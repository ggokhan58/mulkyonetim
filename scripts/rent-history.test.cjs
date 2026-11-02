const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { DatabaseSync } = require('node:sqlite');

const main = fs.readFileSync(path.join(__dirname, '../electron/main.ts'), 'utf8');
const tree = ts.createSourceFile('main.ts', main, ts.ScriptTarget.Latest, true);
const sql = [], handlers = {};
function visit(node) {
  if (ts.isCallExpression(node)) {
    const name = node.expression.getText(tree);
    if (name === 'db.exec' && ts.isStringLiteralLike(node.arguments[0])) sql.push(node.arguments[0].text);
    if (name === 'ipcMain.handle') handlers[node.arguments[0].text] = node.arguments[1].getText(tree);
  }
  ts.forEachChild(node, visit);
}
visit(tree);
function compile(source) {
  return ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
}
const context = { exports: {} };
vm.runInNewContext(compile(fs.readFileSync(path.join(__dirname, '../src/payment-values.ts'), 'utf8')), context);
const { collectedAmount, initialPaymentAmount } = context.exports;
const rentContext = { exports: {} };
vm.runInNewContext(compile(fs.readFileSync(path.join(__dirname, '../src/rent-history.ts'), 'utf8')), rentContext);
const { updatedRentHistory, periodRent, readRentHistory } = rentContext.exports;

function setup() {
  const db = new DatabaseSync(':memory:', { enableDoubleQuotedStringLiterals: true });
  for (const statement of sql) {
    try { db.exec(statement); } catch (error) {
      if (!error.message.includes('duplicate column name')) throw error;
    }
  }
  const call = (name, data) => {
    const scope = { db, data, updatedRentHistory };
    return vm.runInNewContext(compile(`const handler = ${handlers[name]}; handler(null, data);`), scope);
  };
  return { db, call };
}

test('rent increases preserve past payments, deductions, account balances and summary after reload', () => {
  const { db, call } = setup();
  try {
    const property = { name: 'Daire', address: '', status: 'Dolu', rent_amount: 10000,
      tenant: 'Kiracı', tenant_phone: '', tenant_email: '', payment_status: '—', deduction_amount: 0,
      deduction_note: '', payment_due_day: 1, contract_start_date: '2026-01-01', contract_end_date: '', note: '', contract_file_path: '' };
    const saved = call('add-property', property);
    const payment = { property_id: saved.id, month: 'Aralık', year: 2026, payment_date: '2026-12-05',
      status: 'Ödendi', amount_paid: 8500, deduction_amount: 1500, note: 'Tadilat', payment_method: 'Nakit',
      tahir_payi: 0, bekir_payi: 0, tarik_payi: 0, halil_payi: 0, ozcam_payi: 0, ozcam_insaat_payi: 0 };
    call('save-payment', payment);
    const before = call('get-payments', saved.id);
    const cashBefore = call('get-finance-summary');
    for (const rent_amount of [15000, 20000]) {
      call('update-property', { ...saved, rent_amount, rent_effective_date: rent_amount === 15000 ? '2027-01-01' : '2027-02-01' });
      const summary = call('get-properties')[0];
      assert.equal(summary.rent_amount, rent_amount);
      const history = readRentHistory(summary.rent_history);
      assert.equal(periodRent(history, 2026, 11, 1), 10000);
      assert.equal(periodRent(history, 2027, 0, 1), 15000);
      if (rent_amount === 20000) assert.equal(periodRent(history, 2027, 1, 1), 20000);
      assert.throws(() => call('update-property', { ...saved, rent_amount: 30000, rent_effective_date: '2025-01-01' }), /sonra/);
      assert.equal(call('get-properties')[0].rent_amount, rent_amount);
      assert.equal(collectedAmount(summary), 8500);
      assert.equal(summary.deduction_amount, 1500);
      assert.deepEqual(call('get-payments', saved.id), before);
      assert.deepEqual(call('get-finance-summary'), cashBefore);
    }
    call('save-payment', { ...payment, month: 'Ocak', year: 2027, amount_paid: 20000, deduction_amount: 0 });
    assert.equal(collectedAmount(call('get-properties')[0]), 20000);
    assert.equal(call('get-payments', saved.id).find(p => p.year === 2026).amount_paid, 8500);
    assert.equal(call('get-finance-summary').nakit_total, 28500);
  } finally { db.close(); }
});

test('new property and unpaid or zero-valued records do not become collected rent', () => {
  assert.equal(collectedAmount({ payment_status: '—' }), 0);
  assert.equal(collectedAmount({ payment_status: 'Bekliyor', last_amount_paid: 10000 }), 0);
  assert.equal(collectedAmount({ payment_status: 'Gecikti', last_amount_paid: 10000 }), 0);
  assert.equal(collectedAmount({ payment_status: 'Ödendi', last_amount_paid: 0 }), 0);
});

test('historical entries keep their amounts; missing past periods require explicit entry', () => {
  const today = new Date(2027, 0, 15);
  assert.equal(initialPaymentAmount({ amount_paid: 8500 }, 20000, 2026, 11, today), '8500');
  assert.equal(initialPaymentAmount({ amount_paid: 0 }, 20000, 2026, 11, today), '0');
  assert.equal(initialPaymentAmount(undefined, 20000, 2026, 11, today), '');
  assert.equal(initialPaymentAmount(undefined, 20000, 2027, 0, today), '20000');
  assert.equal(initialPaymentAmount(undefined, 20000, 2027, 1, today), '20000');
});

test('dated agreements retain each period, boundaries, future rents and unknown history', () => {
  let previous = { rent_amount: 10000, rent_history: JSON.stringify([{ effective_date: '2026-01-01', amount: 10000 }]) };
  let history = updatedRentHistory(previous, { rent_amount: 15000, rent_effective_date: '2026-06-15' });
  previous = { ...previous, rent_amount: 15000, rent_history: history };
  history = updatedRentHistory(previous, { rent_amount: 20000, rent_effective_date: '2027-01-01' });
  const periods = readRentHistory(history);
  assert.equal(periodRent(periods, 2025, 11, 1), undefined);
  assert.equal(periodRent(periods, 2026, 5, 14), 10000);
  assert.equal(periodRent(periods, 2026, 5, 15), 15000);
  assert.equal(periodRent(periods, 2026, 11, 31), 15000);
  assert.equal(periodRent(periods, 2027, 0, 1), 20000);
  assert.equal(periodRent(periods, 2028, 1, 31), 20000);
  assert.throws(() => updatedRentHistory(previous, { rent_amount: 25000 }), /tarihi/);
  assert.throws(() => updatedRentHistory(previous, { rent_amount: 25000, rent_effective_date: '2026-06-15' }), /sonra/);
  assert.throws(() => updatedRentHistory(previous, { rent_amount: 25000, rent_effective_date: '2027-02-30' }), /tarihi/);
  assert.equal(updatedRentHistory(previous, { rent_amount: 15000 }), previous.rent_history);
  const unknown = readRentHistory(updatedRentHistory({ rent_amount: 10000 }, { rent_amount: 20000, rent_effective_date: '2026-11-01' }));
  assert.equal(periodRent(unknown, 2026, 9, 1), undefined);
});
