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

function setup() {
  const db = new DatabaseSync(':memory:', { enableDoubleQuotedStringLiterals: true });
  for (const statement of sql) {
    try { db.exec(statement); } catch (error) {
      if (!error.message.includes('duplicate column name')) throw error;
    }
  }
  const call = (name, data) => {
    const scope = { db, data };
    return vm.runInNewContext(compile(`const handler = ${handlers[name]}; handler(null, data);`), scope);
  };
  return { db, call };
}

test('rent increases preserve past payments, deductions, account balances and summary after reload', () => {
  const { db, call } = setup();
  try {
    const property = { name: 'Daire', address: '', status: 'Dolu', rent_amount: 10000,
      tenant: 'Kiracı', tenant_phone: '', tenant_email: '', payment_status: '—', deduction_amount: 0,
      deduction_note: '', payment_due_day: 1, contract_start_date: '', contract_end_date: '', note: '', contract_file_path: '' };
    const saved = call('add-property', property);
    const payment = { property_id: saved.id, month: 'Aralık', year: 2026, payment_date: '2026-12-05',
      status: 'Ödendi', amount_paid: 8500, deduction_amount: 1500, note: 'Tadilat', payment_method: 'Nakit',
      tahir_payi: 0, bekir_payi: 0, tarik_payi: 0, halil_payi: 0, ozcam_payi: 0, ozcam_insaat_payi: 0 };
    call('save-payment', payment);
    const before = call('get-payments', saved.id);
    const cashBefore = call('get-finance-summary');
    for (const rent_amount of [15000, 20000]) {
      call('update-property', { ...saved, rent_amount });
      const summary = call('get-properties')[0];
      assert.equal(summary.rent_amount, rent_amount);
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
