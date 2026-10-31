const { test } = require('node:test')
const assert = require('node:assert/strict')
const vm = require('node:vm')
const fs = require('node:fs')
const { EventEmitter } = require('node:events')
const ts = require('typescript')
const path = require('node:path')
const source = ts.transpileModule(fs.readFileSync(path.join(__dirname, '../electron/updates.ts'), 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true },
}).outputText

function setup(packaged = true) {
  const app = Object.assign(new EventEmitter(), { isPackaged: packaged, getPath: () => '/test', getVersion: () => '1.0.6' })
  const updater = new EventEmitter()
  const messages = [], timers = [], cleared = []
  let check, calls = 0, installs = 0, response = 1
  updater.checkForUpdates = async () => { calls++; return { updateInfo: { version: '1.0.6' } } }
  updater.quitAndInstall = () => installs++
  const modules = {
    electron: { app, ipcMain: { handle: (_name, fn) => { check = fn } }, dialog: {
      showMessageBox: async options => { messages.push(options); return { response } },
    } },
    'electron-updater': { autoUpdater: updater },
    'node:fs': { existsSync: () => false, appendFileSync: () => {} },
    'node:path': path,
  }
  const context = { exports: {}, require: name => modules[name], console,
    setTimeout: fn => { timers.push(fn); return 1 }, setInterval: fn => { timers.push(fn); return 2 },
    clearTimeout: id => cleared.push(id), clearInterval: id => cleared.push(id),
  }
  vm.runInNewContext(source, context)
  context.exports.startUpdates()
  return { app, updater, messages, timers, cleared, check: () => check(), calls: () => calls,
    installs: () => installs, answer: value => { response = value } }
}
test('development mode skips network and explains manual check', async () => {
  const s = setup(false)
  await s.check()
  assert.equal(s.calls(), 0)
  assert.equal(s.timers.length, 0)
  assert.match(s.messages[0].message, /kurulu uygulamada/)
})
test('failed check is handled and can be retried; timers clear on quit', async () => {
  const s = setup()
  s.updater.checkForUpdates = async () => { s.updater.emit('error', new Error('404')); throw new Error('404') }
  await s.check()
  assert.match(s.messages[0].message, /indirilemedi/)
  s.updater.checkForUpdates = async () => ({ updateInfo: { version: '1.0.6' } })
  await s.check()
  assert.match(s.messages[1].message, /güncel/)
  s.app.emit('before-quit')
  assert.deepEqual(s.cleared, [1, 2])
})
test('download failures are caught and simultaneous checks are suppressed', async () => {
  const s = setup()
  let rejectDownload, checks = 0
  s.updater.checkForUpdates = async () => {
    checks++
    return { updateInfo: { version: '1.0.7' }, downloadPromise: new Promise((_resolve, reject) => { rejectDownload = reject }) }
  }
  const pending = s.check()
  await s.check()
  assert.equal(checks, 1)
  rejectDownload(new Error('missing installer'))
  await pending
  assert.ok(s.messages.some(m => /indirilemedi/.test(m.message)))
})
test('downloaded update can be deferred then explicitly installed', async () => {
  const s = setup()
  s.updater.emit('update-downloaded', { version: '1.0.7' })
  await new Promise(resolve => setImmediate(resolve))
  assert.equal(s.installs(), 0)
  s.answer(0)
  await s.check()
  assert.equal(s.installs(), 1)
  assert.equal(s.calls(), 0)
  assert.equal(s.updater.autoInstallOnAppQuit, true)
})
