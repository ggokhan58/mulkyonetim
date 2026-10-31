// Run with Electron after building; all data stays in release/ui-smoke.
const { app, BrowserWindow, dialog } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const assert = require('node:assert/strict');
const output = path.resolve(__dirname, '../release/ui-smoke');
fs.mkdirSync(output, { recursive: true });
app.setPath('userData', output);
app.disableHardwareAcceleration();
BrowserWindow.prototype.show = function () {};
BrowserWindow.prototype.focus = function () {};
const messages = [];
dialog.showMessageBox = async options => { messages.push(options); return { response: 1 }; };
const timeout = setTimeout(() => { console.error('UI smoke timeout'); app.exit(1); }, 30000);
app.on('browser-window-created', (_event, window) => {
  window.webContents.setBackgroundThrottling(false);
  window.webContents.once('did-finish-load', async () => {
    try {
      const evaluate = code => window.webContents.executeJavaScript(code);
      for (let i = 0; i < 100; i++) {
        if (await evaluate(`!!document.querySelector('aside button') && document.querySelector('aside').textContent.includes('v' + ${JSON.stringify(app.getVersion())})`)) break;
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      assert.equal(window.getTitle(), `Mülk Yönetim — ${app.getVersion()}`);
      assert.equal(await evaluate(`document.documentElement.lang`), 'tr');
      assert.ok(await evaluate(`document.querySelector('link[rel="icon"]').href.endsWith('/icon.ico')`));
      assert.ok(await evaluate(`document.querySelector('aside').textContent.includes('Güncellemeleri Kontrol Et')`));
      await evaluate(`[...document.querySelectorAll('nav button')].find(b => b.textContent.trim() === 'Ayarlar').click()`);
      await new Promise(resolve => setTimeout(resolve, 200));
      assert.ok(await evaluate(`document.querySelector('main').textContent.includes('Kurulu sürüm: ' + ${JSON.stringify(app.getVersion())})`));
      await evaluate(`[...document.querySelectorAll('main button')].find(b => b.textContent.includes('Güncellemeleri Kontrol Et')).click()`);
      for (let i = 0; i < 30 && !messages.length; i++) await new Promise(resolve => setTimeout(resolve, 100));
      assert.ok(messages.some(m => m.message.includes('kurulu uygulamada')));
      await new Promise(resolve => setTimeout(resolve, 1500));
      assert.equal(await evaluate(`[...document.querySelectorAll('aside button')].find(b => b.textContent.includes('Güncellemeleri Kontrol Et')).disabled`), false);
      fs.writeFileSync(path.join(output, 'settings.png'), (await window.webContents.capturePage()).toPNG());
      console.log('PASS: title, icon, version, sidebar/settings buttons, preload IPC and update dialog');
      clearTimeout(timeout);
      app.exit(0);
    } catch (error) { console.error(error); app.exit(1); }
  });
});
import(pathToFileURL(path.resolve(__dirname, '../dist-electron/main.js')).href).catch(error => {
  console.error(error); app.exit(1);
});
