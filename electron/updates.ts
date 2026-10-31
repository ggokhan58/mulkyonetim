import { app, dialog, ipcMain } from 'electron'
import { autoUpdater } from 'electron-updater'
import fs from 'node:fs'
import path from 'node:path'

export function startUpdates(): void {
  const logPath = path.join(app.getPath('userData'), 'updater.log')
  const log = (message: string) => {
    try {
      if (fs.existsSync(logPath) && fs.statSync(logPath).size > 1024 * 1024) {
        fs.renameSync(logPath, `${logPath}.previous`)
      }
      fs.appendFileSync(logPath, `${new Date().toISOString()} ${message}\n`)
    } catch (error) { console.error('Güncelleme günlüğü yazılamadı:', error) }
  }
  const show = async (message: string, detail = '') => {
    await dialog.showMessageBox({ type: 'info', title: 'Mülk Yönetim — Güncelleme', message, detail })
  }
  let busy = false
  let downloadedVersion: string | undefined
  let prompting = false
  const offerInstall = async () => {
    if (prompting || !downloadedVersion) return
    prompting = true
    try {
      const result = await dialog.showMessageBox({
        type: 'info', title: 'Güncelleme hazır',
        message: `${downloadedVersion} sürümü indirildi.`,
        detail: 'Açık formlarınızı kaydettikten sonra yeniden başlatabilirsiniz. Ertelerseniz güncelleme uygulamayı kapattığınızda kurulur.',
        buttons: ['Yeniden Başlat ve Kur', 'Daha Sonra'], defaultId: 1, cancelId: 1,
      })
      if (result.response === 0) autoUpdater.quitAndInstall(false, true)
    } finally { prompting = false }
  }
  const check = async (manual = false) => {
    if (!app.isPackaged) {
      if (manual) await show('Güncelleme kontrolü kurulu uygulamada kullanılabilir.', 'Geliştirme modunda otomatik güncelleme yapılmaz.')
      return
    }
    if (downloadedVersion) {
      if (manual) await offerInstall()
      return
    }
    if (busy) {
      if (manual) await show('Güncelleme kontrolü veya indirme devam ediyor.')
      return
    }
    busy = true
    try {
      log(`Kontrol başladı; mevcut sürüm: ${app.getVersion()}`)
      const result = await autoUpdater.checkForUpdates()
      if (result?.downloadPromise) {
        if (manual) void show(`${result.updateInfo.version} sürümü indiriliyor.`, 'İndirme tamamlandığında kurulum seçeneği gösterilecek.').catch(error => log(String(error)))
        await result.downloadPromise
      } else if (manual) {
        await show(`Uygulamanız güncel (${app.getVersion()}).`)
      }
    } catch (error) {
      log(`Güncelleme başarısız: ${String(error)}`)
      if (manual) await show('Güncelleme kontrol edilemedi veya indirilemedi.', `İnternet bağlantınızı kontrol edip yeniden deneyin. Sorun sürerse günlük dosyası: ${logPath}`)
    } finally { busy = false }
  }
  ipcMain.handle('check-for-updates', () => check(true))
  if (!app.isPackaged) return
  autoUpdater.autoDownload = true
  autoUpdater.autoInstallOnAppQuit = true
  autoUpdater.allowPrerelease = false
  autoUpdater.allowDowngrade = false
  autoUpdater.on('error', error => log(`Updater: ${String(error)}`))
  autoUpdater.on('update-available', info => log(`Yeni sürüm: ${info.version}`))
  autoUpdater.on('update-not-available', () => log('Uygulama güncel'))
  autoUpdater.on('update-downloaded', info => {
    downloadedVersion = info.version
    log(`İndirildi: ${info.version}`)
    void offerInstall().catch(error => log(String(error)))
  })
  const run = () => { void check().catch(error => log(String(error))) }
  const initial = setTimeout(run, 10_000)
  const interval = setInterval(run, 60 * 60 * 1000)
  app.once('before-quit', () => { clearTimeout(initial); clearInterval(interval) })
}
