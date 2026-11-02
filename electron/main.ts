import { updatedRentHistory } from '../src/rent-history';
import { app, shell, BrowserWindow, ipcMain, dialog } from 'electron'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'
import fs from 'node:fs'
import { buildExcelWorkbook } from './excel-export'
import { startUpdates } from './updates';

const currentDir = path.dirname(fileURLToPath(import.meta.url))
const require = createRequire(import.meta.url)
const Database = require('better-sqlite3')

// ─── GÜVENLİ VERİTABANI BAĞLANTISI ───
// 1. userData klasörünün yolunu al
const userDataPath = app.getPath('userData');

// 2. Klasörün var olup olmadığını kontrol et, yoksa güvenli bir şekilde oluştur
if (!fs.existsSync(userDataPath)) {
  fs.mkdirSync(userDataPath, { recursive: true });
}

//PDF İçin Gerekli
const documentsPath = path.join(userDataPath, 'documents');
if (!fs.existsSync(documentsPath)) {
  fs.mkdirSync(documentsPath, { recursive: true });
}

// 3. Veritabanı dosyasının tam yolunu belirle
const dbPath = path.join(userDataPath, 'mulk_yonetim.sqlite');

// 4. Veritabanına bağlan (Dosya yoksa better-sqlite3 kendisi oluşturacaktır)
let db: any;
try {
  db = new Database(dbPath);
} catch (error) {
  console.error("Veritabanı başlatılamadı:", error);
  // Burada isterseniz kullanıcıya bir uyarı mesajı da gösterebilirsiniz.
}

// Tablo Oluşturma
db.exec(`
  CREATE TABLE IF NOT EXISTS properties (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    address TEXT,
    status TEXT,
    rent_amount REAL,
    tenant TEXT,
    tenant_phone TEXT,
    tenant_email TEXT,
    payment_status TEXT,
    deduction_amount REAL,
    deduction_note TEXT,
    payment_due_day INTEGER
  );

  CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    property_id INTEGER,
    month TEXT,          -- Örn: "Ocak", "Şubat"
    year INTEGER,         -- Örn: 2026
    payment_date TEXT,    -- Örn: "2026-06-24" (Ödeme tarihi)
    status TEXT,          -- Örn: "Ödendi", "Gecikti", "Bekliyor"
    amount_paid REAL,     -- Ödenen miktar (kesintiler düşülmüş net miktar için)
    note TEXT,            -- Ödeme ile ilgili özel notlar
    FOREIGN KEY(property_id) REFERENCES properties(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS transfers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_account TEXT,  -- Örn: "Nakit", "Havale (Tahir Çam)"
    amount REAL,          -- Çıkan miktar
    description TEXT,     -- Nereye/Niçin gönderildiği
    transfer_date TEXT    -- İşlem tarihi
  );

`);

// Mevcut tabloların altına bu kodu ekle:
try {
  // Mevcut payments tablosuna deduction_amount sütununu ekler (varsa sessizce atlar)
  db.exec('ALTER TABLE payments ADD COLUMN deduction_amount REAL DEFAULT 0');
} catch (error) {
  // Sütun zaten eklenmişse hata verir, sorun yok.
}

try {
  db.exec('ALTER TABLE properties ADD COLUMN payment_due_day INTEGER DEFAULT 1');
} catch (error) {}
try { db.exec('ALTER TABLE properties ADD COLUMN tenant_phone TEXT'); } catch (e) {}
try { db.exec('ALTER TABLE properties ADD COLUMN tenant_email TEXT'); } catch (e) {}
try { db.exec('ALTER TABLE properties ADD COLUMN contract_start_date TEXT'); } catch (e) {}
try { db.exec('ALTER TABLE properties ADD COLUMN contract_end_date TEXT'); } catch (e) {}
try { db.exec('ALTER TABLE properties ADD COLUMN note TEXT'); } catch (e) {}
try { db.exec('ALTER TABLE properties ADD COLUMN contract_file_path TEXT'); } catch (e) {}
try { db.exec("ALTER TABLE properties ADD COLUMN rent_history TEXT DEFAULT '[]'"); } catch (e) {}
// --- YENİ EKLENEN HAVALE VE PAY SÜTUNLARI ---
try { db.exec("ALTER TABLE payments ADD COLUMN payment_method TEXT DEFAULT 'Nakit'"); } catch (e) {}
try { db.exec('ALTER TABLE payments ADD COLUMN tahir_payi REAL DEFAULT 0'); } catch (e) {}
try { db.exec('ALTER TABLE payments ADD COLUMN bekir_payi REAL DEFAULT 0'); } catch (e) {}
try { db.exec('ALTER TABLE payments ADD COLUMN tarik_payi REAL DEFAULT 0'); } catch (e) {}
try { db.exec('ALTER TABLE payments ADD COLUMN halil_payi REAL DEFAULT 0'); } catch (e) {}

try { db.exec('ALTER TABLE payments ADD COLUMN ozcam_payi REAL DEFAULT 0'); } catch (e) {}
// Mevcut kayıtları koruyarak yeni kasanın pay sütununu ekle.
if (!db.prepare('PRAGMA table_info(payments)').all().some((column: any) => column.name === 'ozcam_insaat_payi')) {
  db.exec('ALTER TABLE payments ADD COLUMN ozcam_insaat_payi REAL DEFAULT 0');
}
try { db.exec('ALTER TABLE transfers ADD COLUMN destination_account TEXT DEFAULT "Dışarıya Harcama / Gider"'); } catch (e) {}

function createWindow(): void {
  // Derlenmiş dosya genelde aynı klasörde yer alır
  const preloadPath = path.join(currentDir, 'preload.cjs'); // Vite standart olarak .mjs veya .js kullanır

  const mainWindow = new BrowserWindow({
    title: `Mülk Yönetim — ${app.getVersion()}`,
    width: 1280,
    height: 800,
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      preload: preloadPath, // Yeni yolu buraya verdik
      sandbox: false
    }
  })

  mainWindow.on('page-title-updated', (event) => {
    event.preventDefault();
    mainWindow.setTitle(`Mülk Yönetim — ${app.getVersion()}`);
  });

  mainWindow.on('ready-to-show', () => {
    mainWindow.show()
    mainWindow.focus()
  })

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url)
    return { action: 'deny' }
  })

  // Vite'in kullandığı geliştirme URL'sini bulma (Hatayı çözen asıl kısım)
  const devUrl = process.env.VITE_DEV_SERVER_URL || process.env.ELECTRON_RENDERER_URL;
  
  if (devUrl) {
    mainWindow.loadURL(devUrl)
  } else {
    mainWindow.loadFile(path.join(currentDir, '../dist/index.html'))
  }
}

app.whenReady().then(() => {
  ipcMain.handle('get-app-version', () => app.getVersion());
  startUpdates();
  
  // ─── ANA TABLOYU ÇEKME SORGUSU (Güncellendi) ───
  ipcMain.handle('get-properties', () => {
    const stmt = db.prepare(`
      WITH LatestPayments AS (
        SELECT property_id, status, note, amount_paid, deduction_amount, month, year,
               ROW_NUMBER() OVER(
                 PARTITION BY property_id 
                 ORDER BY year DESC, 
                 CASE month 
                   WHEN 'Ocak' THEN 1 WHEN 'Şubat' THEN 2 WHEN 'Mart' THEN 3 WHEN 'Nisan' THEN 4 
                   WHEN 'Mayıs' THEN 5 WHEN 'Haziran' THEN 6 WHEN 'Temmuz' THEN 7 WHEN 'Ağustos' THEN 8 
                   WHEN 'Eylül' THEN 9 WHEN 'Ekim' THEN 10 WHEN 'Kasım' THEN 11 WHEN 'Aralık' THEN 12 
                 END DESC
               ) as rn
        FROM payments
      )
      SELECT 
        p.id, p.name, p.address, p.status, p.rent_history, p.rent_amount, p.tenant, p.tenant_phone, p.tenant_email, p.payment_due_day,p.contract_start_date, p.contract_end_date,p.note, p.contract_file_path,
        COALESCE(lp.status, '—') as payment_status,
        COALESCE(lp.note, '') as deduction_note,
        COALESCE(lp.deduction_amount, 0) as deduction_amount,
        lp.amount_paid as last_amount_paid,
        lp.month as last_month,
        lp.year as last_year
      FROM properties p
      LEFT JOIN LatestPayments lp ON p.id = lp.property_id AND lp.rn = 1
    `);
    return stmt.all();
  });

ipcMain.handle('add-property', (_event, property) => {
    const history = updatedRentHistory(null, property);
    try {
      // Sütun sayısı, soru işareti sayısı ve gönderilen veri sayısı (11) eşitlendi.
      const stmt = db.prepare(`
        INSERT INTO properties (
          name, address, status, rent_amount, tenant, tenant_phone, tenant_email, payment_status, deduction_amount, deduction_note, payment_due_day,contract_start_date, contract_end_date,note, contract_file_path, rent_history
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      
      const info = stmt.run(
        property.name, 
        property.address, 
        property.status, 
        property.rent_amount, 
        property.tenant, 
        property.tenant_phone,
        property.tenant_email,
        property.payment_status, 
        property.deduction_amount, 
        property.deduction_note, 
        property.payment_due_day,
        property.contract_start_date, 
        property.contract_end_date,
        property.note,
        property.contract_file_path,
        history
      );
      
      return { ...property, id: info.lastInsertRowid, rent_history: history };
    } catch (error) {
      console.error("Veritabanına kayıt eklenirken hata oluştu:", error);
      // Hatayı fırlatarak React tarafındaki catch bloğunun (alert'in) çalışmasını sağlıyoruz.
      // Böylece uygulama çökmek yerine kullanıcıya uyarı verir.
      throw error; 
    }
  });

  // --- DÜZENLEME YAPMA ---
  ipcMain.handle('update-property', (_event, property) => {
    const previous = db.prepare('SELECT * FROM properties WHERE id = ?').get(property.id);
    if (!previous) throw new Error('Mülk bulunamadı.');
    const history = updatedRentHistory(previous, property);
    const stmt = db.prepare(`
      UPDATE properties 
      SET name = ?, address = ?, status = ?, rent_amount = ?, tenant = ?, tenant_phone = ?, tenant_email = ?, payment_status = ?, deduction_amount = ?, deduction_note = ?, payment_due_day = ?, contract_start_date = ?, contract_end_date = ?, note = ?, contract_file_path = ?, rent_history = ?
      WHERE id = ?
    `);
    
    stmt.run(
      property.name, 
      property.address, 
      property.status, 
      property.rent_amount, 
      property.tenant, 
      property.tenant_phone, 
      property.tenant_email,
      property.payment_status, 
      property.deduction_amount, 
      property.deduction_note, 
      property.payment_due_day, 
      property.contract_start_date, // YENİ
      property.contract_end_date,
      property.note,
      property.contract_file_path,
      history,
      property.id // Hangi satırın güncelleneceğini belirten ID
    );
    
    return property; // Güncellenmiş veriyi arayüze geri dönüyoruz
  });

  // ─── YENİ: ÖDEME GEÇMİŞİNİ GETİRME (Mülk ID'sine göre) ───
  ipcMain.handle('get-payments', (_event, propertyId) => {
    const stmt = db.prepare('SELECT * FROM payments WHERE property_id = ? ORDER BY year DESC, month DESC');
    return stmt.all(propertyId);
  });

  // ─── PDF SEÇME VE KOPYALAMA İŞLEMİ (ÇOKLU SEÇİM İÇİN GÜNCELLENDİ) ───
  ipcMain.handle('select-pdf', async () => {
    const result = await dialog.showOpenDialog({
      // multiSelections parametresi eklendi
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'PDF Dosyaları', extensions: ['pdf'] }]
    });

    if (result.canceled || result.filePaths.length === 0) return null;

    const currentUserDataPath = app.getPath('userData');
    const safeDocumentsPath = path.join(currentUserDataPath, 'documents');

    if (!fs.existsSync(safeDocumentsPath)) {
      fs.mkdirSync(safeDocumentsPath, { recursive: true });
    }

    const copiedPaths: string[] = [];

    // Seçilen tüm dosyalar üzerinde döngü kuruyoruz
    for (const sourcePath of result.filePaths) {
      const fileName = path.basename(sourcePath);
      const uniqueFileName = `${Date.now()}_${fileName}`; 
      const destPath = path.join(safeDocumentsPath, uniqueFileName);

      try {
        fs.copyFileSync(sourcePath, destPath);
        console.log("✅ PDF Kopyalandı! Hedef Yol:", destPath);
        copiedPaths.push(destPath);
      } catch (error) {
        console.error("❌ PDF kopyalanırken hata oluştu:", error);
      }
    }

    // Başarıyla kopyalanan dosyaların yollarını dizi olarak döndürüyoruz
    return copiedPaths.length > 0 ? copiedPaths : null;
  });

  // ─── KAYITLI PDF'İ SİSTEMİN VARSAYILAN UYGULAMASIYLA AÇMA ───
  ipcMain.handle('open-pdf', async (_event, filePath) => {
    if (filePath && fs.existsSync(filePath)) {
      await shell.openPath(filePath);
      return true;
    }
    return false;
  });
  
// ─── ÖDEME KAYDETME/GÜNCELLEME ───
  ipcMain.handle('save-payment', (_event, payment) => {
    if (payment.id) {
      const stmt = db.prepare(`
        UPDATE payments 
        SET month = ?, year = ?, payment_date = ?, status = ?, amount_paid = ?, deduction_amount = ?, note = ?, payment_method = ?, tahir_payi = ?, bekir_payi = ?, tarik_payi = ?, halil_payi = ?, ozcam_payi = ?, ozcam_insaat_payi = ?
        WHERE id = ?
      `);
      stmt.run(payment.month, payment.year, payment.payment_date, payment.status, payment.amount_paid, payment.deduction_amount, payment.note, payment.payment_method, payment.tahir_payi, payment.bekir_payi, payment.tarik_payi, payment.halil_payi, payment.ozcam_payi, payment.ozcam_insaat_payi ?? 0, payment.id);
      return payment;
    } else {
      const stmt = db.prepare(`
        INSERT INTO payments (property_id, month, year, payment_date, status, amount_paid, deduction_amount, note, payment_method, tahir_payi, bekir_payi, tarik_payi, halil_payi, ozcam_payi, ozcam_insaat_payi)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      const info = stmt.run(payment.property_id, payment.month, payment.year, payment.payment_date, payment.status, payment.amount_paid, payment.deduction_amount, payment.note, payment.payment_method, payment.tahir_payi, payment.bekir_payi, payment.tarik_payi, payment.halil_payi, payment.ozcam_payi, payment.ozcam_insaat_payi ?? 0);
      return { id: info.lastInsertRowid, ...payment };
    }
  });

  // ─── YENİ: YANLIŞ GİRİLEN AYI SİLME ───
  ipcMain.handle('delete-payment', (_event, id) => {
    const stmt = db.prepare('DELETE FROM payments WHERE id = ?');
    stmt.run(id);
    return true;
  });


// Tüm dönemlerdeki kayıtları, arayüz filtrelerinden bağımsız olarak dışa aktarır.
ipcMain.handle('export-transactions-excel', async () => {
  let temporaryPath: string | undefined;
  try {
    const result = await dialog.showSaveDialog({
      title: 'Tüm İşlemleri Excel Olarak Kaydet',
      defaultPath: `islemler_${new Date().toISOString().slice(0, 10)}.xlsx`,
      filters: [{ name: 'Excel Çalışma Kitabı', extensions: ['xlsx'] }]
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    // İki liste aynı veritabanı anlık görüntüsünden alınır.
    const { payments, transfers } = db.transaction(() => ({
      payments: db.prepare(`
        SELECT pay.*, COALESCE(p.name, 'Silinmiş mülk') AS property_name,
               COALESCE(p.tenant, '') AS tenant
        FROM payments pay LEFT JOIN properties p ON p.id = pay.property_id
        ORDER BY pay.payment_date, pay.id
      `).all(),
      transfers: db.prepare('SELECT * FROM transfers ORDER BY transfer_date, id').all()
    }))();
    const workbook = buildExcelWorkbook([
      {
        name: 'Ödeme Kayıtları',
        headers: ['Kayıt No', 'Mülk No', 'Mülk', 'Kiracı (Güncel)', 'Ay', 'Yıl', 'Ödeme Tarihi', 'Durum', 'Hesap / Ödeme Yöntemi', 'Kayıtlı Tutar', 'Kesinti', 'Gerçekleşen Tahsilat', 'Not'],
        moneyColumns: [9, 10, 11], dateColumns: [6],
        rows: payments.map((p: any) => [p.id, p.property_id, p.property_name, p.tenant, p.month, p.year,
          p.payment_date, p.status, p.payment_method, p.amount_paid ?? 0, p.deduction_amount ?? 0,
          p.status === 'Ödendi' ? p.amount_paid ?? 0 : 0, p.note])
      },
      {
        name: 'Transfer ve Giderler',
        headers: ['Kayıt No', 'İşlem Tarihi', 'İşlem Türü', 'Çıkış Hesabı', 'Hedef Hesap', 'Tutar', 'Açıklama'],
        moneyColumns: [5], dateColumns: [1],
        rows: transfers.map((t: any) => [t.id, t.transfer_date,
          !t.destination_account || t.destination_account === 'Dışarıya Harcama / Gider' ? 'Gider' : 'Hesaplar Arası Transfer',
          t.source_account, t.destination_account || 'Dışarıya Harcama / Gider', t.amount ?? 0, t.description])
      }
    ]);
    const outputPath = result.filePath.toLowerCase().endsWith('.xlsx') ? result.filePath : result.filePath + '.xlsx';
    temporaryPath = outputPath + `.${Date.now()}.tmp`;
    await fs.promises.writeFile(temporaryPath, workbook, { flag: 'wx' });
    await fs.promises.rename(temporaryPath, outputPath);
    return { success: true, path: outputPath };
  } catch (error) {
    if (temporaryPath) await fs.promises.unlink(temporaryPath).catch(() => {});
    console.error('Excel dışa aktarma hatası:', error);
    return { success: false, error: 'Excel dosyası kaydedilemedi. Dosyanın açık olmadığını ve klasöre yazma izniniz olduğunu kontrol edin.' };
  }
});

// ─── SADECE DIŞA AKTARMA (YEDEKLEME) KANALI ───
ipcMain.handle('export-database', async () => {
  const result = await dialog.showSaveDialog({
    title: 'Veritabanını Dışa Aktar (Yedekle)',
    defaultPath: 'mulk_yonetim_yedek.sqlite',
    filters: [{ name: 'SQLite Veritabanı', extensions: ['sqlite'] }]
  });

  if (!result.canceled && result.filePath) {
    try {
      // Mevcut veritabanı dosyasını seçilen yere kopyala
      fs.copyFileSync(dbPath, result.filePath);
      return { success: true, path: result.filePath };
    } catch (error) {
      console.error("Yedekleme sırasında hata:", error);
      return { success: false };
    }
  }
  return null;
});

  createWindow()
  // ─── YENİ: KAYITLI MÜLKÜ SİLME SORGUSU ───
  ipcMain.handle('delete-property', (_event, id) => {
    const stmt = db.prepare('DELETE FROM properties WHERE id = ?');
    stmt.run(id);
    return true;
  });

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// ─── GÜNCELLENEN: HESAPLAR ARASI TRANSFER DESTEKLİ KASA ÖZETİ ───
  ipcMain.handle('get-finance-summary', () => {
    const stmt = db.prepare(`
      WITH Incomes AS (
        SELECT 
          COALESCE(SUM(tahir_payi), 0) as t_in,
          COALESCE(SUM(bekir_payi), 0) as b_in,
          COALESCE(SUM(tarik_payi), 0) as tr_in,
          COALESCE(SUM(halil_payi), 0) as h_in,
          COALESCE(SUM(ozcam_payi), 0) as oz_in,
          COALESCE(SUM(ozcam_insaat_payi), 0) as oi_in,
          COALESCE(SUM(CASE WHEN payment_method = 'Nakit' THEN amount_paid ELSE 0 END), 0) as n_in
        FROM payments WHERE status = 'Ödendi'
      ),
      Outgoes AS (
        SELECT
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Tahir Çam)' THEN amount ELSE 0 END), 0) as t_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Bekir Çam)' THEN amount ELSE 0 END), 0) as b_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Tarık Çam)' THEN amount ELSE 0 END), 0) as tr_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Halil Çam)' THEN amount ELSE 0 END), 0) as h_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Özçam Yönetim Kurulu)' THEN amount ELSE 0 END), 0) as oz_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Özçam İnşaat)' THEN amount ELSE 0 END), 0) as oi_out,
          COALESCE(SUM(CASE WHEN source_account = 'Nakit' THEN amount ELSE 0 END), 0) as n_out
        FROM transfers
      ),
      TransferIncomes AS (
        SELECT
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Tahir Çam)' THEN amount ELSE 0 END), 0) as t_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Bekir Çam)' THEN amount ELSE 0 END), 0) as b_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Tarık Çam)' THEN amount ELSE 0 END), 0) as tr_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Halil Çam)' THEN amount ELSE 0 END), 0) as h_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Özçam Yönetim Kurulu)' THEN amount ELSE 0 END), 0) as oz_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Özçam İnşaat)' THEN amount ELSE 0 END), 0) as oi_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Nakit' THEN amount ELSE 0 END), 0) as n_ti
        FROM transfers
      )
      SELECT 
        (i.t_in + ti.t_ti - o.t_out) as tahir_total,
        (i.b_in + ti.b_ti - o.b_out) as bekir_total,
        (i.tr_in + ti.tr_ti - o.tr_out) as tarik_total,
        (i.h_in + ti.h_ti - o.h_out) as halil_total,
        (i.oz_in + ti.oz_ti - o.oz_out) as ozcam_total,
        (i.oi_in + ti.oi_ti - o.oi_out) as ozcam_insaat_total,
        (i.n_in + ti.n_ti - o.n_out) as nakit_total
      FROM Incomes i, Outgoes o, TransferIncomes ti;
    `);
    return stmt.get();
  });

  // ─── YENİ: BELİRLİ BİR HESABA/YÖNTEME GÖRE TAHSİLAT DETAYLARINI GETİRME ───
  ipcMain.handle('get-payments-by-method', (_event, method) => {
    const stmt = db.prepare(`
      SELECT 
        p.name as property_name, 
        p.tenant, 
        pay.month, 
        pay.year, 
        pay.payment_date, 
        pay.amount_paid, 
        pay.deduction_amount
      FROM payments pay
      JOIN properties p ON pay.property_id = p.id
      WHERE pay.payment_method = ? AND pay.status = 'Ödendi'
      ORDER BY 
        pay.year DESC, 
        CASE pay.month 
          WHEN 'Ocak' THEN 1 WHEN 'Şubat' THEN 2 WHEN 'Mart' THEN 3 WHEN 'Nisan' THEN 4 
          WHEN 'Mayıs' THEN 5 WHEN 'Haziran' THEN 6 WHEN 'Temmuz' THEN 7 WHEN 'Ağustos' THEN 8 
          WHEN 'Eylül' THEN 9 WHEN 'Ekim' THEN 10 WHEN 'Kasım' THEN 11 WHEN 'Aralık' THEN 12 
        END DESC
    `);
    return stmt.all(method);
  });

  // ─── PARA ÇIKIŞI / TRANSFER EKLEME ───
ipcMain.handle('add-transfer', (_event, transferData) => {
    const stmt = db.prepare(`
      INSERT INTO transfers (source_account, destination_account, amount, description, transfer_date)
      VALUES (?, ?, ?, ?, ?)
    `);
    const info = stmt.run(transferData.source_account, transferData.destination_account, transferData.amount, transferData.description, transferData.transfer_date);
    return { id: info.lastInsertRowid, ...transferData };
  });

  // ─── TRANSFER / PARA ÇIKIŞI GEÇMİŞİNİ GETİRME ───
  ipcMain.handle('get-transfers', () => {
    const stmt = db.prepare('SELECT * FROM transfers ORDER BY transfer_date DESC, id DESC');
    return stmt.all();
  });

  // ─── TRANSFER SİLME (Geri Alma) ───
  ipcMain.handle('delete-transfer', (_event, id) => {
    const stmt = db.prepare('DELETE FROM transfers WHERE id = ?');
    stmt.run(id);
    return true;
  });

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
