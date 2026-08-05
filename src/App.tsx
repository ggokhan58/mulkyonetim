import React, { useState, useEffect } from 'react';

// TypeScript için veritabanı yapımızın tip tanımı
interface Property {
  id?: number;
  name: string;
  address: string;
  status: string;
  rent_amount: string | number;
  tenant: string;
  tenant_phone?: string; 
  tenant_email?: string;
  payment_status: string;
  deduction_amount: string | number;
  deduction_note: string;
  payment_due_day: string | number;
  last_month?: string;
  last_year?: number;
  contract_start_date?: string; 
  contract_end_date?: string;   
  note?: string;
  contract_file_path?: string;
}

function App() {
  const [properties, setProperties] = useState<Property[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('dashboard');
  
  // Modal State'leri
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [formData, setFormData] = useState<Property>({
      name: '', address: '', tenant: '',tenant_phone: '', tenant_email: '', rent_amount: '', status: 'Boş', payment_status: '—', deduction_amount: '', deduction_note: '', payment_due_day: 1,contract_start_date: '',contract_end_date: '', note: '',contract_file_path: ''
    });

  // ─── ÖDEME GEÇMİŞİ MODAL STATE'LERİ ───
  const [isPaymentModalOpen, setIsPaymentModalOpen] = useState(false);
  const [selectedProperty, setSelectedProperty] = useState<Property | null>(null);
  const [propertyPayments, setPropertyPayments] = useState<any[]>([]);
  // YENİ: Hangi yılın görüntülendiğini tutan state (Varsayılan olarak bulunduğumuz yıl)
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());

  // Ödeme Ekleme/Düzenleme Form State'i
  const [isPaymentFormOpen, setIsPaymentFormOpen] = useState(false);
  const [paymentFormData, setPaymentFormData] = useState({
    id: undefined,
    month: 'Ocak',
    year: 2026,
    status: 'Bekliyor',
    payment_method: 'Nakit',
    amount_paid: '',
    deduction_amount: '', // YENİ
    payment_date: new Date().toISOString().split('T')[0], // Bugünün tarihi varsayılan
    note: ''
  });

  // ─── PARA ÇIKIŞI / TRANSFER MODAL STATE'LERİ ───
  const [isTransferFormOpen, setIsTransferFormOpen] = useState(false);
  const [isTransferHistoryOpen, setIsTransferHistoryOpen] = useState(false);
  const [transferHistoryList, setTransferHistoryList] = useState<any[]>([]);
  const [transferFormData, setTransferFormData] = useState({
    source_account: 'Nakit', amount: '', description: '', transfer_date: new Date().toISOString().split('T')[0]
  });

  const fetchTransferHistory = async () => {
    try {
      // @ts-ignore
      const data = await window.api.getTransfers();
      setTransferHistoryList(data || []);
    } catch (error) { console.error(error); }
  };

  const handleTransferSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      const payload = {
        ...transferFormData,
        amount: Number(transferFormData.amount) || 0
      };
      // @ts-ignore
      await window.api.addTransfer(payload);
      setIsTransferFormOpen(false);
      setTransferFormData({ source_account: 'Nakit', amount: '', description: '', transfer_date: new Date().toISOString().split('T')[0] });
      fetchFinanceSummary(); // Kasa toplamlarını güncelle (Düşülmüş halini ekrana yansıtır)
    } catch (error) { console.error("Transfer kaydedilemedi:", error); }
  };

  const handleDeleteTransfer = async (id: number) => {
    if (!window.confirm("Bu para çıkışını (transferi) silmek istediğinize emin misiniz? Kasa bakiyesi tekrar artacaktır.")) return;
    try {
      // @ts-ignore
      await window.api.deleteTransfer(id);
      fetchTransferHistory();
      fetchFinanceSummary();
    } catch (error) { console.error(error); }
  };

  const [financeSummary, setFinanceSummary] = useState({
    tahir_total: 0, bekir_total: 0, tarik_total: 0, halil_total: 0, nakit_total: 0
  });

  // Türkçe Ay Listesi (Grid oluşturmak için)
  const monthsList = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];

  // ─── SATIRA ÇİFT TIKLANDIĞINDA ÇALIŞACAK FONKSİYON ───
  const handlePropertyDoubleClick = async (property: Property) => {
    if (!property.id) return;
    setSelectedProperty(property);
    try {
      // @ts-ignore
      const data = await window.api.fetchPayments(property.id);
      setPropertyPayments(data);
      setIsPaymentModalOpen(true);
    } catch (error) {
      console.error("Ödeme geçmişi yüklenirken hata oluştu:", error);
    }
  };

// ─── HESAP DETAY MODALI STATE'LERİ ───
  const [isAccountModalOpen, setIsAccountModalOpen] = useState(false);
  const [selectedAccountName, setSelectedAccountName] = useState('');
  const [accountDetails, setAccountDetails] = useState<any[]>([]);

  // Hangi karta tıklandığını yakalayacak fonksiyon
  const handleAccountClick = async (accountTitle: string, methodString: string) => {
    setSelectedAccountName(accountTitle);
    setIsAccountModalOpen(true);
    setAccountDetails([]); // Yüklenirken tabloyu temizle
    try {
      // @ts-ignore
      // NOT: Backend (Electron) tarafında 'getPaymentsByMethod' adında bir fonksiyon yazacağız
      const data = await window.api.getPaymentsByMethod(methodString);
      setAccountDetails(data || []);
    } catch (error) {
      console.error("Hesap detayları çekilemedi:", error);
    }
  };

// Yeni Bir Ay Ödemesi Kaydetme veya Güncelleme
  const handlePaymentSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedProperty) return;

    try {
      const amountPaid = Number(paymentFormData.amount_paid) || 0;
      const method = paymentFormData.payment_method; // Hangi hesabın seçildiği

      const formattedPayment = {
        ...paymentFormData,
        property_id: selectedProperty.id,
        amount_paid: amountPaid,
        deduction_amount: Number(paymentFormData.deduction_amount) || 0,
        year: selectedYear,
        
        // YENİ MANTIK: Sadece paranın yattığı hesaba tutarı yaz, diğerleri 0 kalsın
        tahir_payi: method === 'Havale (Tahir Çam)' ? amountPaid : 0,
        bekir_payi: method === 'Havale (Bekir Çam)' ? amountPaid : 0,
        tarik_payi: method === 'Havale (Tarık Çam)' ? amountPaid : 0,
        halil_payi: method === 'Havale (Halil Çam)' ? amountPaid : 0
      };

      // @ts-ignore
      await window.api.savePayment(formattedPayment);
      
      // Listeyi yenile
      // @ts-ignore
      const updatedPayments = await window.api.fetchPayments(selectedProperty.id);
      setPropertyPayments(updatedPayments);
      setIsPaymentFormOpen(false);
      
      fetchProperties(); 
      fetchFinanceSummary();
    } catch (error) {
      console.error("Ödeme kaydedilemedi:", error);
    }
  };

  const fetchFinanceSummary = async () => {
    try {
      // @ts-ignore
      const data = await window.api.getFinanceSummary();
      if (data) setFinanceSummary(data);
    } catch (error) {
      console.error("Finans özeti çekilemedi:", error);
    }
  };

  // ─── OTOMATİK GECİKME KONTROLÜ ───
  const isOverdue = (prop: Property) => {
    if (prop.status === 'Boş') return false; 
    
    const currentDay = new Date().getDate();
    const currentMonthIndex = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    const currentMonthName = monthsList[currentMonthIndex];

    // Eğer son ödeme "bu ay" ve "bu yıla" aitse ve durumu 'Ödendi' ise gecikme yoktur.
    const isPaidThisMonth = prop.last_month === currentMonthName && prop.last_year === currentYear && prop.payment_status === 'Ödendi';

    // Bu ay ödenmediyse VE bugünün tarihi ödeme gününü geçtiyse -> Gecikmiştir
    if (!isPaidThisMonth && currentDay > Number(prop.payment_due_day)) {
      return true;
    }
    return false;
  };

  const handleDeletePayment = async () => {
    if (!paymentFormData.id || !selectedProperty) return;
    const confirmDelete = window.confirm(`${paymentFormData.month} ayına ait kaydı silmek istediğinize emin misiniz?`);
    if (!confirmDelete) return;
  
    try {
      // @ts-ignore
      await window.api.deletePayment(paymentFormData.id);
      
      // @ts-ignore
      const updatedPayments = await window.api.fetchPayments(selectedProperty.id);
      setPropertyPayments(updatedPayments);
      setIsPaymentFormOpen(false);
      fetchProperties(); // Ana tabloyu da tazele ki Aralık silinince Haziran geri gelsin
      fetchFinanceSummary();
    } catch (error) {
      console.error("Ödeme silinemedi:", error);
    }
  };

  // ─── MÜLK SİLME FONKSİYONU ───
  const handleDeleteProperty = async (id: number, name: string) => {
    const confirmDelete = window.confirm(
      `"${name}" isimli mülkü ve bu mülke ait TÜM geçmiş ödeme kayıtlarını silmek istediğinize emin misiniz?\nBu işlem geri alınamaz!`
    );
    if (!confirmDelete) return;

    try {
      // @ts-ignore
      await window.api.deleteProperty(id);
      
      // Listeyi tazelemek için mevcut mülkleri yeniden çekiyoruz
      fetchProperties();
      
      // Eğer mülk düzenleme modalı açıksa kapatıyoruz
      setIsModalOpen(false);
    } catch (error) {
      console.error("Mülk silinirken hata oluştu:", error);
      alert("Mülk silme işlemi başarısız oldu.");
    }
  };

  // Belirli bir ayın veritabanı kaydını bulma yardımcısı
  const getPaymentForMonth = (monthName: string) => {
    return propertyPayments.find(p => p.month === monthName && p.year === selectedYear);
  };

  // ─── ELECTRON ÜZERİNDEN SQLITE'TAN VERİLERİ ÇEK (GET) ───
  useEffect(() => {
    fetchProperties();
    fetchFinanceSummary();
  }, []);

  const fetchProperties = async () => {
    try {
      // json-server yerine artık doğrudan kendi masaüstü API'mize bağlanıyoruz!
      // @ts-ignore (TypeScript'in window.api'yi tanıması için hızlı çözüm)
      const data = await window.api.fetchProperties();
      setProperties(data);
    } catch (error) {
      console.error("Veri çekilirken hata oluştu:", error);
    } finally {
      setIsLoading(false);
    }
  };

  // ─── DİNAMİK KPI HESAPLAMALARI ───
  const totalProperties = properties.length;
  const occupiedProperties = properties.filter(p => p.status !== 'Boş').length;
  const occupancyRate = totalProperties > 0 ? Math.round((occupiedProperties / totalProperties) * 100) : 0;
  
  const expectedMonthlyRevenue = properties.reduce((acc, curr) => {
    const rent = Number(curr.rent_amount) || 0;
    const deduction = Number(curr.deduction_amount) || 0;
    if (curr.status !== 'Boş') {
      return acc + (rent - deduction);
    }
    return acc;
  }, 0);

  const delayedPayments = properties.filter(p => p.payment_status === 'Gecikti').length;

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement| HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

// --- EKLENECEK YARDIMCI FONKSİYONLAR ---
  
  // Geriye dönük uyumlu yol çözümleyici (Eski tekli kayıtları ve yeni çoklu kayıtları okur)
  const parsePaths = (pathString?: string): string[] => {
    if (!pathString) return [];
    try {
      const parsed = JSON.parse(pathString);
      return Array.isArray(parsed) ? parsed : [pathString];
    } catch (error) {
      // JSON formatında değilse eski formattaki tek bir dosya yoludur
      return [pathString];
    }
  };

// PDF Ekleme
  const handlePdfSelect = async () => {
    try {
      // @ts-ignore
      const filePaths = await window.api.selectPdf();
      if (filePaths && filePaths.length > 0) {
        // Mevcut seçili yolları al
        const existingPaths = parsePaths(formData.contract_file_path);
        // Yeni seçilenleri üzerine ekle
        const updatedPaths = [...existingPaths, ...filePaths];
        const newPathsString = JSON.stringify(updatedPaths);
        
        // 1. Arayüzü anında güncelle
        setFormData(prev => ({ ...prev, contract_file_path: newPathsString }));

        // 2. Eğer mevcut bir kayıt düzenleniyorsa, anında veritabanına kaydet (Auto-Save)
        if (formData.id) {
          const formattedData = {
            ...formData,
            rent_amount: Number(formData.rent_amount) || 0,
            deduction_amount: Number(formData.deduction_amount) || 0,
            payment_due_day: Number(formData.payment_due_day) || 1,
            contract_file_path: newPathsString // Yeni dosya yolunu ekle
          };
          
          // @ts-ignore
          await window.api.updateProperty(formattedData);
          fetchProperties(); // Arka planda ana tabloyu da tazele
        }
      }
    } catch (error) {
      console.error("PDF seçilemedi:", error);
    }
  };

// Listeden PDF Silme
  const removePdf = async (indexToRemove: number) => {
    const existingPaths = parsePaths(formData.contract_file_path);
    const updatedPaths = existingPaths.filter((_, idx) => idx !== indexToRemove);
    const newPathsString = JSON.stringify(updatedPaths);

    // 1. Arayüzü güncelle
    setFormData(prev => ({ ...prev, contract_file_path: newPathsString }));

    // 2. Veritabanını anında güncelle
    if (formData.id) {
       const formattedData = {
            ...formData,
            rent_amount: Number(formData.rent_amount) || 0,
            deduction_amount: Number(formData.deduction_amount) || 0,
            payment_due_day: Number(formData.payment_due_day) || 1,
            contract_file_path: newPathsString
       };
       // @ts-ignore
       await window.api.updateProperty(formattedData);
       fetchProperties();
    }
  };

  // Formu sıfırlamak için yardımcı fonksiyon
  const resetForm = () => {
      setFormData({ 
        name: '', address: '', tenant: '',tenant_phone: '', tenant_email: '', rent_amount: '', status: 'Boş', payment_status: '—', deduction_amount: '', deduction_note: '', payment_due_day: 1,contract_start_date: '',contract_end_date: '',note: '', contract_file_path: '' 
      });
    };

  // Yeni Kayıt Butonuna Tıklanınca
  const handleNewRecordClick = () => {
    resetForm(); // Önce formu temizle (önceden düzenleme yapılmış olabilir)
    setIsModalOpen(true);
  };

  // Düzenle Butonuna Tıklanınca
  const handleEditClick = (property: Property) => {
    setFormData(property); // Seçili mülkün verilerini forma doldur
    setIsModalOpen(true);  // Modalı aç
  };

  // ─── ELECTRON ÜZERİNDEN SQLITE'A KAYIT EKLE VEYA GÜNCELLE ───
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    try {
      const formattedData = {
        ...formData,
        rent_amount: Number(formData.rent_amount) || 0,
        deduction_amount: Number(formData.deduction_amount) || 0,
        payment_due_day: Number(formData.payment_due_day) || 1
      };

      if (formData.id) {
        // ID VARSA: GÜNCELLEME İŞLEMİ
        // @ts-ignore
        const updatedProperty = await window.api.updateProperty(formattedData);
        
        // Tablodaki eski veriyi yeni veri ile değiştiriyoruz
        setProperties(prev => prev.map(p => p.id === formData.id ? updatedProperty : p));
      } else {
        // ID YOKSA: YENİ KAYIT İŞLEMİ (Eski Kodun)
        // @ts-ignore
        const savedProperty = await window.api.saveProperty(formattedData);
        setProperties(prev => [...prev, savedProperty]);
      }
      
      // Arayüzü temizle ve modalı kapat
      setIsModalOpen(false);
      resetForm();
      
    } catch (error) {
      console.error("Kayıt işlemi sırasında hata oluştu:", error);
      alert("İşlem gerçekleştirilemedi.");
    }
  };

  // Arayüz
  return (
    <div className="flex min-h-screen bg-[#F4F7F9] font-sans text-gray-900 antialiased relative">

      {/* ─── SİYAH & MAVİ KONSEPTLİ MODERN SOL MENÜ ─── */}
      <aside className="w-64 bg-[#09090B] text-gray-400 flex flex-col justify-between shrink-0 transition-all border-r border-gray-800">
        <div>
          <div className="flex items-center gap-3 px-6 py-8 mb-2">
            <div className="bg-blue-600 p-2 rounded-xl text-white shadow-lg shadow-blue-500/20">
              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" /></svg>
            </div>
            <div>
              <h2 className="text-white font-extrabold tracking-wide text-lg">MÜLK<span className="text-blue-500">YÖNETİM</span></h2>
            </div>
          </div>

          <nav className="space-y-1.5 px-4 mt-4">
            {[
              { id: 'dashboard', label: 'Genel Bakış', icon: 'M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h3m-6 0a1 1 0 001-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 001 1m-6 0h6' },
              { id: 'properties', label: 'Mülk Envanteri', icon: 'M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4' },
              { id: 'payments', label: 'Finans / Kira', icon: 'M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z' },
              { id: 'tenants', label: 'Kiracılar', icon: 'M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z' },
              { id: 'settings', label: 'Ayarlar', icon: 'M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z' }
            ].map((item) => (
              <button
                key={item.id}
                onClick={() => setActiveTab(item.id)}
                className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all duration-300 ${
                  activeTab === item.id
                    ? 'bg-blue-600 text-white shadow-md shadow-blue-900/20'
                    : 'hover:bg-[#18181B] hover:text-white'
                }`}
              >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d={item.icon} />
                </svg>
                {item.label}
              </button>
            ))}
          </nav>
        </div>

        <div className="p-4 m-4 bg-[#18181B] rounded-xl flex items-center gap-3 border border-white/5">
          <div className="w-10 h-10 rounded-full bg-blue-600 flex items-center justify-center font-bold text-white shadow-inner">
            CG
          </div>
          <div className="overflow-hidden">
            <h4 className="text-sm font-bold text-white truncate">Çam Grup</h4>
            <p className="text-[11px] text-gray-400 truncate mt-0.5">Sistem Yöneticisi</p>
          </div>
        </div>
      </aside>

      {/* ─── ANA İÇERİK ─── */}
      <main className="flex-1 flex flex-col min-w-0 h-screen overflow-hidden">
        
        {/* ÜST BAR */}
        <header className="h-20 bg-white/80 backdrop-blur-md border-b border-gray-200/60 px-10 flex items-center justify-between shrink-0 z-10 sticky top-0">
          <div className="w-96 relative group">
            <input
              type="text"
              placeholder="Adres, kiracı veya mülk ara..."
              className="w-full bg-gray-100/50 border border-transparent rounded-xl pl-11 pr-4 py-2.5 text-sm focus:outline-none focus:bg-white focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm"
            />
            <svg className="w-4 h-4 text-gray-400 absolute left-4 top-3.5 group-focus-within:text-blue-500 transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>
          </div>

          <div className="flex items-center gap-4">
            <button 
              onClick={handleNewRecordClick}
              className="bg-[#09090B] hover:bg-gray-800 text-white font-medium text-sm px-5 py-2.5 rounded-xl transition-all flex items-center gap-2 shadow-lg shadow-black/5"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4v16m8-8H4" /></svg>
              Yeni Kayıt
            </button>
          </div>
        </header>

        {/* İÇERİK */}
        <div className="p-10 max-w-[1400px] w-full mx-auto overflow-y-auto pb-20">
          
          {/* ─── 1. GENEL BAKIŞ (DASHBOARD) SEKMESİ ─── */}
          {activeTab === 'dashboard' && (
            <div className="animate-fade-in">
              <div className="mb-8">
                <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Mülk Yönetim Paneli</h1>
                <p className="text-gray-500 text-sm mt-1">Veritabanından çekilen canlı finansal durum özeti.</p>
              </div>

              {/* KPI Kartları */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
                {[
                  { title: "Toplam Kayıt", val: totalProperties.toString(), desc: "Aktif mülk portföyü", trendColor: "text-emerald-500" },
                  { title: "Doluluk Oranı", val: `%${occupancyRate}`, desc: `${occupiedProperties} Dolu / ${totalProperties - occupiedProperties} Boş`, trendColor: "text-gray-400" },
                  { title: "Aylık Net Hacim", val: `₺${expectedMonthlyRevenue.toLocaleString('tr-TR')}`, desc: "Kesintiler sonrası beklenen", trendColor: "text-emerald-500" },
                  { title: "Tahsilat Durumu", val: `${delayedPayments} Gecikme`, desc: delayedPayments > 0 ? "İşlem bekleyen" : "Tümü ödendi", trendColor: delayedPayments > 0 ? "text-red-500" : "text-emerald-500" },
                ].map((card, i) => (
                  <div key={i} className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow group relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-transparent via-blue-500/0 to-transparent group-hover:via-blue-500/50 transition-all duration-500"></div>
                    <div className="text-sm font-semibold text-gray-400 mb-2">{card.title}</div>
                    <div className="text-3xl font-black text-gray-900 mb-2">{card.val}</div>
                    <div className="flex items-center justify-between mt-4">
                      <span className="text-xs text-gray-500 font-medium">{card.desc}</span>
                    </div>
                  </div>
                ))}
              </div>

              {/* Mevcut Tablo Görünümü */}
              <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
                <div className="px-6 py-5 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
                  <h3 className="font-bold text-gray-900">Kayıtlı Mülkler (Finansal Görünüm)</h3>
                </div>
                
                {/* Orijinal Tablo Yapısı */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-white border-b border-gray-100 text-gray-400 text-xs font-semibold uppercase tracking-wider">
                        <th className="py-4 px-6">Mülk Detayı</th>
                        <th className="py-4 px-6">Kiracı Sorumlusu</th>
                        <th className="py-4 px-6 text-center">Durum</th>
                        <th className="py-4 px-6 text-center">Tahsilat</th>
                        <th className="py-4 px-6 w-1/5">Tadilat / Kesinti</th>
                        <th className="py-4 px-6 text-right">Kira Bedeli</th>
                        <th className="py-4 px-6 text-right">İşlemler</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50 text-sm">
                      {isLoading ? (
                        <tr>
                          <td colSpan={7} className="py-10 text-center text-gray-500 font-medium">Veriler yükleniyor...</td>
                        </tr>
                      ) : properties.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="py-10 text-center text-gray-500 font-medium">Henüz kayıtlı mülk bulunmuyor.</td>
                        </tr>
                      ) : properties.map((prop) => (
                        <tr 
                            key={prop.id} 
                            onDoubleClick={() => handlePropertyDoubleClick(prop)}
                            className="hover:bg-blue-50/30 transition-colors group cursor-pointer select-none"
                            title="Ödeme geçmişini görmek için çift tıklayın"
                          >
                          <td className="py-4 px-6">
                            <div className="font-bold text-gray-900">{prop.name}</div>
                            <div className="text-xs text-gray-500 mt-1 flex items-center gap-1.5">
                              <svg className="w-3.5 h-3.5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" /></svg>
                              {prop.address}
                            </div>
                          </td>
                          <td className="py-4 px-6 font-medium text-gray-700">
                            {prop.tenant && prop.tenant !== "—" ? (
                              <div className="flex items-center gap-3">
                                <div className="w-7 h-7 rounded-full bg-gray-100 flex items-center justify-center text-xs font-bold text-gray-600 uppercase">
                                  {prop.tenant.substring(0,2)}
                                </div>
                                {prop.tenant}
                              </div>
                            ) : (
                              <span className="text-gray-400">—</span>
                            )}
                          </td>
                          <td className="py-4 px-6 text-center">
                            <span className={`inline-flex items-center justify-center px-2.5 py-1 rounded-md text-xs font-bold ${
                              prop.status === 'Boş' ? 'bg-emerald-100/50 text-emerald-700 border border-emerald-200/50' : 'bg-gray-100 text-gray-500 border border-gray-200'
                            }`}>
                              <span className={`w-1.5 h-1.5 rounded-full mr-1.5 ${prop.status === 'Boş' ? 'bg-emerald-500' : 'bg-gray-400'}`}></span>
                              {prop.status}
                            </span>
                          </td>
                          <td className="py-4 px-6 text-center">
                            <div className="flex flex-col items-center gap-1.5">
                              {isOverdue(prop) ? (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-bold bg-red-50 text-red-700 border border-red-200 animate-pulse">
                                  <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                                  Gecikmede
                                </span>
                              ) : prop.payment_status === '—' ? (
                                <span className="text-gray-300">—</span>
                              ) : (
                                <span className={`inline-flex px-2.5 py-1 rounded-md text-xs font-bold ${
                                  prop.payment_status === 'Ödendi' ? 'bg-blue-50 text-blue-700 border border-blue-200/50' :
                                  prop.payment_status === 'Bekliyor' ? 'bg-orange-50 text-orange-700 border border-orange-200' :
                                  'bg-red-50 text-red-700 border border-red-200'
                                }`}>
                                  {prop.payment_status}
                                </span>
                              )}
                              
                              {/* Ödeme Günü Bilgisi */}
                              <span className="text-[10px] text-gray-500 font-semibold tracking-wide flex items-center gap-1">
                                <svg className="w-3 h-3 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" /></svg>
                                Her ayın {prop.payment_due_day}. günü
                              </span>
                            </div>
                        </td>
                          
                          <td className="py-4 px-6">
                            {Number(prop.deduction_amount) > 0 ? (
                              <div>
                                <span className="text-xs font-bold text-orange-600 bg-orange-50 border border-orange-100 px-2 py-0.5 rounded-md inline-block mb-1">
                                  - ₺{Number(prop.deduction_amount).toLocaleString('tr-TR')}
                                </span>
                                <div className="text-[11px] text-gray-500 leading-tight" title={prop.deduction_note}>
                                  {prop.deduction_note?.length > 40 ? prop.deduction_note.substring(0, 40) + "..." : prop.deduction_note}
                                </div>
                              </div>
                            ) : (
                              <span className="text-gray-300 text-xs">—</span>
                            )}
                          </td>

                          <td className="py-4 px-6 text-right">
                            {Number(prop.deduction_amount) > 0 ? (
                              <div className="flex flex-col items-end">
                                <span className="text-xs text-gray-400 line-through">₺{Number(prop.rent_amount).toLocaleString('tr-TR')}</span>
                                <span className="font-black text-gray-900 text-base">₺{(Number(prop.rent_amount) - Number(prop.deduction_amount)).toLocaleString('tr-TR')}</span>
                              </div>
                            ) : (
                              <span className="font-black text-gray-900 text-base">₺{Number(prop.rent_amount).toLocaleString('tr-TR')}</span>
                            )}
                          </td>
                          <td className="py-4 px-6 text-right">
                            <button 
                              onClick={() => handleEditClick(prop)}
                              className="text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 p-2 rounded-lg transition-colors inline-flex"
                              title="Düzenle"
                            >
                              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z" />
                              </svg>
                            </button>
                            {/* YENİ SİLME BUTONU */}
                            <button 
                              onClick={(e) => {
                                e.stopPropagation(); // Satır çift tıklama olayını tetiklememesi için
                                if (prop.id) handleDeleteProperty(prop.id, prop.name);
                              }}
                              className="text-red-600 hover:text-red-800 bg-red-50 hover:bg-red-100 p-2 rounded-lg transition-colors inline-flex"
                              title="Mülkü Sil"
                            >
                              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-4v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                              </svg>
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

              </div>
            </div>
          )}

          {/* ─── 2. MÜLK ENVANTERİ SEKMESİ ─── */}
          {activeTab === 'properties' && (
            <div className="animate-fade-in">
              <div className="mb-8 flex justify-between items-end">
                <div>
                  <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Mülk Envanteri</h1>
                  <p className="text-gray-500 text-sm mt-1">Sistemde kayıtlı mülklerin fiziksel portföy görünümü.</p>
                </div>
                <button 
                  onClick={handleNewRecordClick}
                  className="bg-white border border-gray-200 text-gray-700 hover:bg-gray-50 font-medium text-sm px-4 py-2 rounded-xl transition-all shadow-sm flex items-center gap-2"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4v16m8-8H4" /></svg>
                  Hızlı Mülk Ekle
                </button>
              </div>

              {/* Kart Grid Yapısı */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                {isLoading ? (
                  <div className="col-span-full py-10 text-center text-gray-500 font-medium">Veriler yükleniyor...</div>
                ) : properties.length === 0 ? (
                  <div className="col-span-full py-16 text-center text-gray-500 bg-white rounded-2xl border border-gray-100 border-dashed">
                    Envanterde mülk bulunmuyor.
                  </div>
                ) : properties.map((prop) => (
                  <div key={prop.id} className="bg-white rounded-2xl border border-gray-100 shadow-sm hover:shadow-lg transition-all duration-300 overflow-hidden flex flex-col group">
                    
                    {/* Kart Üst Kısım (Header/Görsel Alan) */}
                    <div className="h-32 bg-[#09090B] relative overflow-hidden">
                      {/* Dekoratif Arka Plan Deseni */}
                      <div className="absolute inset-0 opacity-10 bg-[radial-gradient(circle_at_top_right,_var(--tw-gradient-stops))] from-blue-400 via-transparent to-transparent"></div>
                      
                      <div className="absolute top-4 right-4 z-10">
                        <span className={`px-2.5 py-1 rounded-md text-[10px] font-black uppercase tracking-wider shadow-sm backdrop-blur-md ${
                          prop.status === 'Boş' ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-gray-500/20 text-gray-300 border border-gray-500/30'
                        }`}>
                          {prop.status}
                        </span>
                      </div>
                      
                      <div className="absolute bottom-4 left-4 right-4 z-10">
                        <h3 className="text-white font-bold text-lg truncate" title={prop.name}>{prop.name}</h3>
                        <p className="text-gray-400 text-xs truncate flex items-center gap-1 mt-1" title={prop.address}>
                          <svg className="w-3.5 h-3.5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" /></svg>
                          {prop.address}
                        </p>
                      </div>
                    </div>
                    
                    {/* Kart İçerik */}
                    <div className="p-5 flex-1 flex flex-col justify-between">
                      <div className="space-y-4">
                        <div className="flex items-center justify-between border-b border-gray-50 pb-3">
                          <div>
                            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-0.5">Mevcut Kiracı</p>
                            <p className="text-sm font-semibold text-gray-800 truncate max-w-[120px]">
                              {prop.tenant && prop.tenant !== '—' ? prop.tenant : <span className="text-gray-400 italic">Atanmadı</span>}
                            </p>
                          </div>
                          <div className="text-right">
                            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-0.5">Aylık Kira</p>
                            <p className="text-sm font-black text-gray-900">₺{Number(prop.rent_amount).toLocaleString('tr-TR')}</p>
                          </div>
                        </div>

                        {/* Son Tahsilat Durumu Mini Gösterge */}
                        <div className="flex items-center gap-2">
                          <div className={`w-2 h-2 rounded-full ${prop.payment_status === 'Ödendi' ? 'bg-blue-500' : prop.payment_status === 'Gecikti' ? 'bg-red-500' : 'bg-orange-400'}`}></div>
                          <span className="text-xs text-gray-600 font-medium flex-1">Son Durum: {prop.payment_status}</span>
                          {prop.last_month && (
                            <span className="text-[10px] text-gray-400 bg-gray-100 px-2 py-0.5 rounded">{prop.last_month} {prop.last_year}</span>
                          )}
                        </div>
                      </div>

                      <div className="mt-3 pt-3 border-t border-gray-50 flex flex-col gap-1 text-xs text-gray-600">
                        <div className="flex justify-between">
                          <span className="text-gray-400 font-medium">Sözleşme Başlangıç:</span>
                          <span className="font-semibold text-gray-700">{prop.contract_start_date ? new Date(prop.contract_start_date).toLocaleDateString('tr-TR') : '—'}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-gray-400 font-medium">Sözleşme Bitiş:</span>
                          <span className={`font-semibold ${prop.contract_end_date ? 'text-gray-700' : 'text-gray-400'}`}>
                            {prop.contract_end_date ? new Date(prop.contract_end_date).toLocaleDateString('tr-TR') : '—'}
                          </span>
                        </div>
                      </div>
                      
                      {prop.note && (
                        <div className="mt-2 p-2 bg-gray-50 border border-gray-100 rounded-lg">
                          <p className="text-[11px] text-gray-500 italic line-clamp-2" title={prop.note}>
                            {prop.note}
                          </p>
                        </div>
                      )}

                      {/* ÇOKLU BELGE GÖRÜNTÜLEME ALANI */}
                      {(() => {
                        const docs = parsePaths(prop.contract_file_path);
                        if (docs.length > 0) {
                          return (
                            <div className="mt-3 space-y-1.5 max-h-[80px] overflow-y-auto pr-1">
                              {docs.map((docPath, idx) => (
                                <button 
                                  key={idx}
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    // @ts-ignore
                                    window.api.openPdf(docPath);
                                  }}
                                  className="w-full bg-blue-50/50 hover:bg-blue-100 text-blue-700 border border-blue-100 text-xs font-semibold py-1.5 px-2 rounded-md transition-all flex items-center justify-between group"
                                  title={docPath.split('\\').pop()?.split('/').pop()}
                                >
                                  <div className="flex items-center gap-1.5 overflow-hidden">
                                    <svg className="w-3.5 h-3.5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                                    <span className="truncate">
                                      Belge {docs.length > 1 ? idx + 1 : ''}
                                    </span>
                                  </div>
                                  <svg className="w-3 h-3 opacity-0 group-hover:opacity-100 transition-opacity" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" /></svg>
                                </button>
                              ))}
                            </div>
                          );
                        }
                        return null;
                      })()}

                      {/* Kart Aksiyon Butonları */}
                      <div className="pt-4 mt-4 border-t border-gray-100 flex gap-2">
                        <button 
                          onClick={() => handleEditClick(prop)} 
                          className="flex-1 bg-white border border-gray-200 hover:border-blue-300 hover:bg-blue-50 text-gray-700 hover:text-blue-700 text-xs font-bold py-2 rounded-lg transition-all text-center flex items-center justify-center gap-1.5"
                        >
                          <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z" /></svg>
                          Düzenle
                        </button>
                        <button 
                          onClick={() => handlePropertyDoubleClick(prop)} 
                          className="flex-1 bg-gray-900 hover:bg-gray-800 text-white text-xs font-bold py-2 rounded-lg shadow-md transition-all text-center"
                        >
                          Ödemeler
                        </button>
                        {/* YENİ SİLME BUTONU */}
                        <button 
                          onClick={() => {
                            if (prop.id) handleDeleteProperty(prop.id, prop.name);
                          }} 
                          className="bg-red-50 hover:bg-red-100 text-red-600 p-2 rounded-lg transition-all text-center flex items-center justify-center"
                          title="Mülkü Sil"
                        >
                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-4v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                          </svg>
                        </button>
                      </div>
                    </div>

                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ─── 3. FİNANS VE KİRA SEKMESİ ─── */}
          {activeTab === 'payments' && (
             <div className="animate-fade-in space-y-6">
                
                {/* Başlık ve Tarih */}
                <div className="flex items-end justify-between border-b border-gray-200/50 pb-4">
                  <div>
                    <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Finans & Kira Durumu</h1>
                    <p className="text-gray-500 text-sm mt-1">Nakit akışı, kesintiler ve tahsilat performans raporu.</p>
                  </div>
                  <div className="text-right">
                    <p className="text-xs font-bold text-gray-400 uppercase tracking-wider">Mevcut Dönem</p>
                    <p className="text-lg font-black text-blue-600">{new Date().toLocaleString('tr-TR', { month: 'long', year: 'numeric' })}</p>
                  </div>
                </div>

                {/* Finansal Hesaplamalar (IIFE kullanarak sadece bu sekme açıldığında hesaplatıyoruz) */}
                {(() => {
                  
                  const occupiedProps = properties.filter(p => p.status !== 'Boş');
                  
                  const brutKira = occupiedProps.reduce((acc, curr) => acc + (Number(curr.rent_amount) || 0), 0);
                  const toplamKesinti = occupiedProps.reduce((acc, curr) => acc + (Number(curr.deduction_amount) || 0), 0);
                  const netKasa = brutKira - toplamKesinti;
                  
                  const count = occupiedProps.length || 1; // Sıfıra bölünme hatasını önlemek için
                  const odendiSayisi = occupiedProps.filter(p => p.payment_status === 'Ödendi').length;
                  const geciktiSayisi = occupiedProps.filter(p => p.payment_status === 'Gecikti').length;
                  const bekleyenSayisi = occupiedProps.filter(p => p.payment_status === 'Bekliyor' || p.payment_status === '—').length;
                  
                  const odendiYuzde = (odendiSayisi / count) * 100;
                  const geciktiYuzde = (geciktiSayisi / count) * 100;
                  const bekleyenYuzde = (bekleyenSayisi / count) * 100;

                  return (
                    <>
                      {/* Premium Koyu Tema Metrik Kartları */}
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                        {/* Brüt Beklenen Kartı */}
                        <div className="bg-[#09090B] rounded-2xl p-6 relative overflow-hidden group transition-all duration-500 hover:-translate-y-1 hover:neon-shadow-blue border border-gray-800/60">
                          <div className="absolute inset-0 bg-gradient-to-br from-blue-600/10 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-500"></div>
                          <div className="relative z-10">
                            <h3 className="text-gray-400 text-xs font-bold uppercase tracking-wider mb-1">Brüt Beklenen</h3>
                            <div className="text-3xl font-black text-white">₺{brutKira.toLocaleString('tr-TR')}</div>
                            <p className="text-gray-500 text-[11px] mt-2 flex items-center gap-1.5">
                              <span className="w-1.5 h-1.5 rounded-full bg-blue-500 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></span>
                              Aktif kontratların toplamı
                            </p>
                          </div>
                        </div>

                        {/* Toplam Kesinti Kartı */}
                        <div className="bg-[#09090B] rounded-2xl p-6 relative overflow-hidden group transition-all duration-500 hover:-translate-y-1 hover:neon-shadow-orange border border-gray-800/60">
                          <div className="absolute inset-0 bg-gradient-to-br from-orange-500/10 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-500"></div>
                          <div className="relative z-10">
                            <h3 className="text-gray-400 text-xs font-bold uppercase tracking-wider mb-1">Toplam Kesinti</h3>
                            <div className="text-3xl font-black text-white">₺{toplamKesinti.toLocaleString('tr-TR')}</div>
                            <p className="text-gray-500 text-[11px] mt-2 flex items-center gap-1.5">
                              <span className="w-1.5 h-1.5 rounded-full bg-orange-500 shadow-[0_0_8px_rgba(249,115,22,0.8)]"></span>
                              Tadilat ve masraf düşümleri
                            </p>
                          </div>
                        </div>

                        {/* Net Kasa Kartı (Vurgulu) */}
                        <div className="bg-[#09090B] rounded-2xl p-6 relative overflow-hidden group transition-all duration-500 hover:-translate-y-1 hover:neon-shadow-emerald border border-gray-700">
                          <div className="absolute top-0 right-0 p-5">
                            <div className="animate-pulse flex h-3 w-3">
                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                              <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
                            </div>
                          </div>
                          <div className="absolute inset-0 bg-gradient-to-br from-emerald-500/10 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-500"></div>
                          <div className="relative z-10">
                            <h3 className="text-gray-300 text-xs font-bold uppercase tracking-wider mb-1">Aylık Net Kasa</h3>
                            <div className="text-4xl font-black text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-emerald-200">
                              ₺{netKasa.toLocaleString('tr-TR')}
                            </div>
                            <p className="text-emerald-500/80 text-[11px] mt-2 flex items-center gap-1.5 font-medium">
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                              Ele geçecek tahmini tutar
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* ─── YENİ EKLENEN: ORTAKLAR & NAKİT KASA DAĞILIMI ─── */}
                      <div className="mt-8 mb-4">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3">
                          <h3 className="font-bold text-gray-900 text-sm flex items-center gap-2">
                            <svg className="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z" /></svg>
                            Tüm Zamanlar Kasa ve Ortak Dağılımı
                          </h3>
                          <div className="flex gap-2">
                            <button 
                              onClick={() => { fetchTransferHistory(); setIsTransferHistoryOpen(true); }}
                              className="bg-gray-100 hover:bg-gray-200 text-gray-700 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm flex items-center gap-1.5"
                            >
                              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" /></svg>
                              Çıkış Geçmişi
                            </button>
                            <button 
                              onClick={() => setIsTransferFormOpen(true)}
                              className="bg-orange-50 hover:bg-orange-100 border border-orange-200 text-orange-700 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm flex items-center gap-1.5"
                            >
                              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4v16m8-8H4" /></svg>
                              Para Çıkışı Yap
                            </button>
                          </div>
                        </div>
                        <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                          
                          {/* NAKİT (KASA) KARTI */}
                        <div 
                          onClick={() => handleAccountClick('Nakit Kasa', 'Nakit')}
                          className="bg-emerald-50 border border-emerald-100 rounded-xl p-4 flex flex-col justify-center relative overflow-hidden group hover:shadow-lg hover:-translate-y-1 hover:border-emerald-300 transition-all cursor-pointer"
                        >
                          <div className="text-emerald-600 text-[10px] font-black uppercase tracking-wider mb-1">Nakit Kasa</div>
                          <div className="text-xl font-black text-emerald-900">₺{financeSummary.nakit_total.toLocaleString('tr-TR')}</div>
                          <div className="absolute -right-4 -bottom-4 opacity-10 group-hover:scale-110 group-hover:opacity-20 transition-all">
                            <svg className="w-16 h-16 text-emerald-600" fill="currentColor" viewBox="0 0 20 20"><path d="M4 4a2 2 0 00-2 2v1h16V6a2 2 0 00-2-2H4z" /><path fillRule="evenodd" d="M18 9H2v5a2 2 0 002 2h12a2 2 0 002-2V9zM4 13a1 1 0 011-1h1a1 1 0 110 2H5a1 1 0 01-1-1zm5-1a1 1 0 100 2h1a1 1 0 100-2H9z" clipRule="evenodd" /></svg>
                          </div>
                        </div>

                        {/* ORTAK KARTLARI */}
                        {[
                          { name: 'Tahir Çam', val: financeSummary.tahir_total, bg: 'bg-blue-50', border: 'border-blue-100', text: 'text-blue-600', valText: 'text-blue-900', method: 'Havale (Tahir Çam)' },
                          { name: 'Bekir Çam', val: financeSummary.bekir_total, bg: 'bg-indigo-50', border: 'border-indigo-100', text: 'text-indigo-600', valText: 'text-indigo-900', method: 'Havale (Bekir Çam)' },
                          { name: 'Tarık Çam', val: financeSummary.tarik_total, bg: 'bg-purple-50', border: 'border-purple-100', text: 'text-purple-600', valText: 'text-purple-900', method: 'Havale (Tarık Çam)' },
                          { name: 'Halil Çam', val: financeSummary.halil_total, bg: 'bg-sky-50', border: 'border-sky-100', text: 'text-sky-600', valText: 'text-sky-900', method: 'Havale (Halil Çam)' }
                        ].map((ortak, idx) => (
                          <div 
                            key={idx} 
                            onClick={() => handleAccountClick(ortak.name, ortak.method)}
                            className={`${ortak.bg} ${ortak.border} border rounded-xl p-4 flex flex-col justify-center transition-all hover:shadow-lg hover:-translate-y-1 hover:brightness-95 cursor-pointer group relative overflow-hidden`}
                          >
                            <div className={`${ortak.text} text-[10px] font-black uppercase tracking-wider mb-1`}>{ortak.name}</div>
                            <div className={`text-xl font-black ${ortak.valText} relative z-10`}>₺{ortak.val.toLocaleString('tr-TR')}</div>
                            <div className="absolute -right-3 -bottom-3 opacity-[0.03] group-hover:opacity-10 group-hover:scale-110 transition-all">
                              <svg className="w-16 h-16" fill="currentColor" viewBox="0 0 24 24"><path d="M12 12a5 5 0 110-10 5 5 0 010 10zm0-2a3 3 0 100-6 3 3 0 000 6zm9 11a1 1 0 01-2 0c0-2.336-2.833-4-7-4s-7 1.664-7 4a1 1 0 01-2 0c0-3.336 3.833-6 9-6s9 2.664 9 6z" /></svg>
                            </div>
                          </div>
                        ))}
                          
                        </div>
                      </div>

                      {/* Tahsilat İlerleme Çubuğu (Progress Bar) */}
                      <div className="bg-white rounded-2xl p-7 border border-gray-100 shadow-sm mt-2">
                        <div className="flex justify-between items-end mb-3">
                          <h3 className="font-bold text-gray-900 text-sm">Tahsilat Performansı</h3>
                          <span className="text-xs font-black text-gray-400 bg-gray-100 px-2 py-1 rounded-md">%{(odendiYuzde).toFixed(0)} TAMAMLANDI</span>
                        </div>
                        <div className="h-4 w-full bg-gray-100 rounded-full overflow-hidden flex shadow-inner">
                          {/* Yüzdelik dilimlere göre genişleyen renkli alanlar */}
                          <div className="h-full bg-gradient-to-r from-emerald-600 to-emerald-400 transition-all duration-1000 ease-out" style={{ width: `${odendiYuzde}%` }} title="Ödendi"></div>
                          <div className="h-full bg-gradient-to-r from-blue-500 to-blue-400 transition-all duration-1000 ease-out border-l border-white/20" style={{ width: `${bekleyenYuzde}%` }} title="Bekliyor"></div>
                          <div className="h-full bg-gradient-to-r from-red-600 to-red-500 transition-all duration-1000 ease-out border-l border-white/20" style={{ width: `${geciktiYuzde}%` }} title="Gecikti"></div>
                        </div>
                        <div className="flex gap-8 mt-5 border-t border-gray-50 pt-4">
                          <div className="flex items-center gap-2">
                            <span className="w-3 h-3 rounded bg-emerald-500 shadow-sm"></span>
                            <span className="text-sm text-gray-700 font-bold">{odendiSayisi} <span className="text-gray-400 font-normal">Ödendi</span></span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="w-3 h-3 rounded bg-blue-400 shadow-sm"></span>
                            <span className="text-sm text-gray-700 font-bold">{bekleyenSayisi} <span className="text-gray-400 font-normal">Bekliyor</span></span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="w-3 h-3 rounded bg-red-500 shadow-sm"></span>
                            <span className="text-sm text-gray-700 font-bold">{geciktiSayisi} <span className="text-gray-400 font-normal">Gecikti</span></span>
                          </div>
                        </div>
                      </div>

                      {/* Şık Finans Döküm Tablosu */}
                      <div className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden mt-6">
                        <div className="px-6 py-5 border-b border-gray-800 bg-[#09090B] flex justify-between items-center">
                          <h3 className="font-bold text-white text-sm flex items-center gap-2">
                            <svg className="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                            Aktif Kontratlar Finansal Dökümü
                          </h3>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-gray-50/80 border-b border-gray-100 text-gray-400 text-[11px] font-black uppercase tracking-wider">
                                <th className="py-4 px-6 w-1/3">Mülk & Kiracı</th>
                                <th className="py-4 px-6 text-right">Brüt Kira</th>
                                <th className="py-4 px-6 text-right">Kesinti (Tadilat)</th>
                                <th className="py-4 px-6 text-right">Net Ödenecek</th>
                                <th className="py-4 px-6 text-center">Durum</th>
                                <th className="py-4 px-6 text-center">Detay</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-50 text-sm">
                              {occupiedProps.length === 0 ? (
                                <tr>
                                  <td colSpan={6} className="py-12 text-center text-gray-400 font-medium">Finansal döküm için sistemde 'Dolu' mülk bulunmuyor.</td>
                                </tr>
                              ) : occupiedProps.map((prop, index) => {
                                const rent = Number(prop.rent_amount) || 0;
                                const deduction = Number(prop.deduction_amount) || 0;
                                const net = rent - deduction;
                                
                                return (
                                  <tr 
                                    key={prop.id} 
                                    className="hover:bg-blue-50/40 transition-colors group" 
                                    // CSS'e eklediğimiz animasyonun her satırda sırayla çalışması için gecikme ekliyoruz
                                    style={{ animationDelay: `${index * 50}ms` }}
                                  >
                                    <td className="py-4 px-6">
                                      <div className="font-bold text-gray-900 group-hover:text-blue-700 transition-colors">{prop.name}</div>
                                      <div className="text-xs text-gray-500 mt-1 flex items-center gap-1.5">
                                        <div className="w-4 h-4 rounded-full bg-gray-200 flex items-center justify-center text-[8px] font-bold text-gray-600">
                                          {prop.tenant.substring(0,1)}
                                        </div>
                                        {prop.tenant}
                                      </div>
                                    </td>
                                    
                                    <td className="py-4 px-6 text-right font-medium text-gray-600">
                                      ₺{rent.toLocaleString('tr-TR')}
                                    </td>
                                    
                                    <td className="py-4 px-6 text-right">
                                      {deduction > 0 ? (
                                        <div className="inline-flex flex-col items-end">
                                          <span className="text-xs font-bold text-orange-600 bg-orange-50 px-2 py-0.5 rounded border border-orange-100 shadow-sm">
                                            - ₺{deduction.toLocaleString('tr-TR')}
                                          </span>
                                          {prop.deduction_note && (
                                            <span className="text-[10px] text-gray-400 mt-1.5 truncate max-w-[140px] italic" title={prop.deduction_note}>
                                              {prop.deduction_note}
                                            </span>
                                          )}
                                        </div>
                                      ) : (
                                        <span className="text-gray-300">—</span>
                                      )}
                                    </td>
                                    
                                    <td className="py-4 px-6 text-right">
                                      <span className="font-black text-gray-900 text-base">₺{net.toLocaleString('tr-TR')}</span>
                                    </td>
                                    
                                    <td className="py-4 px-6 text-center">
                                      <span className={`inline-flex items-center justify-center min-w-[80px] gap-1.5 px-3 py-1.5 rounded-md text-[10px] font-black uppercase tracking-wider shadow-sm ${
                                        prop.payment_status === 'Ödendi' ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' :
                                        prop.payment_status === 'Gecikti' ? 'bg-red-50 text-red-700 border border-red-200' :
                                        'bg-blue-50 text-blue-700 border border-blue-200'
                                      }`}>
                                        {prop.payment_status === 'Gecikti' && <span className="w-1.5 h-1.5 rounded-full bg-red-500 animate-pulse"></span>}
                                        {prop.payment_status}
                                      </span>
                                    </td>
                                    
                                    <td className="py-4 px-6 text-center">
                                      <button 
                                        onClick={() => handlePropertyDoubleClick(prop)}
                                        className="text-gray-400 hover:text-blue-600 p-2 rounded-lg hover:bg-white border border-transparent hover:border-blue-100 hover:shadow-sm transition-all inline-flex bg-gray-50 group-hover:bg-transparent"
                                        title="Geçmiş Aylara Git"
                                      >
                                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5l7 7-7 7" /></svg>
                                      </button>
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </>
                  );
                })()}
             </div>
          )}

{/* ─── 4. KİRACILAR SEKMESİ ─── */}
          {activeTab === 'tenants' && (
             <div className="animate-fade-in space-y-6">
                
                {/* Başlık */}
                <div className="flex justify-between items-end border-b border-gray-200/50 pb-4">
                  <div>
                    <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Kiracı Fihristi</h1>
                    <p className="text-gray-500 text-sm mt-1">Aktif kiracıların iletişim bilgileri ve mülk eşleşmeleri.</p>
                  </div>
                </div>
                

                {(() => {
                  // Sadece durumu 'Dolu' olan ve atanmış bir kiracısı bulunan mülkleri filtrele
                  const tenantsList = properties.filter(p => p.status !== 'Boş' && p.tenant && p.tenant.trim() !== '' && p.tenant !== '—');

                  if (tenantsList.length === 0) {
                    return (
                      <div className="py-16 text-center text-gray-500 bg-white rounded-2xl border border-gray-100 border-dashed flex flex-col items-center">
                        <svg className="w-12 h-12 text-gray-300 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197M13 7a4 4 0 11-8 0 4 4 0 018 0z" /></svg>
                        Sistemde kayıtlı aktif bir kiracı bulunmuyor.
                      </div>
                    );
                  }

                  return (
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                      {tenantsList.map((prop, index) => {
                        // İsimden baş harfleri çıkartan küçük bir algoritma (Örn: Gökhan Bey -> GB)
                        const initials = prop.tenant.trim().split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase() || 'K';

                        return (
                          <div 
                            key={prop.id} 
                            className="bg-white rounded-2xl border border-gray-100 shadow-sm hover:shadow-xl hover:-translate-y-1 transition-all duration-300 overflow-hidden group"
                            style={{ animationDelay: `${index * 50}ms` }}
                          >
                            {/* Kart Üst (Header) - Gradient Arka Plan */}
                            <div className="h-20 bg-gradient-to-r from-blue-700 to-blue-900 relative">
                              <div className="absolute -bottom-8 left-6">
                                <div className="w-16 h-16 rounded-full bg-[#09090B] border-4 border-white flex items-center justify-center text-xl font-black text-white shadow-md">
                                  {initials}
                                </div>
                              </div>
                              <div className="absolute top-3 right-3">
                                <span className="bg-white/20 backdrop-blur-md text-white text-[10px] font-bold px-2 py-1 rounded shadow-sm border border-white/10 uppercase tracking-wider">
                                  Aktif
                                </span>
                              </div>
                            </div>
                            
                            {/* Kart Gövdesi (Body) */}
                            <div className="pt-10 p-6">
                              <h3 className="font-bold text-gray-900 text-lg truncate" title={prop.tenant}>{prop.tenant}</h3>
                              <p className="text-blue-600 text-xs font-bold mt-0.5 truncate flex items-center gap-1">
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" /></svg>
                                {prop.name}
                              </p>
                              
                              <div className="mt-5 space-y-3">
                                {/* Telefon Gösterimi */}
                                <div className="flex items-center gap-3 text-sm">
                                  <div className={`p-1.5 rounded-md ${prop.tenant_phone ? 'bg-blue-50 text-blue-600' : 'bg-gray-50 text-gray-400'}`}>
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z" /></svg>
                                  </div>
                                  {prop.tenant_phone ? (
                                    <span className="text-gray-700 font-medium">{prop.tenant_phone}</span>
                                  ) : (
                                    <span className="italic text-[11px] font-medium tracking-wide text-gray-400">Numara girilmedi</span>
                                  )}
                                </div>
                                
                                {/* E-posta Gösterimi */}
                                <div className="flex items-center gap-3 text-sm mt-3">
                                  <div className={`p-1.5 rounded-md ${prop.tenant_email ? 'bg-blue-50 text-blue-600' : 'bg-gray-50 text-gray-400'}`}>
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" /></svg>
                                  </div>
                                  {prop.tenant_email ? (
                                    <span className="text-gray-700 font-medium">{prop.tenant_email}</span>
                                  ) : (
                                    <span className="italic text-[11px] font-medium tracking-wide text-gray-400">E-posta girilmedi</span>
                                  )}
                                </div>

                                {/* Kira Bedeli */}
                                <div className="flex items-center justify-between border-t border-gray-50 pt-4 mt-4">
                                  <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Aylık Kira Bedeli</span>
                                  <span className="font-black text-gray-900 text-sm">₺{Number(prop.rent_amount).toLocaleString('tr-TR')}</span>
                                </div>
                              </div>
                              
                              {/* Aksiyon Butonu */}
                              <div className="mt-4 pt-4 border-t border-gray-100">
                                <button 
                                  onClick={() => handleEditClick(prop)}
                                  className="w-full bg-white border border-gray-200 hover:border-blue-400 hover:text-blue-700 hover:bg-blue-50 text-gray-600 text-xs font-bold py-2.5 rounded-xl transition-all shadow-sm flex items-center justify-center gap-2"
                                >
                                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" /></svg>
                                  Kiracı Bilgisini Güncelle
                                </button>
                              </div>

                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
             </div>
          )}
        {/* ─── 5. AYARLAR SEKMESİ ─── */}
          {activeTab === 'settings' && (
             <div className="animate-fade-in space-y-6">
                
                <div className="flex justify-between items-end border-b border-gray-200/50 pb-4">
                  <div>
                    <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Sistem Ayarları</h1>
                    <p className="text-gray-500 text-sm mt-1">Veritabanı yönetimi ve dışa aktarım işlemleri.</p>
                  </div>
                </div>

                {/* VERİTABANI YEDEKLEME KARTI */}
                <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-8 max-w-2xl">
                  <h3 className="font-bold text-gray-900 text-lg mb-2">Veritabanı Yedekleme</h3>
                  <p className="text-sm text-gray-500 mb-6">
                    Sistemdeki tüm mülk, kiracı ve finans kayıtlarını dışa aktararak güvenli bir kopyasını alabilirsiniz. 
                    Farklı bir bilgisayara veri aktarımı yapmak veya yedekten dönmek (içeri aktarmak) için teknik destek ile iletişime geçiniz.
                  </p>
                  
                  <button 
                    onClick={async () => {
                      // @ts-ignore
                      const res = await window.api.exportDatabase();
                      if (res?.success) {
                        alert("Veritabanı başarıyla dışa aktarıldı!\n\nKonum: " + res.path);
                      }
                    }}
                    className="bg-blue-600 hover:bg-blue-700 text-white font-medium text-sm px-6 py-3 rounded-xl transition-all flex items-center gap-2 shadow-lg shadow-blue-500/30"
                  >
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
                    Veritabanını Dışa Aktar
                  </button>
                  
                  <div className="mt-8 p-4 bg-orange-50 border border-orange-100 rounded-xl flex gap-3">
                    <svg className="w-6 h-6 text-orange-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                    <div className="text-sm text-orange-800">
                      <span className="font-bold">Bilgilendirme:</span> Dışa aktarılan veritabanını sisteme geri yükleme işlemi güvenlik protokolleri gereği kısıtlanmıştır. Kurulum gerektiren işlemler için bizlerle iletişime geçiniz.
                    </div>
                  </div>
                </div>

                {/* YENİ EKLENEN: İLETİŞİM BİLGİLERİ KARTI */}
                <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-8 max-w-2xl">
                  <h3 className="font-bold text-gray-900 text-lg mb-4 flex items-center gap-2">
                    <svg className="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M18.364 5.636l-3.536 3.536m0 5.656l3.536 3.536M9.172 9.172L5.636 5.636m3.536 9.192l-3.536 3.536M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-5 0a4 4 0 11-8 0 4 4 0 018 0z" /></svg>
                    Teknik Destek & İletişim
                  </h3>
                  <p className="text-sm text-gray-500 mb-6">
                    Uygulama ile ilgili teknik destek talepleriniz, veri taşıma işlemleri veya yeni özellik önerileriniz için aşağıdaki kanallardan bize ulaşabilirsiniz.
                  </p>
                  
                  <div className="space-y-4">
                    {/* Adres Bölümü */}
                    <div className="flex items-start gap-4 p-4 bg-gray-50 rounded-xl border border-gray-100">
                      <div className="bg-blue-100 p-2 rounded-lg text-blue-600 shrink-0">
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" /></svg>
                      </div>
                      <div>
                        <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-0.5">Merkez Ofis</div>
                        <div className="text-sm font-semibold text-gray-800">Cızırtı Bilişim</div>
                        <div className="text-sm text-gray-600 mt-1">Kunduracılar Çarşısı Kat:2 No:120 Sivas Merkez, Türkiye</div>
                      </div>
                    </div>

                    {/* Telefon ve E-Posta Bölümü */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="flex items-center gap-4 p-4 bg-gray-50 rounded-xl border border-gray-100">
                        <div className="bg-emerald-100 p-2 rounded-lg text-emerald-600 shrink-0">
                          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z" /></svg>
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-0.5">Telefon</div>
                          <div className="text-sm font-semibold text-gray-800">+90 544 911 60 98</div>
                        </div>
                      </div>

                      <div className="flex items-center gap-4 p-4 bg-gray-50 rounded-xl border border-gray-100">
                        <div className="bg-purple-100 p-2 rounded-lg text-purple-600 shrink-0">
                          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" /></svg>
                        </div>
                        <div>
                          <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-0.5">E-Posta</div>
                          <div className="text-sm font-semibold text-gray-800">cizirtibilisim@gmail.com</div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
                
             </div>
          )}


        </div>
      </main>

      {/* ─── YENİ KAYIT MODAL ─── */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-gray-900/40 backdrop-blur-sm transition-all">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xl overflow-hidden flex flex-col">
            
            <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
              <h3 className="text-lg font-bold text-gray-900">
                {formData.id ? "Mülkü Düzenle" : "Yeni Mülk Ekle"}
              </h3>
              <button 
                onClick={() => setIsModalOpen(false)}
                className="text-gray-400 hover:text-red-500 transition-colors p-1"
              >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>

            <form onSubmit={handleSubmit} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-1 md:col-span-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Mülk Adı</label>
                  <input type="text" name="name" value={formData.name} onChange={handleInputChange} required className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 bg-gray-50 focus:bg-white" placeholder="Örn: Gül Apt." />
                </div>
                <div className="space-y-1 md:col-span-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Aylık Kira (₺)</label>
                  <input type="number" name="rent_amount" value={formData.rent_amount} onChange={handleInputChange} required className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 bg-gray-50 focus:bg-white" placeholder="Örn: 22000" />
                </div>
                {/* YENİ EKLENEN ÖDEME GÜNÜ ALANI */}
                <div className="space-y-1 md:col-span-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Ödeme Günü</label>
                  <div className="relative">
                    <span className="absolute left-3 top-2.5 text-sm text-gray-400 font-bold">Ayın</span>
                    <input type="number" min="1" max="31" name="payment_due_day" value={formData.payment_due_day} onChange={handleInputChange} required className="w-full border border-gray-200 rounded-lg pl-12 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 bg-gray-50 focus:bg-white" placeholder="1-31" />
                  </div>
                </div>
              </div>           
              
              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Açık Adres</label>
                <input type="text" name="address" value={formData.address} onChange={handleInputChange} required className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" placeholder="Örn: Kadıköy, İstanbul" />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Kiracı Adı Soyadı</label>
                <input type="text" name="tenant" value={formData.tenant} onChange={handleInputChange} className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" placeholder="Mülk boş ise boş bırakın" />
              </div>
              {/* YENİ EKLENEN İLETİŞİM ALANLARI */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Kiracı Telefon</label>
                  <input type="text" name="tenant_phone" value={formData.tenant_phone || ''} onChange={handleInputChange} className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" placeholder="Örn: 0555 555 55 55" />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Kiracı E-Posta</label>
                  <input type="email" name="tenant_email" value={formData.tenant_email || ''} onChange={handleInputChange} className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" placeholder="Örn: ornek@mail.com" />
                </div>
              </div>

              {/* YENİ EKLENEN SÖZLEŞME TARİHLERİ ALANI */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Sözleşme Başlangıç Tarihi</label>
                  <input 
                    type="date" 
                    name="contract_start_date" 
                    value={formData.contract_start_date || ''} 
                    onChange={handleInputChange} 
                    className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" 
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Sözleşme Bitiş Tarihi</label>
                  <input 
                    type="date" 
                    name="contract_end_date" 
                    value={formData.contract_end_date || ''} 
                    onChange={handleInputChange} 
                    className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white" 
                  />
                </div>
              </div>

              {/* PDF YÜKLEME VE LİSTELEME ALANI (GÜNCELLENDİ) */}
              <div className="space-y-2 md:col-span-2 pt-2 border-t border-gray-100">
                <label className="text-xs font-bold text-gray-500 uppercase tracking-wide flex justify-between items-center">
                  Sözleşme / Evraklar (PDF)
                </label>
                
                <div className="flex flex-col gap-3">
                  <button
                    type="button"
                    onClick={handlePdfSelect}
                    className="px-4 py-2.5 bg-gray-100 hover:bg-gray-200 border border-gray-300 text-gray-700 text-sm font-semibold rounded-lg transition-all flex items-center justify-center gap-2 w-full md:w-auto"
                  >
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 4v16m8-8H4" /></svg>
                    Belge Ekle
                  </button>
                  
                  {/* Eklenen Dosyaların Listesi */}
                  <div className="flex flex-wrap gap-2 mt-2">
                    {parsePaths(formData.contract_file_path).map((path, idx) => {
                      // Dosya yolundan sadece dosya adını çekme (Windows ve Mac uyumlu)
                      const fileName = path.split('\\').pop()?.split('/').pop();
                      return (
                        <div key={idx} className="flex items-center gap-2 bg-emerald-50 border border-emerald-100 pl-3 pr-1 py-1.5 rounded-lg shadow-sm">
                          <svg className="w-3.5 h-3.5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" /></svg>
                          <span className="text-xs font-semibold text-emerald-700 truncate max-w-[180px]" title={path}>
                            {fileName}
                          </span>
                          <button
                            type="button"
                            onClick={() => removePdf(idx)}
                            className="text-gray-400 hover:text-red-500 hover:bg-red-50 p-1.5 rounded-md transition-colors"
                            title="Dosyayı Kaldır"
                          >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                          </button>
                        </div>
                      )
                    })}
                    
                    {parsePaths(formData.contract_file_path).length === 0 && (
                      <span className="text-xs text-gray-400 italic py-2">Henüz mülke ait belge yüklenmedi.</span>
                    )}
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Mülk Durumu</label>
                  <select name="status" value={formData.status} onChange={handleInputChange} className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white appearance-none cursor-pointer">
                    <option value="Boş">Boş</option>
                    <option value="Dolu - Eşyasız">Dolu - Eşyasız</option>
                    <option value="Dolu - Eşyalı">Dolu - Eşyalı</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Tahsilat</label>
                  <select name="payment_status" value={formData.payment_status} onChange={handleInputChange} className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white appearance-none cursor-pointer">
                    <option value="—">—</option>
                    <option value="Ödendi">Ödendi</option>
                    <option value="Gecikti">Gecikti</option>
                  </select>
                </div>
              </div>

              <div className="space-y-1">
              <label className="text-xs font-bold text-gray-500 uppercase tracking-wide">Mülk Hakkında Özel Notlar</label>
              <textarea 
                name="note" 
                value={formData.note || ''} 
                onChange={handleInputChange} 
                rows={3}
                placeholder="Mülkle ilgili tapu bilgileri, abonelik numaraları veya özel şartlar..." 
                className="w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-gray-50 focus:bg-white resize-none"
              />
            </div>

              <div className="pt-6 flex justify-end gap-3 mt-2 border-t border-gray-100">
                <button type="button" onClick={() => setIsModalOpen(false)} className="px-5 py-2.5 text-sm font-bold text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 rounded-xl transition-all">
                  İptal
                </button>
                <button type="submit" className="px-5 py-2.5 text-sm font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-lg shadow-blue-500/30 transition-all">
                  {formData.id ? "Güncelle" : "Kaydet"}
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

      {/* ─── ÖDEME GEÇMİŞİ VE AYLIK TAKİP MODALI ─── */}
      {isPaymentModalOpen && selectedProperty && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-gray-900/40 backdrop-blur-sm transition-all">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl overflow-hidden flex flex-col max-h-[90vh]">
            
            {/* Başlık Alanı */}
            <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-[#09090B] text-white">
              <div>
                <h3 className="text-lg font-bold">{selectedProperty.name} — Ödeme Geçmişi</h3>
                <p className="text-xs text-gray-400 mt-0.5">{selectedProperty.tenant || 'Kiracı Yok'} | Sabit Kira: ₺{Number(selectedProperty.rent_amount).toLocaleString('tr-TR')}</p>
              </div>
              <button 
                onClick={() => { setIsPaymentModalOpen(false); setIsPaymentFormOpen(false); }}
                className="text-gray-400 hover:text-red-400 transition-colors p-1"
              >
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>

            <div className="p-6 overflow-y-auto flex-1 grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* Sol Taraf: 12 Ayın Grid Listesi */}
              <div className="md:col-span-2 space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <h4 className="font-bold text-gray-900 text-sm uppercase tracking-wide">Dönem Durumu:</h4>
                    {/* YENİ: Yıl Seçici Dropdown */}
                    <select 
                      value={selectedYear}
                      onChange={(e) => setSelectedYear(Number(e.target.value))}
                      className="bg-gray-100 border-none rounded-lg px-3 py-1 font-bold text-sm text-gray-800 focus:ring-2 focus:ring-blue-500/20 outline-none cursor-pointer"
                    >
                      {[2024, 2025, 2026, 2027, 2028].map(year => (
                        <option key={year} value={year}>{year} Yılı</option>
                      ))}
                    </select>
                  </div>
                  <span className="text-xs bg-blue-50 text-blue-700 px-2.5 py-1 rounded-md font-semibold">Çift tıklanan mülk detayları</span>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {monthsList.map((month) => {
                    const payment = getPaymentForMonth(month);
                    let statusStyle = "bg-gray-50 border-gray-200 text-gray-400";
                    if (payment?.status === 'Ödendi') statusStyle = "bg-emerald-50 border-emerald-200 text-emerald-800 shadow-sm";
                    if (payment?.status === 'Gecikti') statusStyle = "bg-red-50 border-red-200 text-red-800 shadow-sm";

                    return (
                      <div 
                        key={month}
                        onClick={() => {
                          setPaymentFormData({
                            id: payment?.id || undefined,
                            month: month,
                            year: 2026,
                            status: payment?.status || 'Ödendi',
                            payment_method: payment?.payment_method || 'Nakit',
                            amount_paid: payment ? payment.amount_paid.toString() : (Number(selectedProperty.rent_amount) - Number(selectedProperty.deduction_amount)).toString(),
                            deduction_amount: payment?.deduction_amount?.toString() || '',
                            payment_date: payment?.payment_date || new Date().toISOString().split('T')[0],
                            note: payment?.note || ''
                          });
                          setIsPaymentFormOpen(true);
                        }}
                        className={`border p-3 rounded-xl transition-all cursor-pointer hover:scale-[1.02] relative group flex flex-col justify-between h-24 ${statusStyle}`}
                      >
                        <div className="flex justify-between items-start">
                            <div className="flex flex-col gap-1">
                              <span className="font-bold text-sm text-gray-900">{month}</span>
                              {/* EĞER NOT VARSA GÖSTERİLECEK ROZET */}
                              {payment?.note && (
                                <span 
                                  title={payment.note} 
                                  className="cursor-help inline-flex items-center gap-1 bg-orange-100 text-orange-700 border border-orange-200 text-[9px] font-black uppercase px-1.5 py-0.5 rounded"
                                >
                                  <svg className="w-2.5 h-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                                  Not
                                </span>
                              )}
                            </div>
                            <span className={`w-2 h-2 rounded-full shadow-sm mt-1 ${payment?.status === 'Ödendi' ? 'bg-emerald-500' : payment?.status === 'Gecikti' ? 'bg-red-500' : 'bg-gray-300'}`}></span>
                          </div>
                        
                        {payment ? (
                          <div className="mt-2">
                            <div className="text-xs font-black">₺{Number(payment.amount_paid).toLocaleString('tr-TR')}</div>
                            <div className="text-[10px] opacity-75 mt-0.5">{payment.payment_date}</div>
                          </div>
                        ) : (
                          <span className="text-[11px] text-gray-400 italic mt-auto">Kayıt Girilmedi</span>
                        )}
                        
                        <div className="absolute inset-0 bg-blue-600/5 rounded-xl opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center text-xs font-bold text-blue-700">
                          {payment ? 'Düzenle' : 'Ödeme Gir'}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Sağ Taraf: Dinamik Ödeme Giriş / Düzenleme Formu */}
              <div className="border-l border-gray-100 pl-0 md:pl-6">
                {isPaymentFormOpen ? (
                  <form onSubmit={handlePaymentSubmit} className="space-y-4 sticky top-0">
                    <div className="flex justify-between items-center border-b border-gray-100 pb-2">
                      <h4 className="font-black text-gray-900 text-sm">
                        {paymentFormData.id ? `${paymentFormData.month} Ayını Düzenle` : `${paymentFormData.month} İçin Ödeme Gir`}
                      </h4>
                      {/* YENİ: SİL BUTONU (Sadece önceden kaydedilmiş aylarda görünür) */}
                      {paymentFormData.id && (
                        <button 
                          type="button" 
                          onClick={handleDeletePayment}
                          className="text-xs font-bold text-red-500 hover:text-red-700 bg-red-50 hover:bg-red-100 px-2 py-1 rounded transition-colors"
                        >
                          Kaydı Sil
                        </button>
                      )}
                    </div>
                    
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-gray-500 uppercase">Tahsilat Durumu</label>
                      <select 
                        value={paymentFormData.status} 
                        onChange={(e) => setPaymentFormData(prev => ({ ...prev, status: e.target.value }))}
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20"
                      >
                        <option value="Ödendi">Ödendi</option>
                        <option value="Gecikti">Gecikti</option>
                        <option value="Bekliyor">Bekliyor</option>
                      </select>
                    </div>
                    
                    {/* GÜNCELLENEN: Ödeme Yöntemi Seçimi */}
                    <div className="space-y-1 mt-3">
                      <label className="text-xs font-bold text-gray-500 uppercase">Tahsilat Yöntemi (Hesap)</label>
                      <select 
                        value={paymentFormData.payment_method} 
                        onChange={(e) => setPaymentFormData(prev => ({ ...prev, payment_method: e.target.value }))}
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20"
                      >
                        <option value="Nakit">Nakit (Kasa)</option>
                        <option value="Havale (Tahir Çam)">Havale (Tahir Çam)</option>
                        <option value="Havale (Bekir Çam)">Havale (Bekir Çam)</option>
                        <option value="Havale (Tarık Çam)">Havale (Tarık Çam)</option>
                        <option value="Havale (Halil Çam)">Havale (Halil Çam)</option>
                      </select>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-gray-500 uppercase">Alınan Net (₺)</label>
                        <input 
                          type="number" 
                          value={paymentFormData.amount_paid} 
                          onChange={(e) => setPaymentFormData(prev => ({ ...prev, amount_paid: e.target.value }))}
                          required
                          className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-orange-500 uppercase">Kesinti (₺)</label>
                        <input 
                          type="number" 
                          value={paymentFormData.deduction_amount} 
                          onChange={(e) => setPaymentFormData(prev => ({ ...prev, deduction_amount: e.target.value }))}
                          className="w-full border border-orange-200 rounded-lg px-3 py-2 text-sm bg-orange-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                          placeholder="Yoksa boş bırak"
                        />
                      </div>
                    </div>

                    <div className="space-y-1">
                      <label className="text-xs font-bold text-gray-500 uppercase">Ödeme Tarihi</label>
                      <input 
                        type="date" 
                        value={paymentFormData.payment_date} 
                        onChange={(e) => setPaymentFormData(prev => ({ ...prev, payment_date: e.target.value }))}
                        required
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20"
                      />
                    </div>

                    <div className="space-y-1">
                      <label className="text-xs font-bold text-gray-500 uppercase flex items-center gap-1.5">
                        Not / Tadilat / Kesinti
                        {paymentFormData.note && <span className="w-1.5 h-1.5 rounded-full bg-orange-500"></span>}
                      </label>
                      <textarea 
                        rows={3}
                        value={paymentFormData.note} 
                        onChange={(e) => setPaymentFormData(prev => ({ ...prev, note: e.target.value }))}
                        placeholder="Örn: Kombi bakımı yapıldığı için kiradan 1500 TL kesilerek yatırıldı..."
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/30 focus:border-orange-500 transition-all resize-none shadow-sm"
                      />
                    </div>

                    <div className="pt-2 flex gap-2">
                      <button 
                        type="button" 
                        onClick={() => setIsPaymentFormOpen(false)}
                        className="w-1/3 bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-bold py-2 rounded-lg transition-colors"
                      >
                        İptal
                      </button>
                      <button 
                        type="submit" 
                        className="w-2/3 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold py-2 rounded-lg shadow-md shadow-blue-500/10 transition-colors"
                      >
                        {paymentFormData.id ? 'Güncelle' : 'Kaydet'}
                      </button>
                    </div>
                  </form>
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-center p-4 text-gray-400 border-2 border-dashed border-gray-100 rounded-xl">
                    <svg className="w-8 h-8 mb-2 opacity-60" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 15l-2 5L9 9l11 4-5 2zm0 0l5 5M7.188 2.239l.777 2.897M5.136 7.965l-2.898-.777M13.95 4.05l-2.122 2.122m-5.657 5.656l-2.12 2.122" /></svg>
                    <p className="text-xs font-medium">İşlem yapmak için soldan bir aya tıklayın.</p>
                  </div>
                )}
              </div>

            </div>
            
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end">
              <button 
                onClick={() => { setIsPaymentModalOpen(false); setIsPaymentFormOpen(false); }}
                className="px-5 py-2 text-sm font-bold bg-gray-900 text-white hover:bg-gray-800 rounded-xl transition-colors"
              >
                Kapat
              </button>
            </div>

          </div>
        </div>
      )}

      {/* ─── HESAP DETAYLARI (DAİRELER) MODALI ─── */}
      {isAccountModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-gray-900/60 backdrop-blur-sm transition-all animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden flex flex-col max-h-[85vh]">
            
            <div className="px-6 py-5 border-b border-gray-800 bg-[#09090B] flex justify-between items-center">
              <div>
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <svg className="w-5 h-5 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z" /></svg>
                  {selectedAccountName} — Kasa Dökümü
                </h3>
                <p className="text-xs text-gray-400 mt-1">Bu hesaba giriş yapan mülklerin detaylı listesi.</p>
              </div>
              <button 
                onClick={() => setIsAccountModalOpen(false)}
                className="text-gray-400 hover:text-white bg-gray-800 hover:bg-gray-700 rounded-lg p-2 transition-colors"
              >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>

            <div className="p-0 overflow-y-auto flex-1 bg-gray-50/50">
              {accountDetails.length === 0 ? (
                <div className="py-16 flex flex-col items-center justify-center text-gray-400">
                  <div className="w-12 h-12 mb-3 rounded-full bg-gray-100 flex items-center justify-center animate-pulse">
                    <svg className="w-6 h-6 text-gray-300" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                  </div>
                  <p className="text-sm font-medium">Bu hesaba ait tahsilat bulunamadı veya aranıyor...</p>
                </div>
              ) : (
                <table className="w-full text-left border-collapse bg-white">
                  <thead>
                    <tr className="bg-gray-100/80 border-b border-gray-200 text-gray-500 text-[11px] font-black uppercase tracking-wider sticky top-0 shadow-sm">
                      <th className="py-4 px-6">Mülk Adı / Kiracı</th>
                      <th className="py-4 px-6">Ait Olduğu Ay</th>
                      <th className="py-4 px-6">İşlem Tarihi</th>
                      <th className="py-4 px-6 text-right">Tahsil Edilen (Net)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-sm">
                    {accountDetails.map((detail, idx) => (
                      <tr key={idx} className="hover:bg-blue-50/40 transition-colors group">
                        <td className="py-3 px-6">
                          <div className="font-bold text-gray-900">{detail.property_name || 'Bilinmeyen Mülk'}</div>
                          <div className="text-xs text-gray-500 mt-0.5">{detail.tenant || 'Bilinmeyen Kiracı'}</div>
                        </td>
                        <td className="py-3 px-6">
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-bold bg-blue-50 text-blue-700 border border-blue-100">
                            {detail.month} {detail.year}
                          </span>
                        </td>
                        <td className="py-3 px-6 text-gray-500 font-medium">
                          {detail.payment_date ? new Date(detail.payment_date).toLocaleDateString('tr-TR') : '—'}
                        </td>
                        <td className="py-3 px-6 text-right">
                          <span className="font-black text-emerald-600 text-base">₺{Number(detail.amount_paid).toLocaleString('tr-TR')}</span>
                          {Number(detail.deduction_amount) > 0 && (
                            <div className="text-[10px] text-orange-500 font-bold mt-1">
                              (-₺{Number(detail.deduction_amount).toLocaleString('tr-TR')} Kesinti)
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ─── PARA ÇIKIŞI (TRANSFER) YAPMA MODALI ─── */}
      {isTransferFormOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-gray-900/60 backdrop-blur-sm transition-all animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden flex flex-col">
            <div className="px-6 py-4 bg-orange-500 border-b border-orange-600 flex justify-between items-center">
              <h3 className="text-lg font-bold text-white flex items-center gap-2">
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 17h8m0 0V9m0 8l-8-8-4 4-6-6" /></svg>
                Para Çıkışı / Aktarım
              </h3>
              <button onClick={() => setIsTransferFormOpen(false)} className="text-orange-100 hover:text-white transition-colors">
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>
            
            <form onSubmit={handleTransferSubmit} className="p-6 space-y-4">
              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase">Hangi Kasadan/Hesaptan Çıkacak?</label>
                <select 
                  value={transferFormData.source_account} 
                  onChange={(e) => setTransferFormData(prev => ({ ...prev, source_account: e.target.value }))}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                >
                  <option value="Nakit">Nakit Kasa</option>
                  <option value="Havale (Tahir Çam)">Tahir Çam Hesabı</option>
                  <option value="Havale (Bekir Çam)">Bekir Çam Hesabı</option>
                  <option value="Havale (Tarık Çam)">Tarık Çam Hesabı</option>
                  <option value="Havale (Halil Çam)">Halil Çam Hesabı</option>
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase">Çıkan Tutar (₺)</label>
                <input 
                  type="number" 
                  required
                  value={transferFormData.amount} 
                  onChange={(e) => setTransferFormData(prev => ({ ...prev, amount: e.target.value }))}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm font-black text-orange-600 bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  placeholder="Örn: 500000"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase">İşlem Tarihi</label>
                <input 
                  type="date" 
                  required
                  value={transferFormData.transfer_date} 
                  onChange={(e) => setTransferFormData(prev => ({ ...prev, transfer_date: e.target.value }))}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-gray-500 uppercase">Açıklama (Nereye / Ne İçin)</label>
                <textarea 
                  required
                  rows={2}
                  value={transferFormData.description} 
                  onChange={(e) => setTransferFormData(prev => ({ ...prev, description: e.target.value }))}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm bg-gray-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20 resize-none"
                  placeholder="Örn: Döviz alımı için bankaya yatırıldı..."
                />
              </div>

              <div className="pt-2 flex justify-end gap-2 border-t border-gray-100">
                <button type="button" onClick={() => setIsTransferFormOpen(false)} className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-bold rounded-lg transition-colors">İptal</button>
                <button type="submit" className="px-4 py-2 bg-orange-500 hover:bg-orange-600 text-white text-xs font-bold rounded-lg shadow-md transition-colors">Kaydet ve Kasadan Düş</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── PARA ÇIKIŞI / TRANSFER GEÇMİŞİ MODALI ─── */}
      {isTransferHistoryOpen && (
        <div className="fixed inset-0 z-[55] flex items-center justify-center p-4 bg-gray-900/60 backdrop-blur-sm transition-all animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl overflow-hidden flex flex-col max-h-[85vh]">
            <div className="px-6 py-5 border-b border-gray-200 bg-gray-50 flex justify-between items-center">
              <div>
                <h3 className="text-lg font-bold text-gray-900 flex items-center gap-2">
                  <svg className="w-5 h-5 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                  Para Çıkışı (Transfer) Geçmişi
                </h3>
                <p className="text-xs text-gray-500 mt-1">Kasadan veya hesaplardan yapılan harcamalar ve dışarı aktarımlar.</p>
              </div>
              <button onClick={() => setIsTransferHistoryOpen(false)} className="text-gray-400 hover:text-red-500 transition-colors p-1">
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>

            <div className="overflow-y-auto flex-1 bg-white">
              {transferHistoryList.length === 0 ? (
                <div className="py-16 text-center text-gray-400">Şu ana kadar herhangi bir para çıkışı kaydedilmemiş.</div>
              ) : (
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-100/50 border-b border-gray-200 text-gray-500 text-[11px] font-black uppercase tracking-wider sticky top-0">
                      <th className="py-4 px-6">İşlem Tarihi</th>
                      <th className="py-4 px-6">Çıkış Yapılan Kasa/Hesap</th>
                      <th className="py-4 px-6 w-1/3">Açıklama</th>
                      <th className="py-4 px-6 text-right">Tutar</th>
                      <th className="py-4 px-6 text-center">İşlem</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 text-sm">
                    {transferHistoryList.map((item, idx) => (
                      <tr key={idx} className="hover:bg-gray-50 transition-colors">
                        <td className="py-3 px-6 text-gray-600 font-medium">{new Date(item.transfer_date).toLocaleDateString('tr-TR')}</td>
                        <td className="py-3 px-6">
                          <span className="inline-flex items-center gap-1.5 px-2 py-1 rounded text-[11px] font-bold bg-gray-100 text-gray-700 border border-gray-200">
                            {item.source_account}
                          </span>
                        </td>
                        <td className="py-3 px-6 text-gray-600 text-xs">{item.description}</td>
                        <td className="py-3 px-6 text-right">
                          <span className="font-black text-orange-600 text-base">-₺{Number(item.amount).toLocaleString('tr-TR')}</span>
                        </td>
                        <td className="py-3 px-6 text-center">
                          <button 
                            onClick={() => handleDeleteTransfer(item.id)}
                            className="text-red-500 hover:text-red-700 bg-red-50 hover:bg-red-100 px-2.5 py-1.5 rounded-lg text-[10px] font-bold transition-colors"
                            title="İşlemi Geri Al"
                          >
                            İptal Et
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

export default App;