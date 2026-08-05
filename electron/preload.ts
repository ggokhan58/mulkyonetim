import { contextBridge, ipcRenderer } from 'electron'

// React'in kullanacağı API köprüsü
const api = {
  fetchProperties: () => ipcRenderer.invoke('get-properties'),
  saveProperty: (propertyData: any) => ipcRenderer.invoke('add-property', propertyData),
  updateProperty: (propertyData: any) => ipcRenderer.invoke('update-property', propertyData),
  deleteProperty: (id: number) => ipcRenderer.invoke('delete-property', id),
  deletePayment: (id: number) => ipcRenderer.invoke('delete-payment', id),
  
  // ─── ÖDEME FONKSİYONLARI ───
  fetchPayments: (propertyId: number) => ipcRenderer.invoke('get-payments', propertyId),
  savePayment: (paymentData: any) => ipcRenderer.invoke('save-payment', paymentData),

  // ─── PDF FONKSİYONLARI ───
  selectPdf: () => ipcRenderer.invoke('select-pdf'),
  openPdf: (path: string) => ipcRenderer.invoke('open-pdf', path),

  getFinanceSummary: () => ipcRenderer.invoke('get-finance-summary'),
  exportDatabase: () => ipcRenderer.invoke('export-database'),
  
  // ─── HESAP DETAYLARI (DAİRELER) İÇİN YENİ EKLENEN KÖPRÜ ───
  getPaymentsByMethod: (method: string) => ipcRenderer.invoke('get-payments-by-method', method),

  // ─── PARA ÇIKIŞI (TRANSFER) KÖPRÜLERİ ───
  addTransfer: (data: any) => ipcRenderer.invoke('add-transfer', data),
  getTransfers: () => ipcRenderer.invoke('get-transfers'),
  deleteTransfer: (id: number) => ipcRenderer.invoke('delete-transfer', id)
}

// Güvenlik izolasyonu kontrolü ile köprüleri dışa aktarma
if (process.contextIsolated) {
  try {
    // 1. Mevcut API köprümüz
    contextBridge.exposeInMainWorld('api', api)
    
    // 2. EKSİK OLAN: ipcRenderer köprümüzü ekliyoruz
    contextBridge.exposeInMainWorld('ipcRenderer', {
      on: (channel: string, listener: (...args: any[]) => void) => {
        ipcRenderer.on(channel, listener)
      },
      send: (channel: string, ...args: any[]) => {
        ipcRenderer.send(channel, ...args)
      },
      invoke: (channel: string, ...args: any[]) => {
        return ipcRenderer.invoke(channel, ...args)
      }
    })
  } catch (error) {
    console.error(error)
  }
} else {
  // @ts-ignore
  window.api = api
  // @ts-ignore
  window.ipcRenderer = ipcRenderer
}