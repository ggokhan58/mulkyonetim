"use strict";
const electron = require("electron");
const api = {
  fetchProperties: () => electron.ipcRenderer.invoke("get-properties"),
  saveProperty: (propertyData) => electron.ipcRenderer.invoke("add-property", propertyData),
  updateProperty: (propertyData) => electron.ipcRenderer.invoke("update-property", propertyData),
  deleteProperty: (id) => electron.ipcRenderer.invoke("delete-property", id),
  deletePayment: (id) => electron.ipcRenderer.invoke("delete-payment", id),
  // ─── ÖDEME FONKSİYONLARI ───
  fetchPayments: (propertyId) => electron.ipcRenderer.invoke("get-payments", propertyId),
  savePayment: (paymentData) => electron.ipcRenderer.invoke("save-payment", paymentData),
  // ─── PDF FONKSİYONLARI ───
  selectPdf: () => electron.ipcRenderer.invoke("select-pdf"),
  openPdf: (path) => electron.ipcRenderer.invoke("open-pdf", path),
  getFinanceSummary: () => electron.ipcRenderer.invoke("get-finance-summary"),
  exportDatabase: () => electron.ipcRenderer.invoke("export-database"),
  // ─── HESAP DETAYLARI (DAİRELER) İÇİN YENİ EKLENEN KÖPRÜ ───
  getPaymentsByMethod: (method) => electron.ipcRenderer.invoke("get-payments-by-method", method),
  // ─── PARA ÇIKIŞI (TRANSFER) KÖPRÜLERİ ───
  addTransfer: (data) => electron.ipcRenderer.invoke("add-transfer", data),
  getTransfers: () => electron.ipcRenderer.invoke("get-transfers"),
  deleteTransfer: (id) => electron.ipcRenderer.invoke("delete-transfer", id)
};
if (process.contextIsolated) {
  try {
    electron.contextBridge.exposeInMainWorld("api", api);
    electron.contextBridge.exposeInMainWorld("ipcRenderer", {
      on: (channel, listener) => {
        electron.ipcRenderer.on(channel, listener);
      },
      send: (channel, ...args) => {
        electron.ipcRenderer.send(channel, ...args);
      },
      invoke: (channel, ...args) => {
        return electron.ipcRenderer.invoke(channel, ...args);
      }
    });
  } catch (error) {
    console.error(error);
  }
} else {
  window.api = api;
  window.ipcRenderer = electron.ipcRenderer;
}
