/// <reference types="vite-plugin-electron/electron-env" />

declare namespace NodeJS {
  interface ProcessEnv {
    /**
     * The built directory structure
     *
     * ```tree
     * ├─┬─┬ dist
     * │ │ └── index.html
     * │ │
     * │ ├─┬ dist-electron
     * │ │ ├── main.js
     * │ │ └── preload.js
     * │
     * ```
     */
    APP_ROOT: string
    /** /dist/ or /public/ */
    VITE_PUBLIC: string
  }
}

// Used in Renderer process, expose in `preload.ts`
interface Window {
  api: {
    getAppVersion: () => Promise<string>;
    checkForUpdates: () => Promise<void>;
    fetchProperties: () => Promise<any[]>;
    saveProperty: (propertyData: any) => Promise<any>;
    updateProperty: (propertyData: any) => Promise<any>;
    deleteProperty: (id: number) => Promise<boolean>;
    deletePayment: (id: number) => Promise<boolean>;
    fetchPayments: (propertyId: number) => Promise<any[]>;
    savePayment: (paymentData: any) => Promise<any>;
    selectPdf: () => Promise<string[] | null>;
    openPdf: (path: string) => Promise<boolean>;
    getFinanceSummary: () => Promise<any>;
    exportDatabase: () => Promise<{ success: boolean; path?: string } | null>;
    getPaymentsByMethod: (method: string) => Promise<any[]>;
    addTransfer: (data: any) => Promise<any>;
    getTransfers: () => Promise<any[]>;
    deleteTransfer: (id: number) => Promise<boolean>;
    exportTransactionsExcel: () => Promise<{ canceled?: boolean; success?: boolean; path?: string; error?: string }>;
  };
  ipcRenderer: import('electron').IpcRenderer
}
