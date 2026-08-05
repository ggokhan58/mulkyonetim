"use strict";
const electron = require("electron");
const api = {
  fetchProperties: () => electron.ipcRenderer.invoke("get-properties"),
  saveProperty: (propertyData) => electron.ipcRenderer.invoke("add-property", propertyData),
  updateProperty: (propertyData) => electron.ipcRenderer.invoke("update-property", propertyData)
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
