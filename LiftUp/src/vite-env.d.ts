/// <reference types="vite/client" />

export {};

declare global {
  interface Window {
    ipcRenderer: import('electron').IpcRenderer
    ipcMain: import('electron').IpcMain
  }
}