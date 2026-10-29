/**
 * Web Digital · PDF Tools — Preload
 * Ponte segura entre o processo principal e a interface.
 * Expõe apenas o necessário via contextBridge (contextIsolation ativo).
 *
 * Desenvolvido por: Murillo Matos · Powered by: Web Digital (https://webdigital.app)
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    checkOffice: () => ipcRenderer.invoke('check-office'),
    selectFiles: () => ipcRenderer.invoke('select-files'),
    selectPdf: () => ipcRenderer.invoke('select-pdf'),
    saveFile: (defaultName) => ipcRenderer.invoke('save-file', defaultName),
    convertOfficeToPdf: (args) => ipcRenderer.invoke('convert-office-to-pdf', args),
    readPdfFile: (filePath) => ipcRenderer.invoke('read-pdf-file', filePath),
    savePdf: (args) => ipcRenderer.invoke('save-pdf', args),
    selectFolder: () => ipcRenderer.invoke('select-folder'),
    savePdfInFolder: (args) => ipcRenderer.invoke('save-pdf-in-folder', args),
    deleteTempFile: (filePath) => ipcRenderer.invoke('delete-temp-file', filePath),
    getTempFolder: () => ipcRenderer.invoke('get-temp-folder'),
    openFile: (filePath) => ipcRenderer.invoke('open-file', filePath),
    openFolder: (folderPath) => ipcRenderer.invoke('open-folder', folderPath),
    openExternal: (url) => ipcRenderer.invoke('open-external', url),
    getVersion: () => ipcRenderer.invoke('get-version'),
    // Atualização automática: status chega do processo principal
    onUpdateStatus: (callback) => ipcRenderer.on('update-status', (event, data) => callback(data)),
    installUpdate: () => ipcRenderer.invoke('install-update')
    // file.path é acessado diretamente no renderer (Electron 28 ainda expõe esta propriedade)
});
