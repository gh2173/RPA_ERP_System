const { contextBridge, ipcRenderer } = require('electron');

// 렌더러 프로세스에 API 노출
contextBridge.exposeInMainWorld('electronAPI', {
  // 인증 정보 설정을 위한 API
  setCredentials: (credentials) => ipcRenderer.invoke('set-credentials', credentials),
  
  // EZ-Voucher 실행을 위한 API
  runEZVoucher: () => ipcRenderer.invoke('run-ezvoucher'),
  
  // 매입송장 처리를 위한 API
  processInvoice: () => ipcRenderer.invoke('process-invoice'),
  
  // A열 값 설정을 위한 API
  setValueA: (valueA) => ipcRenderer.invoke('set-value-a', valueA),
  
  // 다중 A열 값 처리를 위한 API
  processMultipleValueA: (valueArray) => ipcRenderer.invoke('process-multiple-value-a', valueArray),
  
  // 다중모드 진행 상황 업데이트 리스너
  onMultipleModeProgress: (callback) => ipcRenderer.on('multiple-mode-progress', (_, data) => callback(data)),
  
  // 스크린 캡처를 위한 API
  captureFullPage: () => ipcRenderer.invoke('capture-full-page'),
  
  // 기존 API들
  runRPA: () => ipcRenderer.invoke('run-rpa'),
  runTask: (taskName) => ipcRenderer.invoke('run-task', taskName),
  onTaskStatusUpdate: (callback) => ipcRenderer.on('task-status-update', (_, data) => callback(data)),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  processSelectedFiles: (startNumber, endNumber) => ipcRenderer.invoke('process-selected-files', startNumber, endNumber),
  processSingleFile: (fileNumber) => ipcRenderer.invoke('process-single-file', fileNumber)
});

// 호환성을 위해 electron 네임스페이스도 노출
contextBridge.exposeInMainWorld('electron', {
  ipcRenderer: {
    send: (channel, data) => ipcRenderer.send(channel, data),
    invoke: (channel, data) => ipcRenderer.invoke(channel, data),
    on: (channel, callback) => ipcRenderer.on(channel, callback),
    removeAllListeners: (channel) => ipcRenderer.removeAllListeners(channel)
  }
});