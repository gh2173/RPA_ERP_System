const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const ezVoucher = require('./EZVoucher.js');
const ezVoucher2 = require('./EZVoucher2.js');
const puppeteer = require('puppeteer');
const { dialog } = require('electron');

let mainWindow;
let browser = null;
let credentials = {
  username: '',
  password: ''
};

// 수정 후:
async function launchBrowser() {
  if (!browser) {
    browser = await puppeteer.ㅇlaunch({
      headless: false,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--start-maximized'
      ]
    });
  }
  return browser;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    },
    icon: path.join(__dirname, 'assets/app-icon.ico')
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// RPA 상태 업데이트를 렌더러에 전달하는 함수
function forwardStatusUpdate(data) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('task-status-update', data);
  }
}

// 이벤트 리스너 등록
ezVoucher.onStatusUpdate(forwardStatusUpdate);


ipcMain.handle('run-rpa', async () => {
  try {
    // 필요한 경우 브라우저 초기화
    await launchBrowser();
    
    let result = await ezVoucher.runAllTasks(credentials);
    
    return { 
      success: true, 
      data: {
        message: result.message || '작업이 완료되었습니다. 브라우저 창이 자동으로 닫혔습니다.',
        completedAt: result.completedAt || new Date().toISOString(),
        duration: result.duration || '2분',
        browserClosed: true
      }
    };
  } catch (error) {
    console.error('RPA 실행 오류:', error);
    return { success: false, error: error.message };
  }
});

ipcMain.handle('run-task', async (event, taskName) => {
  try {
    // 필요한 경우 브라우저 초기화
    await launchBrowser();
    
    let result = await ezVoucher.runTask(taskName, credentials);
    
    return { 
      success: true, 
      data: {
        message: result.message || '작업이 완료되었습니다. 브라우저 창이 자동으로 닫혔습니다.',
        completedAt: result.completedAt || new Date().toISOString(),
        duration: result.duration || '2분',
        browserClosed: true
      }
    };
  } catch (error) {
    console.error('작업 실행 오류:', error);
    return { success: false, error: error.message };
  }
});

// EZ-Voucher 실행 핸들러 추가
ipcMain.handle('run-ezvoucher', async () => {
  try {
    // ezVoucher 모듈을 사용하여 작업 실행
    const result = await ezVoucher.runAllTasks();
    return result;
  } catch (error) {
    console.error('EZ-Voucher 실행 오류:', error);
    return { success: false, error: error.message };
  }
});

// 파일 범위 처리 핸들러
ipcMain.on('start-file-range-processing', async (event, { userId, userPw, startFileNumber, endFileNumber }) => {
  try {
    console.log(`파일 범위 처리 시작: ${startFileNumber}-${endFileNumber}`);
    
    // 자격 증명 업데이트
    credentials.username = userId;
    credentials.password = userPw;    // 파일 범위 처리 실행
    const result = await ezVoucher.processFileRange(startFileNumber, endFileNumber);
    
    console.log('파일 범위 처리 결과:', result); // 디버깅용
    
    // 결과를 렌더러 프로세스로 전송
    event.reply('file-range-processing-result', {
      success: true,
      message: result.message || `파일 ${startFileNumber}-${endFileNumber} 처리 완료`,
      successCount: result.successCount || 0,
      failCount: result.failCount || 0,
      startFileNumber,
      endFileNumber,
      completedAt: result.completedAt || new Date().toISOString()
    });  } catch (error) {
    console.error('파일 범위 처리 오류:', error);
    event.reply('file-range-processing-result', {
      success: false,
      error: error.message,
      successCount: 0,
      failCount: endFileNumber - startFileNumber + 1,
      startFileNumber,
      endFileNumber
    });
  }
});

// 단일 파일 처리 핸들러
ipcMain.on('start-single-file-processing', async (event, { userId, userPw, fileNumber }) => {
  try {
    console.log(`단일 파일 처리 시작: ${fileNumber}`);
    
    // 자격 증명 업데이트
    credentials.username = userId;
    credentials.password = userPw;    // 단일 파일 처리 실행
    const result = await ezVoucher.processSingleFile(fileNumber);
    
    console.log('단일 파일 처리 결과:', result); // 디버깅용
    
    // 결과를 렌더러 프로세스로 전송
    event.reply('single-file-processing-result', {
      success: true,
      message: result.message || `파일 ${fileNumber} 처리 완료`,
      fileNumber,
      successCount: result.successCount || 1,
      failCount: result.failCount || 0,
      completedAt: result.completedAt || new Date().toISOString()
    });  } catch (error) {
    console.error('단일 파일 처리 오류:', error);
    event.reply('single-file-processing-result', {
      success: false,
      error: error.message,
      fileNumber,
      successCount: 0,
      failCount: 1
    });
  }
});

// 크레덴셜 설정 핸들러
ipcMain.handle('set-credentials', async (event, creds) => {
  try {
    credentials.username = creds.username;
    credentials.password = creds.password;
    
    // EZVoucher에 크레덴셜 설정
    if (ezVoucher && ezVoucher.setCredentials) {
      ezVoucher.setCredentials(creds.username, creds.password);
    }
    
    // EZVoucher2에 크레덴셜 설정
    if (ezVoucher2 && ezVoucher2.setCredentials) {
      ezVoucher2.setCredentials(creds.username, creds.password);
    }
    
    return { success: true };
  } catch (error) {
    console.error('크레덴셜 설정 오류:', error);
    return { success: false, error: error.message };
  }
});

// A열 값 설정 핸들러
ipcMain.handle('set-value-a', async (event, valueA) => {
  try {
    console.log(`[MAIN] A열 값 설정 요청: ${valueA} (타입: ${typeof valueA})`);
    
    if (ezVoucher2 && ezVoucher2.setValueA) {
      console.log('[MAIN] ezVoucher2.setValueA 함수 호출 중...');
      ezVoucher2.setValueA(valueA);
      console.log(`[MAIN] A열 값 설정 완료: ${valueA}`);
      return { success: true, message: `A열 값이 ${valueA}로 설정되었습니다.` };
    } else {
      console.error('[MAIN] ezVoucher2 또는 setValueA 함수를 찾을 수 없음');
      console.error('[MAIN] ezVoucher2:', !!ezVoucher2);
      console.error('[MAIN] ezVoucher2.setValueA:', !!ezVoucher2?.setValueA);
      throw new Error('setValueA 함수를 찾을 수 없습니다.');
    }
  } catch (error) {
    console.error('[MAIN] A열 값 설정 오류:', error);
    return { success: false, error: error.message };
  }
});

// 다중 A열 값 처리 핸들러
ipcMain.handle('process-multiple-value-a', async (event, valueArray) => {
  try {
    console.log(`다중 A열 값 처리 시작: ${valueArray}`);
    
    // 크레덴셜이 설정되어 있는지 확인
    if (!credentials.username || !credentials.password) {
      throw new Error('로그인 정보가 설정되지 않았습니다.');
    }
    
    if (ezVoucher2 && ezVoucher2.processMultipleValueA) {
      // 진행 상황 업데이트 콜백 함수 생성
      const sendProgressUpdate = (groupNumber, currentStep, completedSteps, error) => {
        if (mainWindow) {
          mainWindow.webContents.send('multiple-mode-progress', {
            groupNumber,
            currentStep,
            completedSteps,
            error
          });
        }
      };
      
      const result = await ezVoucher2.processMultipleValueA(valueArray, credentials, sendProgressUpdate);
      console.log('다중 A열 값 처리 완료:', result);
      
      // 다중모드임을 표시하는 플래그 추가
      if (result.success) {
        result.isMultipleMode = true;
      }
      
      return result;
    } else {
      throw new Error('processMultipleValueA 함수를 찾을 수 없습니다.');
    }
  } catch (error) {
    console.error('다중 A열 값 처리 오류:', error);
    return { success: false, error: error.message };
  }
});

// 매입송장 처리 핸들러
ipcMain.handle('process-invoice', async () => {
  try {
    console.log('매입송장 처리 시작...');
    
    // 크레덴셜이 설정되어 있는지 확인
    if (!credentials.username || !credentials.password) {
      throw new Error('로그인 정보가 설정되지 않았습니다.');
    }
    
    // EZVoucher2의 매입송장 처리 실행
    const result = await ezVoucher2.processInvoice(credentials);
    console.log('매입송장 처리 완료:', result);
    
    return result;
  } catch (error) {
    console.error('매입송장 처리 오류:', error);
    return { success: false, error: error.message };
  }
});

// 날짜 범위 설정 핸들러
ipcMain.handle('set-selected-date-range', async (event, dateRangeInfo) => {
  try {
    console.log(`[MAIN] 날짜 범위 설정 요청:`, dateRangeInfo);

    // EZVoucher2에 날짜 범위 설정
    if (ezVoucher2 && ezVoucher2.setSelectedDateRange) {
      ezVoucher2.setSelectedDateRange(dateRangeInfo);
      console.log(`[MAIN] 날짜 범위 설정 완료: ${dateRangeInfo.year}년 ${dateRangeInfo.month}월`);
      return { success: true, message: `날짜 범위가 설정되었습니다: ${dateRangeInfo.year}년 ${dateRangeInfo.month}월` };
    } else {
      throw new Error('EZVoucher2 모듈의 setSelectedDateRange 함수를 찾을 수 없습니다.');
    }
  } catch (error) {
    console.error('[MAIN] 날짜 범위 설정 오류:', error);
    return { success: false, error: error.message };
  }
});

// 전체 페이지 스크린샷 캡처 IPC 핸들러 (진행상황 로그 및 크기 제한 추가)
ipcMain.handle('capture-full-page', async () => {
  try {
    if (!mainWindow) throw new Error('메인 윈도우를 찾을 수 없습니다.');
    console.log('[캡처] 저장 다이얼로그 표시');
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const fileName = `full_page_capture_${timestamp}.png`;
    const result = await dialog.showSaveDialog(mainWindow, {
      title: '전체 페이지 스크롤 캡처 저장',
      defaultPath: fileName,
      filters: [
        { name: 'PNG 이미지', extensions: ['png'] },
        { name: '모든 파일', extensions: ['*'] }
      ]
    });
    if (result.canceled || !result.filePath) {
      console.log('[캡처] 사용자가 저장을 취소함');
      return { success: false, error: '사용자가 저장을 취소했습니다.' };
    }
    const filePath = result.filePath;

    console.log('[캡처] 문서 크기 정보 조회');
    const pageInfo = await mainWindow.webContents.executeJavaScript(`
      ({
        documentWidth: document.documentElement.scrollWidth,
        documentHeight: document.documentElement.scrollHeight,
        scrollX: window.scrollX,
        scrollY: window.scrollY
      })
    `);
    console.log(`[캡처] 문서 크기: ${pageInfo.documentWidth} x ${pageInfo.documentHeight}`);

    // 크기 제한 (예: 8000px 초과 시 경고)
    if (pageInfo.documentHeight > 8000 || pageInfo.documentWidth > 8000) {
      return { success: false, error: '캡처할 페이지가 너무 큽니다. (최대 8000px 제한)' };
    }

    // 스크롤바 숨기기
    await mainWindow.webContents.executeJavaScript(`
      document.documentElement.style.overflow = 'hidden';
    `);
    // 페이지 맨 위로 스크롤
    await mainWindow.webContents.executeJavaScript('window.scrollTo(0, 0)');
    await new Promise(resolve => setTimeout(resolve, 800));

    // 윈도우 크기를 문서 크기에 맞게 조정 (최대 제한 있음)
    const originalBounds = mainWindow.getBounds();
    const newHeight = Math.min(pageInfo.documentHeight + 200, 8000);
    const newWidth = Math.min(pageInfo.documentWidth + 50, 2000);
    mainWindow.setBounds({ ...originalBounds, width: newWidth, height: newHeight });
    await new Promise(resolve => setTimeout(resolve, 1500));

    // 캡처 수행
    console.log('[캡처] capturePage 실행');
    const image = await mainWindow.webContents.capturePage();

    // 원래 윈도우 크기 복원
    mainWindow.setBounds(originalBounds);
    await new Promise(resolve => setTimeout(resolve, 500));

    // 스크롤바 복원 및 원래 스크롤 위치 복원
    await mainWindow.webContents.executeJavaScript(`
      document.documentElement.style.overflow = '';
      window.scrollTo(${pageInfo.scrollX}, ${pageInfo.scrollY});
    `);

    // 이미지 저장
    const fs = require('fs');
    const buffer = image.toPNG();
    fs.writeFileSync(filePath, buffer);
    console.log(`[캡처] 저장 완료: ${filePath}`);

    return {
      success: true,
      path: filePath,
      message: '실제 열려있는 창의 전체 페이지 스크린샷이 저장되었습니다.',
      dimensions: {
        width: pageInfo.documentWidth,
        height: pageInfo.documentHeight
      }
    };
  } catch (error) {
    console.error('[캡처] 오류:', error);
    try {
      await mainWindow.webContents.executeJavaScript(`
        document.documentElement.style.overflow = '';
      `);
    } catch (e) {}
    return { success: false, error: error.message };
  }
});

// 앱 종료 시 브라우저 정리
app.on('before-quit', async () => {
  if (browser) {
    try {
      await browser.close();
    } catch (error) {
      console.error('브라우저 종료 오류:', error);
    }
  }
});