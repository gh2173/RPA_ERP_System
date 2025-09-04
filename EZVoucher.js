const puppeteer = require('puppeteer');
const puppeteerExtra = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const winston = require('winston');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// const { ipcMain } = require('electron');
const { ipcMain, dialog } = require('electron');

// 
let folderPath = null;

// // 폴더 경로 설정(TESTPath - 테스트용)
// const folderPath = 'C:\\Users\\nepes\\OneDrive - 네패스\\바탕 화면\\PROJECT_GIT\\Electron_Test\\my-electron-app\\Tax_Invoice_app\\1.14 엑셀전표업로드';

// 폴더 선택 핸들러 추가
ipcMain.handle('select-folder', async () => {
  try {
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory'],
      title: '엑셀 파일이 있는 폴더를 선택하세요'
    });
    
    if (!result.canceled && result.filePaths.length > 0) {
      folderPath = result.filePaths[0];
      logger.info(`폴더 경로가 설정되었습니다: ${folderPath}`);
      return { success: true, path: folderPath };
    }
    
    return { success: false, message: '폴더 선택이 취소되었습니다.' };
  } catch (error) {
    logger.error(`폴더 선택 중 오류 발생: ${error.message}`);
    return { success: false, message: error.message };
  }
});

// 선택된 범위의 파일들을 처리하는 IPC 핸들러
ipcMain.handle('process-selected-files', async (event, startNumber, endNumber) => {
  try {
    const credentials = getCredentials();
    if (!credentials.username || !credentials.password) {
      throw new Error('로그인 정보가 설정되지 않았습니다. 먼저 로그인을 해주세요.');
    }
    
    const result = await processSelectedFiles(credentials, startNumber, endNumber);
    return result;
  } catch (error) {
    logger.error(`선택된 파일들 처리 중 오류: ${error.message}`);
    return { success: false, error: error.message };
  }
});

// 단일 파일을 처리하는 IPC 핸들러
ipcMain.handle('process-single-file', async (event, fileNumber) => {
  try {
    const credentials = getCredentials();
    if (!credentials.username || !credentials.password) {
      throw new Error('로그인 정보가 설정되지 않았습니다. 먼저 로그인을 해주세요.');
    }
    
    const result = await processSingleFile(credentials, fileNumber);
    return result;
  } catch (error) {
    logger.error(`파일 ${fileNumber} 처리 중 오류: ${error.message}`);
    return { success: false, error: error.message };
  }
});


// 기본 대기 함수 (최적화를 위해 최소한만 사용)
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 성능 최적화를 위한 스마트 대기 시스템 - 강화된 fallback 메커니즘 포함
const smartWait = {
  // 요소가 나타날 때까지 최대 timeout까지 대기 (기본 5초로 단축)
  forElement: async (page, selector, timeout = 5000) => {
    try {
      await page.waitForSelector(selector, { visible: true, timeout });
      return true;
    } catch (error) {
      logger.warn(`요소 대기 시간 초과: ${selector} (${timeout}ms)`);
      return false;
    }
  },

  // 요소가 클릭 가능해질 때까지 대기
  forClickable: async (page, selector, timeout = 5000) => {
    try {
      await page.waitForSelector(selector, { visible: true, timeout });
      await page.waitForFunction(
        (sel) => {
          const el = document.querySelector(sel);
          return el && !el.disabled && el.offsetParent !== null;
        },
        { timeout: 3000 },
        selector
      );
      return true;
    } catch (error) {
      logger.warn(`클릭 가능 요소 대기 시간 초과: ${selector}`);
      return false;
    }
  },

  // 네트워크 활동이 완료될 때까지 대기 (최적화된 시간)
  forNetworkIdle: async (page, timeout = 10000) => {
    try {
      await page.waitForLoadState?.('networkidle', { timeout });
      return true;
    } catch (error) {
      // Puppeteer에서는 waitForLoadState가 없으므로 대안 사용
      try {
        await page.waitForFunction(
          () => document.readyState === 'complete',
          { timeout: timeout / 2 }
        );
        await delay(500); // 최소 0.5초 대기로 단축
        return true;
      } catch (err) {
        logger.warn('네트워크 대기 시간 초과');
        return false;
      }
    }
  },

  // 텍스트 내용이 나타날 때까지 대기
  forText: async (page, text, timeout = 5000) => {
    try {
      await page.waitForFunction(
        (searchText) => document.body.innerText.includes(searchText),
        { timeout },
        text
      );
      return true;
    } catch (error) {
      logger.warn(`텍스트 대기 시간 초과: ${text}`);
      return false;
    }
  },

  // 여러 셀렉터 중 하나라도 나타날 때까지 대기
  forAnyElement: async (page, selectors, timeout = 5000) => {
    try {
      const result = await Promise.race(
        selectors.map(selector => 
          page.waitForSelector(selector, { visible: true, timeout })
            .then(() => selector)
            .catch(() => null)
        )
      );
      return result;
    } catch (error) {
      logger.warn('여러 요소 대기 시간 초과');
      return null;
    }
  },

  // 페이지 로딩이 완료될 때까지 최적화된 대기
  forPageReady: async (page, timeout = 8000) => {
    try {
      await page.waitForFunction(
        () => {
          return document.readyState === 'complete' && 
                 !document.querySelector('.loading, .spinner, [data-loading="true"]');
        },
        { timeout }
      );
      await delay(300); // 최소 대기로 단축
      return true;
    } catch (error) {
      logger.warn('페이지 준비 대기 시간 초과');
      return false;
    }
  }
};

// 스텔스 플러그인 적용 (사이트가 봇을 감지하지 못하도록)
puppeteerExtra.use(StealthPlugin());

// 로거 설정
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.printf(({ timestamp, level, message }) => {
      return `[${timestamp}] ${level.toUpperCase()}: ${message}`;
    })
  ),
  transports: [
    new winston.transports.Console(),
    new winston.transports.File({ filename: 'rpa.log' })
  ]
});

/**
 * Dynamics 365 사이트 자동화 함수
 * @param {Object} credentials - 로그인 정보 객체 (username, password)
 */

// 파일명에서 괄호 안의 텍스트를 추출하는 함수
function extractTextFromParentheses(filePath) {
  try {
    const fileName = path.basename(filePath);
    logger.info(`파일명에서 괄호 안 텍스트 추출 중: ${fileName}`);
    
    // 괄호 패턴을 찾아서 안의 텍스트 추출 (한국어도 지원)
    const parenthesesMatch = fileName.match(/\(([^)]+)\)/);
    
    if (!parenthesesMatch) {
      throw new Error(`파일명에서 괄호 '()' 패턴을 찾을 수 없습니다: ${fileName}`);
    }
    
    const textInParentheses = parenthesesMatch[1];
    
    if (!textInParentheses || textInParentheses.trim() === '') {
      throw new Error(`파일명의 괄호 안에 유효한 텍스트가 없습니다: ${fileName}`);
    }
    
    const extractedText = textInParentheses.trim();
    logger.info(`파일명에서 추출된 괄호 안 텍스트: "${extractedText}"`);
    return extractedText;
  } catch (error) {
    logger.error(`파일명 괄호 텍스트 추출 오류: ${error.message}`);
    throw error;
  }
}

// 특정 번호로 시작하는 엑셀 파일 찾기 함수
function findExcelFileStartingWithNumber(folderPath, fileNumber = 1) {
  try {
    const filePrefix = `${fileNumber}.`;
    const files = fs.readdirSync(folderPath);
    const excelExtensions = ['.xlsx', '.xls', '.xlsm'];
    
    // 지정된 번호로 시작하고 확장자가 엑셀인 파일 필터링
    const matchingFiles = files.filter(file => {
      const isExcel = excelExtensions.some(ext => file.toLowerCase().endsWith(ext));
      return file.startsWith(filePrefix) && isExcel;
    });
    
    if (matchingFiles.length > 0) {
      // 찾은 파일 중 첫 번째 파일 반환
      return path.join(folderPath, matchingFiles[0]);
    }
    
    return null; // 해당하는 파일이 없는 경우
  } catch (error) {
    logger.error(`폴더 내 파일 검색 중 오류 발생: ${error.message}`);
    return null;
  }
}

async function navigateToDynamics365(credentials) {  // 폴더 경로가 설정되지 않은 경우 오류 반환
  if (!folderPath) {
    const errorMsg = '폴더 경로가 설정되지 않았습니다. 먼저 "폴더 지정" 버튼을 클릭하여 폴더를 선택해주세요.';
    logger.error(errorMsg);
    throw new Error(errorMsg);
  }

  logger.info('RPA 프로세스 시작: Dynamics 365 탐색');
  logger.info(`사용할 폴더 경로: ${folderPath}`);
  
  const browser = await puppeteerExtra.launch({
    headless: false,
    defaultViewport: null,
    args: [
      '--start-maximized',
      '--ignore-certificate-errors',
      '--ignore-certificate-errors-spki-list',
      '--allow-insecure-localhost',
      '--no-sandbox',
      '--disable-setuid-sandbox'
    ]
  });

  try {
    const page = await browser.newPage();
    
    // SSL 인증서 오류 처리
    await page.setBypassCSP(true);
    
    // 페이지 요청 인터셉트 설정 (SSL 오류 처리용)
    await page.setRequestInterception(true);
    page.on('request', request => {
      request.continue();
    });
    
    // 대화상자 처리 (인증서 경고 등)
    page.on('dialog', async dialog => {
      logger.info(`대화상자 감지: ${dialog.message()}`);
      await dialog.accept();
    });
    
    // 1. D365 페이지 접속 (재시도 로직 추가)
    logger.info('D365 페이지로 이동 중...');
    let pageLoadSuccess = false;
    let retryCount = 0;
    const maxRetries = 3;
    
    while (!pageLoadSuccess && retryCount < maxRetries) {
      try {
        retryCount++;
        logger.info(`D365 페이지 접속 시도 ${retryCount}/${maxRetries}`);
        
        await page.goto('https://d365.nepes.co.kr/namespaces/AXSF/?cmp=K02&mi=DefaultDashboard', {
          waitUntil: 'networkidle2',
          timeout: 60000 // 60초 타임아웃
        });
        
        pageLoadSuccess = true;
        logger.info('D365 페이지 로드 완료');
      } catch (networkError) {
        logger.error(`D365 페이지 접속 시도 ${retryCount} 실패: ${networkError.message}`);
        
        if (retryCount >= maxRetries) {
          const errorMsg = `네트워크 연결 실패: D365 사이트(https://d365.nepes.co.kr)에 접속할 수 없습니다. 인터넷 연결을 확인하거나 VPN이 필요할 수 있습니다.`;
          logger.error(errorMsg);
          throw new Error(errorMsg);
        }
        
        // 재시도 전 2초 대기 (성능 최적화: navigateToDynamics365 네트워크 재시도)
        logger.info('2초 후 재시도합니다...');
        await delay(2000);
      }
    }

    // 로그인 처리 (필요한 경우)
    if (await page.$('input[type="email"]') !== null || await page.$('#userNameInput') !== null) {
      logger.info('로그인 화면 감지됨, 로그인 시도 중...');
      await handleLogin(page, credentials);
    }    // 로그인 후 페이지가 완전히 로드될 때까지 스마트 대기 (성능 최적화)
    logger.info('로그인 후 페이지 로딩 확인 중...');
    const pageReady = await smartWait.forPageReady(page, 8000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }
    logger.info('페이지 로딩 확인 완료');

    // 2. 즐겨찾기 아이콘 클릭
    logger.info('즐겨찾기 아이콘 찾는 중...');
    
    // 즐겨찾기 아이콘 클릭 시도
    try {
      // 제공된 정확한 선택자로 찾기
      await page.waitForSelector('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"][data-dyn-image-type="Symbol"]', { 
        visible: true,
        timeout: 10000
      });
      
      await page.click('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"][data-dyn-image-type="Symbol"]');
      logger.info('정확한 선택자로 즐겨찾기 아이콘 클릭 성공');
    } catch (error) {
      logger.warn(`정확한 선택자로 즐겨찾기 아이콘을 찾지 못함: ${error.message}`);
      
      // 더 간단한 선택자로 시도
      try {
        await page.waitForSelector('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"]', { 
          visible: true,
          timeout: 5000
        });
        
        await page.click('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"]');
        logger.info('단순 선택자로 즐겨찾기 아이콘 클릭 성공');
      } catch (iconError) {
        logger.warn(`단순 선택자로도 즐겨찾기 아이콘을 찾지 못함: ${iconError.message}`);
        
        // 클래스 이름으로 시도
        try {
          await page.waitForSelector('.StarEmpty-symbol', { visible: true, timeout: 5000 });
          await page.click('.StarEmpty-symbol');
          logger.info('클래스 이름으로 즐겨찾기 아이콘 클릭 성공');
        } catch (classError) {
          logger.warn(`클래스 이름으로도 즐겨찾기 아이콘을 찾지 못함: ${classError.message}`);
          
          // JavaScript로 시도
          try {
            logger.info('JavaScript로 즐겨찾기 아이콘 찾기 시도...');
            
            await page.evaluate(() => {
              // 여러 방법으로 요소 찾기
              const spans = Array.from(document.querySelectorAll('span'));
              const favIcon = spans.find(span => 
                span.getAttribute('data-dyn-title') === '즐겨찾기' || 
                span.classList.contains('StarEmpty-symbol') ||
                (span.className && span.className.includes('StarEmpty'))
              );
              
              if (favIcon) {
                favIcon.click();
              } else {
                throw new Error('JavaScript로도 즐겨찾기 아이콘을 찾을 수 없음');
              }
            });
            
            logger.info('JavaScript로 즐겨찾기 아이콘 클릭 성공');
          } catch (jsError) {
            logger.error(`JavaScript로도 즐겨찾기 아이콘을 찾지 못함: ${jsError.message}`);
            throw new Error('즐겨찾기 아이콘을 찾을 수 없습니다');
          }
        }
      }
    }
      logger.info('즐겨찾기 아이콘 클릭 완료');
    
    // 클릭 후 메뉴가 표시될 때까지 스마트 대기 (성능 최적화)
    const menuVisible = await smartWait.forAnyElement(page, [
      'div.modulesPane-link.modulesFlyout-isFavorite',
      '.modulesPane-link',
      'a[data-dyn-title="엑셀 전표 업로드"]'
    ], 5000);
    
    if (!menuVisible) {
      logger.warn('즐겨찾기 메뉴 표시 확인 실패, 기본 1초 대기로 진행');
      await delay(1000);
    }
    logger.info('즐겨찾기 메뉴 로드 대기 완료');

    // 3. 메뉴에서 "엑셀 전표 업로드" 클릭
    logger.info('"엑셀 전표 업로드" 메뉴 아이템 찾는 중...');
    
    // 정확한 선택자로 시도
    try {
      // 제공된 정확한 선택자로 찾기
      const exactSelector = 'div.modulesPane-link.modulesFlyout-isFavorite[data-dyn-selected="false"][role="treeitem"] a.modulesPane-linkText[data-dyn-title="엑셀 전표 업로드"][role="link"]';
      await page.waitForSelector(exactSelector, { visible: true, timeout: 10000 });
      await page.click(exactSelector);
      logger.info('정확한 선택자로 "엑셀 전표 업로드" 메뉴 클릭 완료');
    } catch (error) {
      logger.warn(`정확한 선택자로 메뉴를 찾지 못함: ${error.message}`);
      
      // 원래 코드의 선택자로 시도
      try {
        const selector = 'div[data-dyn-title="엑셀 전표 업로드"], div.modulesPane-link a[data-dyn-title="엑셀 전표 업로드"], .modulesPane-link a.modulesPane-linkText[data-dyn-title="엑셀 전표 업로드"]';
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        await page.click(selector);
        logger.info('기본 선택자로 "엑셀 전표 업로드" 메뉴 클릭 완료');
      } catch (selectorError) {
        logger.warn(`기본 선택자로도 메뉴를 찾지 못함: ${selectorError.message}`);
        
        // 텍스트 기반으로 요소 찾기
        try {
          const menuItems = await page.$$('.modulesPane-link, .modulesFlyout-isFavorite');
          
          let found = false;
          for (const item of menuItems) {
            const text = await page.evaluate(el => el.textContent, item);
            if (text.includes('엑셀 전표 업로드')) {
              await item.click();
              found = true;
              logger.info('텍스트 검색으로 "엑셀 전표 업로드" 메뉴 클릭 완료');
              break;
            }
          }
          
          if (!found) {
            // 링크 텍스트로 찾기
            try {
              // 페이지에서 JavaScript 실행하여 요소 찾기
              await page.evaluate(() => {
                const links = Array.from(document.querySelectorAll('a.modulesPane-linkText, div.modulesPane-link a, a[role="link"]'));
                const targetLink = links.find(link => link.textContent.includes('엑셀 전표 업로드'));
                if (targetLink) {
                  targetLink.click();
                } else {
                  throw new Error('링크를 찾을 수 없음');
                }
              });
              logger.info('JavaScript 실행으로 "엑셀 전표 업로드" 메뉴 클릭 완료');
            } catch (jsError) {
              logger.error(`JavaScript로도 메뉴를 찾지 못함: ${jsError.message}`);
              throw new Error('메뉴 아이템을 찾을 수 없습니다: 엑셀 전표 업로드');
            }
          }
        } catch (textError) {
          logger.error(`텍스트 검색으로도 메뉴를 찾지 못함: ${textError.message}`);
          throw new Error('모든 방법으로 "엑셀 전표 업로드" 메뉴를 찾지 못했습니다');
        }
      }    }
    
    // 엑셀 전표 업로드 페이지 로드 스마트 대기 (성능 최적화)
    logger.info('엑셀 전표 업로드 페이지 로드 확인 중...');
    const uploadPageReady = await smartWait.forAnyElement(page, [
      '.lookupButton[title="오픈"]',
      '.lookupButton',
      'input[value="일반전표(ARK)"]'
    ], 8000);
    
    if (!uploadPageReady) {
      logger.warn('업로드 페이지 로딩 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }
    logger.info('엑셀 전표 업로드 페이지 로드 완료');

    // 추가 동작 1: lookupButton 클래스를 가진 요소 클릭
    logger.info('lookupButton 클래스 요소 찾는 중...');
    try {
      await page.waitForSelector('.lookupButton[title="오픈"]', { 
        visible: true, 
        timeout: 10000 
      });
      await page.click('.lookupButton[title="오픈"]');
      logger.info('lookupButton 클릭 성공');
    } catch (error) {
      logger.warn(`lookupButton을 찾지 못함: ${error.message}`);
      
      // JavaScript로 시도
      try {
        await page.evaluate(() => {
          const lookupButtons = Array.from(document.querySelectorAll('.lookupButton'));
          const button = lookupButtons.find(btn => 
            btn.getAttribute('title') === '오픈' || 
            btn.getAttribute('data-dyn-bind')?.includes('Input_LookupTooltip')
          );
          
          if (button) {
            button.click();
          } else {
            throw new Error('lookupButton을 찾을 수 없음');
          }
        });
        logger.info('JavaScript로 lookupButton 클릭 성공');
      } catch (jsError) {
        logger.error(`JavaScript로도 lookupButton을 찾지 못함: ${jsError.message}`);
        throw new Error('lookupButton을 찾을 수 없습니다');      }
    }

    // 팝업이 열릴 때까지 스마트 대기 (성능 최적화)
    const popupReady = await smartWait.forAnyElement(page, [
      'input[value="일반전표(ARK)"]',
      'input[title="일반전표(ARK)"]',
      '#SysGen_Name_125_0_0_input'
    ], 5000);
    
    if (!popupReady) {
      logger.warn('팝업 표시 확인 실패, 기본 1초 대기로 진행');
      await delay(1000);
    }
    logger.info('lookupButton 클릭 후 팝업 대기 완료');

    // 추가 동작 2: "일반전표(ARK)" 값을 가진 텍스트 입력 필드 클릭
    logger.info('"일반전표(ARK)" 텍스트 필드 찾는 중...');
    try {
      await page.waitForSelector('input[value="일반전표(ARK)"], input[title="일반전표(ARK)"]', { 
        visible: true, 
        timeout: 10000 
      });
      await page.click('input[value="일반전표(ARK)"], input[title="일반전표(ARK)"]');
      logger.info('"일반전표(ARK)" 텍스트 필드 클릭 성공');
    } catch (error) {
      logger.warn(`"일반전표(ARK)" 텍스트 필드를 찾지 못함: ${error.message}`);
      
      // ID로 시도
      try {
        await page.waitForSelector('#SysGen_Name_125_0_0_input', { visible: true, timeout: 5000 });
        await page.click('#SysGen_Name_125_0_0_input');
        logger.info('ID로 "일반전표(ARK)" 텍스트 필드 클릭 성공');
      } catch (idError) {
        logger.warn(`ID로도 텍스트 필드를 찾지 못함: ${idError.message}`);
        
        // JavaScript로 시도
        try {
          await page.evaluate(() => {
            const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
            const input = inputs.find(inp => 
              inp.value === '일반전표(ARK)' || 
              inp.title === '일반전표(ARK)'
            );
            
            if (input) {
              input.click();
            } else {
              throw new Error('텍스트 필드를 찾을 수 없음');
            }
          });
          logger.info('JavaScript로 "일반전표(ARK)" 텍스트 필드 클릭 성공');
        } catch (jsError) {
          logger.error(`JavaScript로도 텍스트 필드를 찾지 못함: ${jsError.message}`);
          throw new Error('"일반전표(ARK)" 텍스트 필드를 찾을 수 없습니다');
        }
      }
    }    await delay(1000); // 최소 대기로 단축 (성능 최적화)
    logger.info('텍스트 필드 클릭 후 대기 완료');    // 추가 동작 3: 특정 텍스트 박스에 파일명 괄호 안의 텍스트 입력 (필수 동작)
    logger.info('텍스트 박스 찾아 파일명 괄호 안의 텍스트 입력 중...');

    // 먼저 엑셀 파일 경로 찾기
    const excelFilePath = findExcelFileStartingWith1(folderPath);

    if (!excelFilePath) {
      const errorMsg = '1.로 시작하는 엑셀 파일을 찾지 못했습니다. RPA를 중단합니다.';
      logger.error(errorMsg);
      throw new Error(errorMsg);
    }

    let textToInput;
    try {
      // 파일명에서 괄호 안의 텍스트 추출 - 반드시 성공해야 함
      textToInput = extractTextFromParentheses(excelFilePath);
      
      if (!textToInput || textToInput.trim() === '') {
        throw new Error(`파일명 괄호 안에 유효한 텍스트가 없습니다. 파일: ${path.basename(excelFilePath)}`);
      }
      
      logger.info(`파일명 괄호 안의 텍스트 "${textToInput}" 추출 성공`);
    } catch (extractError) {
      const errorMsg = `파일명 괄호 텍스트 추출 실패: ${extractError.message}. 파일: ${path.basename(excelFilePath)}. RPA를 중단합니다.`;
      logger.error(errorMsg);
      throw new Error(errorMsg);
    }

    // 텍스트 입력 요소가 완전히 로드될 때까지 추가 대기
    logger.info('텍스트 입력 페이지 완전 로드 대기 중...');
    await delay(3000);

    // 텍스트 입력 - 강화된 5번 시도
    let inputSuccess = false;
    let lastError = null;

    // 첫 번째 시도: 메인 ID 선택자 (강화된 버전)
    if (!inputSuccess) {
      try {
        logger.info('첫 번째 시도: 메인 ID 선택자로 텍스트 입력 (강화된 대기)');
        
        // 페이지가 완전히 안정화될 때까지 대기
        await page.waitForFunction(
          () => document.readyState === 'complete',
          { timeout: 10000 }
        );
        
        // 여러 가능한 선택자들을 시도
        const possibleSelectors = [
          '#kpc_exceluploadforledgerjournal_2_FormStringControl_Txt_input',
          'input[id*="FormStringControl_Txt_input"]',
          'input[id*="kpc_exceluploadforledgerjournal"][id*="Txt_input"]',
          'input[role="textbox"]'
        ];
        
        let foundElement = null;
        for (const selector of possibleSelectors) {
          try {
            await page.waitForSelector(selector, { visible: true, timeout: 3000 });
            foundElement = selector;
            logger.info(`요소 찾음: ${selector}`);
            break;
          } catch (e) {
            logger.warn(`선택자 실패: ${selector}`);
          }
        }
        
        if (foundElement) {
          // 요소가 클릭 가능할 때까지 대기
          await page.waitForFunction(
            (sel) => {
              const element = document.querySelector(sel);
              return element && !element.disabled && element.offsetParent !== null;
            },
            { timeout: 5000 },
            foundElement
          );
          
          await page.click(foundElement);
          await delay(500); // 클릭 후 잠시 대기
          await page.type(foundElement, textToInput);
          logger.info(`첫 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
          inputSuccess = true;
        } else {
          throw new Error('모든 선택자에서 요소를 찾을 수 없음');
        }
      } catch (error) {
        lastError = error;
        logger.warn(`첫 번째 시도 실패: ${error.message}`);
      }
    }

    // 두 번째 시도: 클래스 선택자 (강화된 버전)
    if (!inputSuccess) {
      try {
        logger.info('두 번째 시도: 클래스 선택자로 텍스트 입력');
        
        const classSelectors = [
          'input.textbox.field.displayoption[role="textbox"]',
          'input.textbox.field.displayoption',
          'input.textbox[role="textbox"]',
          'input[class*="textbox"][class*="field"]'
        ];
        
        let foundElement = null;
        for (const selector of classSelectors) {
          try {
            await page.waitForSelector(selector, { visible: true, timeout: 3000 });
            foundElement = selector;
            logger.info(`클래스 요소 찾음: ${selector}`);
            break;
          } catch (e) {
            logger.warn(`클래스 선택자 실패: ${selector}`);
          }
        }
        
        if (foundElement) {
          await page.click(foundElement);
          await delay(500);
          await page.type(foundElement, textToInput);
          logger.info(`두 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
          inputSuccess = true;
        } else {
          throw new Error('모든 클래스 선택자에서 요소를 찾을 수 없음');
        }
        logger.info(`두 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
        inputSuccess = true;
      } catch (error) {
        lastError = error;
        logger.warn(`두 번째 시도 실패: ${error.message}`);
      }
    }

    // 세 번째 시도: JavaScript로 직접 조작
    if (!inputSuccess) {
      try {
        logger.info('세 번째 시도: JavaScript로 직접 텍스트 입력');
        await page.evaluate((text) => {
          const inputs = Array.from(document.querySelectorAll('input[role="textbox"]'));
          const input = inputs.find(inp => 
            inp.className.includes('textbox') || 
            inp.id?.includes('FormStringControl_Txt_input')
          );
          
          if (input) {
            input.focus();
            input.value = text;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          } else {
            throw new Error('텍스트 박스를 찾을 수 없음');
          }
        }, textToInput);
        logger.info(`세 번째 시도 성공: JavaScript로 텍스트 박스에 "${textToInput}" 입력 완료`);
        inputSuccess = true;
      } catch (error) {
        lastError = error;
        logger.warn(`세 번째 시도 실패: ${error.message}`);
      }
    }

    // 네 번째 시도: 모든 텍스트 입력 필드 검색
    if (!inputSuccess) {
      try {
        logger.info('네 번째 시도: 모든 텍스트 입력 필드에서 검색');
        await page.evaluate((text) => {
          const allInputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type]), textarea'));
          
          // 설명란으로 추정되는 입력 필드 찾기
          const descriptionInput = allInputs.find(inp => 
            inp.placeholder?.includes('설명') ||
            inp.title?.includes('설명') ||
            inp.getAttribute('aria-label')?.includes('설명') ||
            inp.id?.includes('Description') ||
            inp.id?.includes('Txt') ||
            inp.name?.includes('description')
          );
          
          if (descriptionInput) {
            descriptionInput.focus();
            descriptionInput.value = text;
            descriptionInput.dispatchEvent(new Event('input', { bubbles: true }));
            descriptionInput.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          } else {
            throw new Error('설명란으로 추정되는 입력 필드를 찾을 수 없음');
          }
        }, textToInput);
        logger.info(`네 번째 시도 성공: 설명란 입력 필드에 "${textToInput}" 입력 완료`);
        inputSuccess = true;
      } catch (error) {
        lastError = error;
        logger.warn(`네 번째 시도 실패: ${error.message}`);
      }
    }

    // 다섯 번째 시도: 가장 가능성 높은 입력 필드 선택
    if (!inputSuccess) {
      try {
        logger.info('다섯 번째 시도: 가장 가능성 높은 입력 필드 선택');
        await page.evaluate((text) => {
          const allInputs = Array.from(document.querySelectorAll('input, textarea'));
          
          // 보이는 입력 필드만 필터링
          const visibleInputs = allInputs.filter(inp => {
            const style = window.getComputedStyle(inp);
            return style.display !== 'none' && style.visibility !== 'hidden' && inp.offsetWidth > 0 && inp.offsetHeight > 0;
          });
          
          if (visibleInputs.length > 0) {
            // 가장 큰 입력 필드 선택 (설명란일 가능성이 높음)
            const largestInput = visibleInputs.reduce((prev, current) => {
              return (prev.offsetWidth * prev.offsetHeight) > (current.offsetWidth * current.offsetHeight) ? prev : current;
            });
            
            largestInput.focus();
            largestInput.value = text;
            largestInput.dispatchEvent(new Event('input', { bubbles: true }));
            largestInput.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          } else {
            throw new Error('보이는 입력 필드를 찾을 수 없음');
          }
        }, textToInput);
        logger.info(`다섯 번째 시도 성공: 가장 큰 입력 필드에 "${textToInput}" 입력 완료`);
        inputSuccess = true;
      } catch (error) {
        lastError = error;        logger.warn(`다섯 번째 시도 실패: ${error.message}`);      }    }    // 모든 시도 실패 시 RPA 작업 중단 (필수 동작)
    if (!inputSuccess) {
      const errorMsg = `🚨 중요: 추가 동작 3번 실패 - 괄호 안 텍스트 "${textToInput}" 입력에 모든 시도가 실패했습니다. 마지막 오류: ${lastError?.message || '알 수 없는 오류'}. RPA 작업을 중단합니다.`;
      logger.error(errorMsg);
      
      // 브라우저 닫기
      try {
        await browser.close();
      } catch (closeError) {
        logger.error(`브라우저 종료 중 오류: ${closeError.message}`);
      }
      
      // 오류 발생으로 작업 중단
      throw new Error(errorMsg);
    }

    await delay(2000);
    logger.info('텍스트 입력 후 대기 완료');

    // 추가 동작 4: "업로드" 버튼 클릭 (확인 버튼 대신)
    logger.info('"업로드" 버튼 찾는 중...');
    try {
      await page.waitForSelector('#kpc_exceluploadforledgerjournal_2_UploadButton_label', { 
        visible: true, 
        timeout: 10000 
      });
      await page.click('#kpc_exceluploadforledgerjournal_2_UploadButton_label');
      logger.info('ID로 "업로드" 버튼 클릭 성공');
    } catch (error) {
      logger.warn(`ID로 "업로드" 버튼을 찾지 못함: ${error.message}`);
      
      // 텍스트로 시도
      try {
        await page.waitForSelector('span.button-label:contains("업로드")', { 
          visible: true, 
          timeout: 5000 
        });
        await page.click('span.button-label:contains("업로드")');
        logger.info('텍스트로 "업로드" 버튼 클릭 성공');
      } catch (textError) {
        logger.warn(`텍스트로도 "업로드" 버튼을 찾지 못함: ${textError.message}`);
        
        // JavaScript로 시도
        try {
          await page.evaluate(() => {
            const spans = Array.from(document.querySelectorAll('span.button-label, span[id*="UploadButton_label"]'));
            const uploadButton = spans.find(span => span.textContent.trim() === '업로드');
            
            if (uploadButton) {
              uploadButton.click();
            } else {
              // 다른 방식으로 업로드 버튼 찾기
              const buttons = Array.from(document.querySelectorAll('button, div.button-container, [role="button"]'));
              const btn = buttons.find(b => {
                const text = b.textContent.trim();
                return text === '업로드' || text.includes('업로드');
              });
              
              if (btn) {
                btn.click();
              } else {
                throw new Error('업로드 버튼을 찾을 수 없음');
              }
            }
          });
          logger.info('JavaScript로 "업로드" 버튼 클릭 성공');
        } catch (jsError) {
          logger.error(`JavaScript로도 "업로드" 버튼을 찾지 못함: ${jsError.message}`);
          throw new Error('"업로드" 버튼을 찾을 수 없습니다');
        }
      }    }

    // 추가 동작 5: "Browse" 버튼 클릭 및 파일 선택
    // 업로드 버튼 클릭 후 Browse 버튼이 활성화될 때까지 스마트 대기 (성능 최적화)
    const browseButtonReady = await smartWait.forAnyElement(page, [
      '#Dialog_4_UploadBrowseButton',
      'button[name="UploadBrowseButton"]',
      'input[type="file"]'
    ], 5000);
    
    if (!browseButtonReady) {
      logger.warn('Browse 버튼 활성화 확인 실패, 기본 1초 대기로 진행');
      await delay(1000);
    }
    logger.info('"Browse" 버튼 찾는 중...');

    // '1.'로 시작하는 엑셀 파일 찾기 함수
    function findExcelFileStartingWith1(folderPath) {
      try {
        const files = fs.readdirSync(folderPath);
        const excelExtensions = ['.xlsx', '.xls', '.xlsm'];
        
        // '1.'로 시작하고 확장자가 엑셀인 파일 필터링
        const matchingFiles = files.filter(file => {
          const isExcel = excelExtensions.some(ext => file.toLowerCase().endsWith(ext));
          return file.startsWith('1.') && isExcel;
        });
        
        if (matchingFiles.length > 0) {
          // 찾은 파일 중 첫 번째 파일 반환
          return path.join(folderPath, matchingFiles[0]);
        }
        
        return null; // 해당하는 파일이 없는 경우
      } catch (error) {
        logger.error(`폴더 내 파일 검색 중 오류 발생: ${error.message}`);
        return null;
      }
    }

    try {
      // 먼저 파일 경로 확인
      const filePath = findExcelFileStartingWith1(folderPath);
      
      if (!filePath) {
        throw new Error('폴더에서 "1."로 시작하는 엑셀 파일을 찾을 수 없습니다.');
      }
      
      logger.info(`사용할 파일: ${path.basename(filePath)}`);
      
      // 방법 1: fileChooser를 사용하여 파일 선택
      try {
        // 파일 선택기가 열릴 때까지 대기하면서 Browse 버튼 클릭
        const [fileChooser] = await Promise.all([
          page.waitForFileChooser({ timeout: 10000 }),
          page.click('#Dialog_4_UploadBrowseButton, button[name="UploadBrowseButton"]')
        ]);
        
        // 찾은 파일 선택
        await fileChooser.accept([filePath]);
        logger.info(`fileChooser 방식으로 파일 선택 완료: ${path.basename(filePath)}`);
      } catch (chooserError) {
        logger.warn(`fileChooser 방식 실패: ${chooserError.message}`);
        
        // 방법 2: 파일 입력 필드를 직접 찾아 조작
        try {
          // 먼저 Brows 버튼 클릭 취소(이미 클릭했을 수 있기 때문에)
          await page.keyboard.press('Escape');
          await delay(1000);
          
          // 파일 입력 필드 찾기
          const fileInputSelector = 'input[type="file"]';
          const fileInput = await page.$(fileInputSelector);
          
          if (fileInput) {
            // 파일 입력 필드가 있으면 직접 파일 설정
            await fileInput.uploadFile(filePath);
            logger.info(`uploadFile 방식으로 파일 선택 완료: ${path.basename(filePath)}`);
          } else {
            // 파일 입력 필드가 없으면 다시 Browse 버튼 클릭 시도
            await page.click('#Dialog_4_UploadBrowseButton, button[name="UploadBrowseButton"]');
            
            // 사용자에게 파일 선택 안내
            await page.evaluate((fileName) => {
              alert(`자동 파일 선택에 실패했습니다. 파일 탐색기에서 "${fileName}" 파일을 수동으로 선택해주세요.`);
            }, path.basename(filePath));
            
            // 사용자가 파일을 선택할 때까지 대기 (30초)
            logger.info('사용자의 수동 파일 선택 대기 중... (30초)');
            await delay(30000);
          }
        } catch (inputError) {
          logger.error(`파일 입력 방식도 실패: ${inputError.message}`);
          
          // 최후의 방법: 사용자에게 안내
          await page.evaluate((message) => {
            alert(message);
          }, `자동 파일 선택에 실패했습니다. 파일 탐색기에서 "1."로 시작하는 엑셀 파일을 수동으로 선택해주세요.`);
          
          // 사용자가 파일을 선택할 때까지 대기 (30초)
          logger.info('사용자의 수동 파일 선택 대기 중... (30초)');
          await delay(30000);
        }
      }
      
      // 파일 선택 후 대기
      await delay(3000);
      logger.info('파일 선택 후 대기 완료');
      
      // 파일 선택 대화상자에서 "확인" 버튼이 필요한 경우 클릭
      try {
        const confirmButtonSelector = '#Dialog_4_OkButton, #SysOKButton, span.button-label:contains("확인"), span.button-label:contains("OK")';
        const confirmButton = await page.$(confirmButtonSelector);
        
        if (confirmButton) {
          await confirmButton.click();
          logger.info('파일 선택 대화상자의 "확인" 버튼 클릭 성공');
        }
      } catch (confirmError) {
        logger.warn(`확인 버튼 클릭 시도 중 오류: ${confirmError.message}`);
        logger.info('계속 진행합니다...');
      }
      
    } catch (browseError) {
      logger.error(`"Browse" 버튼 처리 오류: ${browseError.message}`);
      
      // 사용자에게 파일 선택 안내
      try {
        await page.evaluate(() => {
          alert('자동화 스크립트에 문제가 발생했습니다. 수동으로 "Browse" 버튼을 클릭하고 "1."로 시작하는 엑셀 파일을 선택해주세요.');
        });
        
        // 사용자가 작업을 완료할 때까지 대기
        logger.info('사용자의 수동 파일 선택 대기 중... (60초)');
        await delay(60000);
      } catch (alertError) {
        logger.error(`알림 표시 중 오류: ${alertError.message}`);
      }    }

    // 추가 동작 6: 파일 선택 후 최종 "확인" 버튼 클릭
    // 파일 선택 후 대화상자 확인 버튼이 활성화될 때까지 스마트 대기 (성능 최적화)
    const confirmButtonReady = await smartWait.forAnyElement(page, [
      '#Dialog_4_OkButton',
      'button[name="OkButton"]',
      'span.button-label:contains("확인")',
      'span.button-label:contains("OK")'
    ], 8000);
    
    if (!confirmButtonReady) {
      logger.warn('파일 선택 후 확인 버튼 활성화 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }
    logger.info('파일 선택 후 최종 "확인" 버튼 찾는 중...');

    try {
      // ID로 시도
      await page.waitForSelector('#Dialog_4_OkButton', { 
        visible: true, 
        timeout: 10000 
      });
      await page.click('#Dialog_4_OkButton');
      logger.info('ID로 최종 "확인" 버튼 클릭 성공');
    } catch (error) {
      logger.warn(`ID로 최종 "확인" 버튼을 찾지 못함: ${error.message}`);
      
      // 이름 속성으로 시도
      try {
        await page.waitForSelector('button[name="OkButton"]', { 
          visible: true, 
          timeout: 5000 
        });
        await page.click('button[name="OkButton"]');
        logger.info('name 속성으로 최종 "확인" 버튼 클릭 성공');
      } catch (nameError) {
        logger.warn(`name 속성으로도 최종 "확인" 버튼을 찾지 못함: ${nameError.message}`);
        
        // 레이블로 시도
        try {
          await page.waitForSelector('#Dialog_4_OkButton_label, span.button-label:contains("확인")', { 
            visible: true, 
            timeout: 5000 
          });
          await page.click('#Dialog_4_OkButton_label, span.button-label:contains("확인")');
          logger.info('레이블로 최종 "확인" 버튼 클릭 성공');
        } catch (labelError) {
          logger.warn(`레이블로도 최종 "확인" 버튼을 찾지 못함: ${labelError.message}`);
          
          // JavaScript로 시도
          try {
            await page.evaluate(() => {
              // 방법 1: ID로 찾기
              const okButton = document.querySelector('#Dialog_4_OkButton');
              if (okButton) {
                okButton.click();
                return;
              }
              
              // 방법 2: 레이블로 찾기
              const okLabel = document.querySelector('#Dialog_4_OkButton_label');
              if (okLabel) {
                okLabel.click();
                return;
              }
              
              // 방법 3: 버튼 텍스트로 찾기
              const buttons = Array.from(document.querySelectorAll('button, span.button-label'));
              const button = buttons.find(btn => btn.textContent.trim() === '확인');
              if (button) {
                button.click();
                return;
              }
              
              // 방법 4: 버튼 클래스와 속성으로 찾기
              const dynamicsButtons = Array.from(document.querySelectorAll('button.dynamicsButton.button-isDefault'));
              const defaultButton = dynamicsButtons.find(btn => {
                const label = btn.querySelector('.button-label');
                return label && label.textContent.trim() === '확인';
              });
              
              if (defaultButton) {
                defaultButton.click();
              } else {
                throw new Error('최종 확인 버튼을 찾을 수 없음');
              }
            });
            logger.info('JavaScript로 최종 "확인" 버튼 클릭 성공');
          } catch (jsError) {
            logger.error(`JavaScript로도 최종 "확인" 버튼을 찾지 못함: ${jsError.message}`);
            logger.warn('사용자가 수동으로 "확인" 버튼을 클릭해야 할 수 있습니다.');
            
            // 사용자에게 안내
            await page.evaluate(() => {
              alert('자동으로 "확인" 버튼을 클릭할 수 없습니다. 수동으로 "확인" 버튼을 클릭해주세요.');
            });
            
            // 사용자가 확인 버튼을 클릭할 때까지 대기 (20초)
            logger.info('사용자의 수동 "확인" 버튼 클릭 대기 중... (20초)');
            await delay(20000);
          }
        }
      }
    }

    // 추가 동작 7: 마지막 "확인" 버튼(kpc_exceluploadforledgerjournal_2_OKButton) 클릭
    await delay(5000);  // 이전 단계 완료 후 충분히 대기
    logger.info('마지막 "확인" 버튼(kpc_exceluploadforledgerjournal_2_OKButton) 찾는 중...');

    try {
      // ID로 시도
      await page.waitForSelector('#kpc_exceluploadforledgerjournal_2_OKButton', { 
        visible: true, 
        timeout: 10000 
      });
      await page.click('#kpc_exceluploadforledgerjournal_2_OKButton');
      logger.info('ID로 마지막 "확인" 버튼 클릭 성공');
    } catch (error) {
      logger.warn(`ID로 마지막 "확인" 버튼을 찾지 못함: ${error.message}`);
      
      // 이름 속성으로 시도
      try {
        await page.waitForSelector('button[name="OKButton"][id*="kpc_exceluploadforledgerjournal"]', { 
          visible: true, 
          timeout: 5000 
        });
        await page.click('button[name="OKButton"][id*="kpc_exceluploadforledgerjournal"]');
        logger.info('name 속성으로 마지막 "확인" 버튼 클릭 성공');
      } catch (nameError) {
        logger.warn(`name 속성으로도 마지막 "확인" 버튼을 찾지 못함: ${nameError.message}`);
        
        // 레이블로 시도
        try {
          await page.waitForSelector('#kpc_exceluploadforledgerjournal_2_OKButton_label', { 
            visible: true, 
            timeout: 5000 
          });
          await page.click('#kpc_exceluploadforledgerjournal_2_OKButton_label');
          logger.info('레이블로 마지막 "확인" 버튼 클릭 성공');
        } catch (labelError) {
          logger.warn(`레이블로도 마지막 "확인" 버튼을 찾지 못함: ${labelError.message}`);
          
          // JavaScript로 시도
          try {
            await page.evaluate(() => {
              // 방법 1: ID로 찾기
              const okButton = document.querySelector('#kpc_exceluploadforledgerjournal_2_OKButton');
              if (okButton) {
                okButton.click();
                return;
              }
              
              // 방법 2: 레이블로 찾기
              const okLabel = document.querySelector('#kpc_exceluploadforledgerjournal_2_OKButton_label');
              if (okLabel) {
                okLabel.click();
                return;
              }
              
              // 방법 3: ID 패턴으로 찾기
              const buttons = Array.from(document.querySelectorAll('button[id*="kpc_exceluploadforledgerjournal"][id*="OKButton"]'));
              if (buttons.length > 0) {
                buttons[0].click();
                return;
              }
              
              // 방법 4: 버튼 텍스트와 위치로 찾기
              const allButtons = Array.from(document.querySelectorAll('button.dynamicsButton'));
              const confirmButton = allButtons.find(btn => {
                const label = btn.querySelector('.button-label');
                return label && (label.textContent.trim() === '확인' || label.textContent.trim() === 'OK');
              });
              
              if (confirmButton) {
                confirmButton.click();
              } else {
                throw new Error('마지막 확인 버튼을 찾을 수 없음');
              }
            });
            logger.info('JavaScript로 마지막 "확인" 버튼 클릭 성공');
          } catch (jsError) {
            logger.error(`JavaScript로도 마지막 "확인" 버튼을 찾지 못함: ${jsError.message}`);
            logger.warn('사용자가 수동으로 마지막 "확인" 버튼을 클릭해야 할 수 있습니다.');
            
            // 사용자에게 안내
            await page.evaluate(() => {
              alert('자동으로 마지막 "확인" 버튼을 클릭할 수 없습니다. 수동으로 "확인" 버튼을 클릭해주세요.');
            });
            
            // 사용자가 수동으로 작업할 시간 제공
            logger.info('사용자의 수동 마지막 "확인" 버튼 클릭 대기 중... (20초)');
            await delay(20000);
          }
        }
      }
    }    // 추가 동작 7 완료 - 작업 완료 처리 (추가 동작 8, 9 제거됨)
    await delay(3000);  // 마지막 확인 버튼 클릭 후 대기
    logger.info('추가 동작 1-7 완료 - 작업 마무리 중...');

    // 작업 완료 팝업 표시 (수정 후)
    await page.evaluate(() => {
      alert('EZVoucher.js 작업완료. 확인 버튼을 누르면 창이 닫힙니다.');
    });

    // 성공적으로 완료되면 브라우저 닫기
    logger.info('RPA 작업 성공적으로 완료됨. 브라우저를 종료합니다...');
    await browser.close(); // 브라우저 창 닫기
    logger.info('브라우저가 성공적으로, 작업 요약 전 종료되었습니다.');

    // 성공적으로 완료되면 브라우저를 유지
    logger.info('RPA 작업 성공적으로 완료됨. 브라우저 유지 중...');
    logger.info('----- 작업 요약 -----');
    logger.info('1. D365 페이지 접속 및 로그인');
    logger.info('2. 5초 대기 후 즐겨찾기 아이콘 클릭');
    logger.info('3. "엑셀 전표 업로드" 메뉴 클릭');
    logger.info('4. lookupButton 클릭');
    logger.info('5. "일반전표(ARK)" 텍스트 필드 클릭');
    logger.info('6. 텍스트 박스에 "test" 입력');
    logger.info('7. "업로드" 버튼 클릭');
    logger.info('8. "Browse" 버튼 클릭');
    logger.info('9. "ARK전표업로드 양식" 파일 선택');
    logger.info('10. 파일 선택 대화상자에서 "확인" 버튼 클릭');
    logger.info('11. 최종 "확인" 버튼 클릭');
    logger.info('12. 마지막 "확인" 버튼(kpc_exceluploadforledgerjournal_2_OKButton) 클릭');
    logger.info('13. 분개장 배치 번호 요소 더블클릭');
    logger.info('---------------------');
    
    // 브라우저 유지 (의도적으로 닫지 않음)
    // await browser.close();
    
    // 반환 부분 (수정 후)
    return { 
      success: true, 
      message: 'RPA가 성공적으로 완료되었습니다.',
      completedAt: new Date().toISOString(),
      browserClosed: true
    };
    
  } catch (error) {
    logger.error(`RPA 오류 발생: ${error.message}`);
    
    
    return { success: false, error: error.message, browser: browser };
  }
}

// 파일 번호 범위를 받아 순차적으로 처리하는 함수 (기존 processAllFiles 수정)
async function processAllFiles(credentials, startFileNumber = 1, endFileNumber = 17) {
  // 폴더 경로가 설정되지 않은 경우 오류 반환
  if (!folderPath) {
    const errorMsg = '폴더 경로가 설정되지 않았습니다. 먼저 "폴더 지정" 버튼을 클릭하여 폴더를 선택해주세요.';
    logger.error(errorMsg);
    throw new Error(errorMsg);
  }

  logger.info(`RPA 프로세스 시작: 파일 ${startFileNumber}-${endFileNumber} 순차 처리`);
  logger.info(`사용할 폴더 경로: ${folderPath}`);
  
  // 성공 및 실패 카운트
  let successCount = 0;
  let failCount = 0;
  
  const browser = await puppeteerExtra.launch({
    headless: false,
    defaultViewport: null,
    args: [
      '--start-maximized',
      '--ignore-certificate-errors',
      '--ignore-certificate-errors-spki-list',
      '--allow-insecure-localhost',
      '--no-sandbox',
      '--disable-setuid-sandbox'
    ]
  });

  try {
    const page = await browser.newPage();
    
    // SSL 인증서 오류 처리
    await page.setBypassCSP(true);
    
    // 페이지 요청 인터셉트 설정 (SSL 오류 처리용)
    await page.setRequestInterception(true);
    page.on('request', request => {
      request.continue();
    });
      // 대화상자 처리 (인증서 경고 등)
    page.on('dialog', async dialog => {
      logger.info(`대화상자 감지: ${dialog.message()}`);
      await dialog.accept();
    });
    
    // 1. D365 페이지 접속 (재시도 로직 추가)
    logger.info('D365 페이지로 이동 중...');
    let pageLoadSuccess = false;
    let retryCount = 0;
    const maxRetries = 3;
    
    while (!pageLoadSuccess && retryCount < maxRetries) {
      try {
        retryCount++;
        logger.info(`D365 페이지 접속 시도 ${retryCount}/${maxRetries}`);
        
        await page.goto('https://d365.nepes.co.kr/namespaces/AXSF/?cmp=K02&mi=DefaultDashboard', {
          waitUntil: 'networkidle2',
          timeout: 60000 // 60초 타임아웃
        });
        
        pageLoadSuccess = true;
        logger.info('D365 페이지 로드 완료');
      } catch (networkError) {
        logger.error(`D365 페이지 접속 시도 ${retryCount} 실패: ${networkError.message}`);
        
        if (retryCount >= maxRetries) {
          const errorMsg = `네트워크 연결 실패: D365 사이트(https://d365.nepes.co.kr)에 접속할 수 없습니다. 인터넷 연결을 확인하거나 VPN이 필요할 수 있습니다.`;
          logger.error(errorMsg);
          throw new Error(errorMsg);
        }
        
        // 재시도 전 5초 대기
        logger.info('5초 후 재시도합니다...');
        await delay(5000);
      }
    }

    // 로그인 처리 (필요한 경우)
    if (await page.$('input[type="email"]') !== null || await page.$('#userNameInput') !== null) {
      logger.info('로그인 화면 감지됨, 로그인 시도 중...');
      await handleLogin(page, credentials);
    }
    
    // 로그인 후 페이지가 완전히 로드될 때까지 5초 대기
    logger.info('로그인 후 페이지가 완전히 로드될 때까지 5초 대기 중...');
    await delay(5000);  // 5초 대기
    logger.info('5초 대기 완료');// 지정된 범위의 파일을 순차적으로 처리
    for (let fileNumber = startFileNumber; fileNumber <= endFileNumber; fileNumber++) {
      logger.info(`======== 파일 ${fileNumber}번 처리 시작 ========`);
      
      // 해당 번호로 시작하는 파일 존재 여부 확인
      const excelFilePath = findExcelFileStartingWithNumber(folderPath, fileNumber);
      
      if (!excelFilePath) {
        logger.warn(`${fileNumber}.로 시작하는 엑셀 파일을 찾지 못했습니다. 다음 번호로 넘어갑니다.`);
        continue;
      }
      
      logger.info(`${fileNumber}. 파일 찾음: ${path.basename(excelFilePath)}`);
      
      // 각 파일에 대한 처리 시작
      try {
        // 첫 번째 파일 처리 시 또는 매 파일처리 시작 시 즐겨찾기 메뉴 클릭
        // 즐겨찾기 아이콘 클릭
        logger.info('즐겨찾기 아이콘 찾는 중...');
        try {
          await page.waitForSelector('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"][data-dyn-image-type="Symbol"]', { 
            visible: true,
            timeout: 10000
          });
          
          await page.click('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"][data-dyn-image-type="Symbol"]');
          logger.info('정확한 선택자로 즐겨찾기 아이콘 클릭 성공');
        } catch (error) {
          // 다른 선택자로 시도
          try {
            await page.waitForSelector('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"]', { 
              visible: true, 
              timeout: 5000 
            });
            await page.click('span.workspace-image.StarEmpty-symbol[data-dyn-title="즐겨찾기"]');
            logger.info('단순 선택자로 즐겨찾기 아이콘 클릭 성공');
          } catch (iconError) {
            // JavaScript로 시도
            try {
              await page.evaluate(() => {
                const spans = Array.from(document.querySelectorAll('span'));
                const favIcon = spans.find(span => 
                  span.getAttribute('data-dyn-title') === '즐겨찾기' || 
                  span.classList.contains('StarEmpty-symbol') ||
                  (span.className && span.className.includes('StarEmpty'))
                );
                
                if (favIcon) {
                  favIcon.click();
                } else {
                  throw new Error('JavaScript로도 즐겨찾기 아이콘을 찾을 수 없음');
                }
              });
              logger.info('JavaScript로 즐겨찾기 아이콘 클릭 성공');
            } catch (jsError) {
              logger.error(`JavaScript로도 즐겨찾기 아이콘을 찾지 못함: ${jsError.message}`);
            }
          }
        }
        
        // 클릭 후 메뉴가 표시될 때까지 잠시 대기
        await delay(3000);  // 3초 대기
        logger.info('즐겨찾기 메뉴 로드 대기 완료');

        // "엑셀 전표 업로드" 메뉴 클릭
        logger.info('"엑셀 전표 업로드" 메뉴 아이템 찾는 중...');
        try {
          const exactSelector = 'div.modulesPane-link.modulesFlyout-isFavorite[data-dyn-selected="false"][role="treeitem"] a.modulesPane-linkText[data-dyn-title="엑셀 전표 업로드"][role="link"]';
          await page.waitForSelector(exactSelector, { visible: true, timeout: 10000 });
          await page.click(exactSelector);
          logger.info('정확한 선택자로 "엑셀 전표 업로드" 메뉴 클릭 완료');
        } catch (error) {
          // 다른 선택자로 시도
          try {
            const selector = 'div[data-dyn-title="엑셀 전표 업로드"], div.modulesPane-link a[data-dyn-title="엑셀 전표 업로드"], .modulesPane-link a.modulesPane-linkText[data-dyn-title="엑셀 전표 업로드"]';
            await page.waitForSelector(selector, { visible: true, timeout: 5000 });
            await page.click(selector);
            logger.info('기본 선택자로 "엑셀 전표 업로드" 메뉴 클릭 완료');
          } catch (selectorError) {
            // JavaScript로 시도
            try {
              await page.evaluate(() => {
                const links = Array.from(document.querySelectorAll('a.modulesPane-linkText, div.modulesPane-link a, a[role="link"]'));
                const targetLink = links.find(link => link.textContent.includes('엑셀 전표 업로드'));
                if (targetLink) {
                  targetLink.click();
                } else {
                  throw new Error('링크를 찾을 수 없음');
                }
              });
              logger.info('JavaScript로 "엑셀 전표 업로드" 메뉴 클릭 완료');
            } catch (jsError) {
              logger.error(`JavaScript로도 메뉴를 찾지 못함: ${jsError.message}`);
            }
          }
        }
        
        // 엑셀 전표 업로드 페이지 로드 대기
        logger.info('엑셀 전표 업로드 페이지 로드 대기 중...');
        await delay(5000);  // 5초 대기
        logger.info('엑셀 전표 업로드 페이지 로드 완료');

        // lookupButton 클릭
        logger.info('lookupButton 클래스 요소 찾는 중...');
        try {
          await page.waitForSelector('.lookupButton[title="오픈"]', { 
            visible: true, 
            timeout: 10000 
          });
          await page.click('.lookupButton[title="오픈"]');
          logger.info('lookupButton 클릭 성공');
        } catch (error) {
          // JavaScript로 시도
          try {
            await page.evaluate(() => {
              const lookupButtons = Array.from(document.querySelectorAll('.lookupButton'));
              const button = lookupButtons.find(btn => 
                btn.getAttribute('title') === '오픈' || 
                btn.getAttribute('data-dyn-bind')?.includes('Input_LookupTooltip')
              );
              
              if (button) {
                button.click();
              } else {
                throw new Error('lookupButton을 찾을 수 없음');
              }
            });
            logger.info('JavaScript로 lookupButton 클릭 성공');
          } catch (jsError) {
            logger.error(`JavaScript로도 lookupButton을 찾지 못함: ${jsError.message}`);
          }
        }

        // 팝업이 열릴 때까지 대기
        await delay(3000);
        logger.info('lookupButton 클릭 후 팝업 대기 완료');

        // "일반전표(ARK)" 텍스트 필드 클릭
        logger.info('"일반전표(ARK)" 텍스트 필드 찾는 중...');
        try {
          await page.waitForSelector('input[value="일반전표(ARK)"], input[title="일반전표(ARK)"]', { 
            visible: true, 
            timeout: 10000 
          });
          await page.click('input[value="일반전표(ARK)"], input[title="일반전표(ARK)"]');
          logger.info('"일반전표(ARK)" 텍스트 필드 클릭 성공');
        } catch (error) {
          // ID로 시도
          try {
            await page.waitForSelector('#SysGen_Name_125_0_0_input', { visible: true, timeout: 5000 });
            await page.click('#SysGen_Name_125_0_0_input');
            logger.info('ID로 "일반전표(ARK)" 텍스트 필드 클릭 성공');
          } catch (idError) {
            // JavaScript로 시도
            try {
              await page.evaluate(() => {
                const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
                const input = inputs.find(inp => 
                  inp.value === '일반전표(ARK)' || 
                  inp.title === '일반전표(ARK)'
                );
                
                if (input) {
                  input.click();
                } else {
                  throw new Error('텍스트 필드를 찾을 수 없음');
                }
              });
              logger.info('JavaScript로 "일반전표(ARK)" 텍스트 필드 클릭 성공');
            } catch (jsError) {
              logger.error(`JavaScript로도 텍스트 필드를 찾지 못함: ${jsError.message}`);
            }
          }
        }

        await delay(2000);        logger.info('텍스트 필드 클릭 후 대기 완료');        // 파일명에서 괄호 안의 텍스트를 추출하여 입력
        logger.info('텍스트 박스 찾아 파일명 괄호 안의 텍스트 입력 중...');
        
        let textToInput;
        try {
          // 파일명에서 괄호 안의 텍스트 추출 - 반드시 성공해야 함
          textToInput = extractTextFromParentheses(excelFilePath);
          
          if (!textToInput || textToInput.trim() === '') {
            throw new Error(`파일명 괄호 안에 유효한 텍스트가 없습니다. 파일: ${path.basename(excelFilePath)}`);
          }
          
          logger.info(`파일명 괄호 안의 텍스트 "${textToInput}" 추출 성공`);
        } catch (extractError) {
          const errorMsg = `파일명 괄호 텍스트 추출 실패: ${extractError.message}. 파일: ${path.basename(excelFilePath)}. 작업을 중단합니다.`;
          logger.error(errorMsg);
          failCount++;          throw new Error(errorMsg);
        }

        // 텍스트 입력 요소가 완전히 로드될 때까지 추가 대기
        logger.info('텍스트 입력 페이지 완전 로드 대기 중...');
        await delay(3000);

        // 텍스트 입력 - 강화된 5번 시도
        let inputSuccess = false;
        let lastError = null;

        // 첫 번째 시도: 메인 ID 선택자 (강화된 버전)
        if (!inputSuccess) {
          try {
            logger.info('첫 번째 시도: 메인 ID 선택자로 텍스트 입력 (강화된 대기)');
            
            // 페이지가 완전히 안정화될 때까지 대기
            await page.waitForFunction(
              () => document.readyState === 'complete',
              { timeout: 10000 }
            );
            
            // 여러 가능한 선택자들을 시도
            const possibleSelectors = [
              '#kpc_exceluploadforledgerjournal_2_FormStringControl_Txt_input',
              'input[id*="FormStringControl_Txt_input"]',
              'input[id*="kpc_exceluploadforledgerjournal"][id*="Txt_input"]',
              'input[role="textbox"]'
            ];
            
            let foundElement = null;
            for (const selector of possibleSelectors) {
              try {
                await page.waitForSelector(selector, { visible: true, timeout: 3000 });
                foundElement = selector;
                logger.info(`요소 찾음: ${selector}`);
                break;
              } catch (e) {
                logger.warn(`선택자 실패: ${selector}`);
              }
            }
            
            if (foundElement) {
              // 요소가 클릭 가능할 때까지 대기
              await page.waitForFunction(
                (sel) => {
                  const element = document.querySelector(sel);
                  return element && !element.disabled && element.offsetParent !== null;
                },
                { timeout: 5000 },
                foundElement
              );
              
              await page.click(foundElement);
              await delay(500); // 클릭 후 잠시 대기
              await page.type(foundElement, textToInput);
              logger.info(`첫 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
              inputSuccess = true;
            } else {
              throw new Error('모든 선택자에서 요소를 찾을 수 없음');
            }
          } catch (error) {
            lastError = error;
            logger.warn(`첫 번째 시도 실패: ${error.message}`);
          }
        }

        // 두 번째 시도: 클래스 선택자 (강화된 버전)
        if (!inputSuccess) {
          try {
            logger.info('두 번째 시도: 클래스 선택자로 텍스트 입력');
            
            const classSelectors = [
              'input.textbox.field.displayoption[role="textbox"]',
              'input.textbox.field.displayoption',
              'input.textbox[role="textbox"]',
              'input[class*="textbox"][class*="field"]'
            ];
            
            let foundElement = null;
            for (const selector of classSelectors) {
              try {
                await page.waitForSelector(selector, { visible: true, timeout: 3000 });
                foundElement = selector;
                logger.info(`클래스 요소 찾음: ${selector}`);
                break;
              } catch (e) {
                logger.warn(`클래스 선택자 실패: ${selector}`);
              }
            }
            
            if (foundElement) {
              await page.click(foundElement);
              await delay(500);
              await page.type(foundElement, textToInput);
              logger.info(`두 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
              inputSuccess = true;
            } else {
              throw new Error('모든 클래스 선택자에서 요소를 찾을 수 없음');
            }
            logger.info(`두 번째 시도 성공: 텍스트 박스에 "${textToInput}" 입력 완료`);
            inputSuccess = true;
          } catch (error) {
            lastError = error;
            logger.warn(`두 번째 시도 실패: ${error.message}`);
          }
        }

        // 세 번째 시도: JavaScript로 직접 조작
        if (!inputSuccess) {
          try {
            logger.info('세 번째 시도: JavaScript로 직접 텍스트 입력');
            await page.evaluate((text) => {
              const inputs = Array.from(document.querySelectorAll('input[role="textbox"]'));
              const input = inputs.find(inp => 
                inp.className.includes('textbox') || 
                inp.id?.includes('FormStringControl_Txt_input')
              );
              
              if (input) {
                input.focus();
                input.value = text;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
              } else {
                throw new Error('텍스트 박스를 찾을 수 없음');
              }
            }, textToInput);
            logger.info(`세 번째 시도 성공: JavaScript로 텍스트 박스에 "${textToInput}" 입력 완료`);
            inputSuccess = true;
          } catch (error) {
            lastError = error;
            logger.warn(`세 번째 시도 실패: ${error.message}`);
          }
        }

        // 네 번째 시도: 모든 텍스트 입력 필드 검색
        if (!inputSuccess) {
          try {
            logger.info('네 번째 시도: 모든 텍스트 입력 필드에서 검색');
            await page.evaluate((text) => {
              const allInputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type]), textarea'));
              
              // 설명란으로 추정되는 입력 필드 찾기
              const descriptionInput = allInputs.find(inp => 
                inp.placeholder?.includes('설명') ||
                inp.title?.includes('설명') ||
                inp.getAttribute('aria-label')?.includes('설명') ||
                inp.id?.includes('Description') ||
                inp.id?.includes('Txt') ||
                inp.name?.includes('description')
              );
              
              if (descriptionInput) {
                descriptionInput.focus();
                descriptionInput.value = text;
                descriptionInput.dispatchEvent(new Event('input', { bubbles: true }));
                descriptionInput.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
              } else {
                throw new Error('설명란으로 추정되는 입력 필드를 찾을 수 없음');
              }
            }, textToInput);
            logger.info(`네 번째 시도 성공: 설명란 입력 필드에 "${textToInput}" 입력 완료`);
            inputSuccess = true;
          } catch (error) {
            lastError = error;
            logger.warn(`네 번째 시도 실패: ${error.message}`);
          }
        }

        // 다섯 번째 시도: 가장 가능성 높은 입력 필드 선택
        if (!inputSuccess) {
          try {
            logger.info('다섯 번째 시도: 가장 가능성 높은 입력 필드 선택');
            await page.evaluate((text) => {
              const allInputs = Array.from(document.querySelectorAll('input, textarea'));
              
              // 보이는 입력 필드만 필터링
              const visibleInputs = allInputs.filter(inp => {
                const style = window.getComputedStyle(inp);
                return style.display !== 'none' && style.visibility !== 'hidden' && inp.offsetWidth > 0 && inp.offsetHeight > 0;
              });
              
              if (visibleInputs.length > 0) {
                // 가장 큰 입력 필드 선택 (설명란일 가능성이 높음)
                const largestInput = visibleInputs.reduce((prev, current) => {
                  return (prev.offsetWidth * prev.offsetHeight) > (current.offsetWidth * current.offsetHeight) ? prev : current;
                });
                
                largestInput.focus();
                largestInput.value = text;
                largestInput.dispatchEvent(new Event('input', { bubbles: true }));
                largestInput.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
              } else {
                throw new Error('보이는 입력 필드를 찾을 수 없음');
              }
            }, textToInput);
            logger.info(`다섯 번째 시도 성공: 가장 큰 입력 필드에 "${textToInput}" 입력 완료`);
            inputSuccess = true;
          } catch (error) {
            lastError = error;
            logger.warn(`다섯 번째 시도 실패: ${error.message}`);
          }        }        // 모든 시도 실패 시 RPA 작업 중단 (필수 동작)
        if (!inputSuccess) {
          const errorMsg = `🚨 중요: 추가 동작 3번 실패 - 괄호 안 텍스트 "${textToInput}" 입력에 모든 시도가 실패했습니다. 마지막 오류: ${lastError?.message || '알 수 없는 오류'}. RPA 작업을 중단합니다.`;
          logger.error(errorMsg);
          
          // 브라우저 닫기
          try {
            await browser.close();
          } catch (closeError) {
            logger.error(`브라우저 종료 중 오류: ${closeError.message}`);
          }
          
          // 오류 발생으로 작업 중단
          throw new Error(errorMsg);
        }

        await delay(2000);
        logger.info('텍스트 입력 후 대기 완료');

        // "업로드" 버튼 클릭
        logger.info('"업로드" 버튼 찾는 중...');
        try {
          await page.waitForSelector('#kpc_exceluploadforledgerjournal_2_UploadButton_label', { 
            visible: true, 
            timeout: 10000 
          });
          await page.click('#kpc_exceluploadforledgerjournal_2_UploadButton_label');
          logger.info('ID로 "업로드" 버튼 클릭 성공');
        } catch (error) {
          // 텍스트로 시도
          try {
            await page.waitForSelector('span.button-label:contains("업로드")', { 
              visible: true, 
              timeout: 5000 
            });
            await page.click('span.button-label:contains("업로드")');
            logger.info('텍스트로 "업로드" 버튼 클릭 성공');
          } catch (textError) {
            // JavaScript로 시도
            try {
              await page.evaluate(() => {
                const spans = Array.from(document.querySelectorAll('span.button-label, span[id*="UploadButton_label"]'));
                const uploadButton = spans.find(span => span.textContent.trim() === '업로드');
                
                if (uploadButton) {
                  uploadButton.click();
                } else {
                  const buttons = Array.from(document.querySelectorAll('button, div.button-container, [role="button"]'));
                  const btn = buttons.find(b => {
                    const text = b.textContent.trim();
                    return text === '업로드' || text.includes('업로드');
                  });
                  
                  if (btn) {
                    btn.click();
                  } else {
                    throw new Error('업로드 버튼을 찾을 수 없음');
                  }
                }
              });
              logger.info('JavaScript로 "업로드" 버튼 클릭 성공');
            } catch (jsError) {
              logger.error(`JavaScript로도 "업로드" 버튼을 찾지 못함: ${jsError.message}`);
            }
          }
        }

        // "Browse" 버튼 클릭 및 파일 선택
        await delay(3000); // 업로드 버튼 클릭 후 잠시 대기
        logger.info('"Browse" 버튼 찾는 중...');
        
        try {
          // fileChooser를 사용하여 파일 선택
          try {
            // 파일 선택기가 열릴 때까지 대기하면서 Browse 버튼 클릭
            const [fileChooser] = await Promise.all([
              page.waitForFileChooser({ timeout: 10000 }),
              page.click('#Dialog_4_UploadBrowseButton, button[name="UploadBrowseButton"]')
            ]);
            
            // 찾은 파일 선택
            await fileChooser.accept([excelFilePath]);
            logger.info(`fileChooser 방식으로 파일 선택 완료: ${path.basename(excelFilePath)}`);
          } catch (chooserError) {
            // 파일 입력 필드를 직접 찾아 조작
            try {
              // 먼저 Brows 버튼 클릭 취소(이미 클릭했을 수 있기 때문에)
              await page.keyboard.press('Escape');
              await delay(1000);
              
              // 파일 입력 필드 찾기
              const fileInputSelector = 'input[type="file"]';
              const fileInput = await page.$(fileInputSelector);
              
              if (fileInput) {
                // 파일 입력 필드가 있으면 직접 파일 설정
                await fileInput.uploadFile(excelFilePath);
                logger.info(`uploadFile 방식으로 파일 선택 완료: ${path.basename(excelFilePath)}`);
              } else {
                // 파일 입력 필드가 없으면 다시 Browse 버튼 클릭 시도
                await page.click('#Dialog_4_UploadBrowseButton, button[name="UploadBrowseButton"]');
                
                // 사용자에게 파일 선택 안내
                await page.evaluate((fileName) => {
                  alert(`자동 파일 선택에 실패했습니다. 파일 탐색기에서 "${fileName}" 파일을 수동으로 선택해주세요.`);
                }, path.basename(excelFilePath));
                
                // 사용자가 파일을 선택할 때까지 대기 (30초)
                logger.info('사용자의 수동 파일 선택 대기 중... (30초)');
                await delay(30000);
              }
            } catch (inputError) {
              // 최후의 방법: 사용자에게 안내
              await page.evaluate((message, fileName) => {
                alert(`${message} 파일 탐색기에서 "${fileName}" 파일을 수동으로 선택해주세요.`);
              }, `자동 파일 선택에 실패했습니다.`, path.basename(excelFilePath));
              
              // 사용자가 파일을 선택할 때까지 대기 (30초)
              logger.info('사용자의 수동 파일 선택 대기 중... (30초)');
              await delay(30000);
            }
          }
          
          // 파일 선택 후 대기
          await delay(3000);
          logger.info('파일 선택 후 대기 완료');
          
          // 파일 선택 대화상자에서 "확인" 버튼이 필요한 경우 클릭
          try {
            const confirmButtonSelector = '#Dialog_4_OkButton, #SysOKButton, span.button-label:contains("확인"), span.button-label:contains("OK")';
            const confirmButton = await page.$(confirmButtonSelector);
            
            if (confirmButton) {
              await confirmButton.click();
              logger.info('파일 선택 대화상자의 "확인" 버튼 클릭 성공');
            }
          } catch (confirmError) {
            logger.warn(`확인 버튼 클릭 시도 중 오류: ${confirmError.message}`);
            logger.info('계속 진행합니다...');
          }
        } catch (browseError) {
          // 사용자에게 파일 선택 안내
          try {
            await page.evaluate((fileNumber) => {
              alert(`자동화 스크립트에 문제가 발생했습니다. 수동으로 "Browse" 버튼을 클릭하고 "${fileNumber}."로 시작하는 엑셀 파일을 선택해주세요.`);
            }, fileNumber);
            
            // 사용자가 작업을 완료할 때까지 대기
            logger.info('사용자의 수동 파일 선택 대기 중... (60초)');
            await delay(60000);
          } catch (alertError) {
            logger.error(`알림 표시 중 오류: ${alertError.message}`);
          }
        }

        // 파일 선택 후 최종 "확인" 버튼 클릭
        await delay(5000);  // 파일 선택 후 충분히 대기
        logger.info('파일 선택 후 최종 "확인" 버튼 찾는 중...');

        try {
          // ID로 시도
          await page.waitForSelector('#Dialog_4_OkButton', { 
            visible: true, 
            timeout: 10000 
          });
          await page.click('#Dialog_4_OkButton');
          logger.info('ID로 최종 "확인" 버튼 클릭 성공');
        } catch (error) {
          // 이름 속성으로 시도
          try {
            await page.waitForSelector('button[name="OkButton"]', { 
              visible: true, 
              timeout: 5000 
            });
            await page.click('button[name="OkButton"]');
            logger.info('name 속성으로 최종 "확인" 버튼 클릭 성공');
          } catch (nameError) {
            // JavaScript로 시도
            try {
              await page.evaluate(() => {
                // 방법 1: ID로 찾기
                const okButton = document.querySelector('#Dialog_4_OkButton');
                if (okButton) {
                  okButton.click();
                  return;
                }
                
                // 방법 2: 레이블로 찾기
                const okLabel = document.querySelector('#Dialog_4_OkButton_label');
                if (okLabel) {
                  okLabel.click();
                  return;
                }
                
                // 방법 3: 버튼 텍스트로 찾기
                const buttons = Array.from(document.querySelectorAll('button, span.button-label'));
                const button = buttons.find(btn => btn.textContent.trim() === '확인');
                if (button) {
                  button.click();
                  return;
                }
                
                // 방법 4: 버튼 클래스와 속성으로 찾기
                const dynamicsButtons = Array.from(document.querySelectorAll('button.dynamicsButton.button-isDefault'));
                const defaultButton = dynamicsButtons.find(btn => {
                  const label = btn.querySelector('.button-label');
                  return label && label.textContent.trim() === '확인';
                });
                
                if (defaultButton) {
                  defaultButton.click();
                } else {
                  throw new Error('최종 확인 버튼을 찾을 수 없음');
                }
              });
              logger.info('JavaScript로 최종 "확인" 버튼 클릭 성공');
            } catch (jsError) {
              // 사용자에게 안내
              await page.evaluate(() => {
                alert('자동으로 "확인" 버튼을 클릭할 수 없습니다. 수동으로 "확인" 버튼을 클릭해주세요.');
              });
              
              // 사용자가 확인 버튼을 클릭할 때까지 대기 (20초)
              logger.info('사용자의 수동 "확인" 버튼 클릭 대기 중... (20초)');
              await delay(20000);
            }
          }
        }

        // 마지막 "확인" 버튼 클릭
        await delay(5000);  // 이전 단계 완료 후 충분히 대기
        logger.info('마지막 "확인" 버튼(kpc_exceluploadforledgerjournal_2_OKButton) 찾는 중...');

        try {
          // ID로 시도
          await page.waitForSelector('#kpc_exceluploadforledgerjournal_2_OKButton', { 
            visible: true, 
            timeout: 10000 
          });
          await page.click('#kpc_exceluploadforledgerjournal_2_OKButton');
          logger.info('ID로 마지막 "확인" 버튼 클릭 성공');
        } catch (error) {
          // JavaScript로 시도
          try {
            await page.evaluate(() => {
              // ID 패턴으로 찾기
              const buttons = Array.from(document.querySelectorAll('button[id*="kpc_exceluploadforledgerjournal"][id*="OKButton"]'));
              if (buttons.length > 0) {
                buttons[0].click();
                return;
              }
              
              // 버튼 텍스트와 위치로 찾기
              const allButtons = Array.from(document.querySelectorAll('button.dynamicsButton'));
              const confirmButton = allButtons.find(btn => {
                const label = btn.querySelector('.button-label');
                return label && (label.textContent.trim() === '확인' || label.textContent.trim() === 'OK');
              });
              
              if (confirmButton) {
                confirmButton.click();
              } else {
                throw new Error('마지막 확인 버튼을 찾을 수 없음');
              }
            });
            logger.info('JavaScript로 마지막 "확인" 버튼 클릭 성공');
          } catch (jsError) {
            // 사용자에게 안내
            await page.evaluate(() => {
              alert('자동으로 마지막 "확인" 버튼을 클릭할 수 없습니다. 수동으로 "확인" 버튼을 클릭해주세요.');
            });
            
            // 사용자가 수동으로 작업할 시간 제공
            logger.info('사용자의 수동 마지막 "확인" 버튼 클릭 대기 중... (20초)');
            await delay(20000);
          }        }

        // 추가 동작 7번까지 완료 - 작업 완료 처리
        logger.info('추가 동작 7번(마지막 확인 버튼 클릭)까지 완료');
        
        // 마지막 확인 버튼 클릭 후 페이지 로드 대기
        await delay(5000);

        // 해당 파일 처리 성공 카운트 증가
        successCount++;
        logger.info(`파일 번호 ${fileNumber} 처리 성공`);
      } catch (fileProcessError) {
        // 파일 처리 중 오류 발생 시 실패 카운트 증가
        failCount++;
        logger.error(`파일 번호 ${fileNumber} 처리 중 오류 발생: ${fileProcessError.message}`);
      }
      
      logger.info(`======== 파일 ${fileNumber}번 처리 완료 ========`);
        // 다음 파일 처리 전 페이지 초기 상태로 돌아가기
      if (fileNumber < endFileNumber) {
        try {
          // 페이지 새로고침
          await page.reload({ waitUntil: 'networkidle2' });
          await delay(5000); // 페이지 로드 대기
          
          logger.info('페이지 초기화 완료');
        } catch (resetError) {
          logger.warn(`페이지 초기화 중 오류 발생: ${resetError.message}. 계속 진행합니다.`);
        }
      }
    }
      // 모든 작업 완료 후 최종 결과 보고
    logger.info("=================================================");
    logger.info(`파일 ${startFileNumber}-${endFileNumber} 처리 완료. 성공: ${successCount}, 실패: ${failCount}`);
    logger.info("=================================================");
    
    // 작업 완료 팝업 표시
    await page.evaluate((successCount, failCount, startFileNumber, endFileNumber) => {
      alert(`파일 ${startFileNumber}-${endFileNumber} 처리가 완료되었습니다.\n성공: ${successCount}, 실패: ${failCount}\n확인 버튼을 누르면 창이 닫힙니다.`);
    }, successCount, failCount, startFileNumber, endFileNumber);
    
    // 모든 작업 완료 후 브라우저 닫기
    await browser.close();
      return {
      success: true,
      message: `파일 ${startFileNumber}-${endFileNumber} 처리 완료. 성공: ${successCount}, 실패: ${failCount}`,
      successCount,
      failCount,
      completedAt: new Date().toISOString()
    };
    
  } catch (error) {
    logger.error(`RPA 오류 발생: ${error.message}`);
    
    // 브라우저 닫기
    try {
      await browser.close();
    } catch (closeError) {
      logger.error(`브라우저 종료 중 오류: ${closeError.message}`);
    }
    
    return { 
      success: false, 
      error: error.message
    };
  }
}
/**
 * 로그인 처리 함수
 * @param {Object} page - Puppeteer 페이지 객체
 * @param {Object} credentials - 로그인 정보 (username, password)
 */
async function handleLogin(page, credentials) {
  try {
  
    // 1. 사용자 이름(이메일) 입력
    logger.info('사용자 이름 입력 중...');
    await page.waitForSelector('#userNameInput', { visible: true, timeout: 10000 });
    await page.type('#userNameInput', credentials.username);
    logger.info('사용자 이름 입력 완료');
    
    // 2. 비밀번호 입력
    logger.info('비밀번호 입력 중...');
    await page.waitForSelector('#passwordInput', { visible: true, timeout: 10000 });
    await page.type('#passwordInput', credentials.password);
    logger.info('비밀번호 입력 완료');
    
    // 3. 로그인 버튼 클릭
    logger.info('로그인 버튼 클릭 중...');
    await page.waitForSelector('#submitButton', { visible: true, timeout: 10000 });
    
    
    await page.click('#submitButton');
    logger.info('로그인 버튼 클릭 완료');
    
    // 로그인 후 페이지 로드 대기
    logger.info('로그인 후 페이지 로드 대기 중...');
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 });
    
    // 로그인 성공 확인
    logger.info('로그인 완료');
    
  } catch (error) {
    // 오류 시 스크린샷
    logger.error(`로그인 오류: ${error.message}`);
    throw error;
  }
}


// 사용자 인증 정보를 저장할 전역 변수
global.authCredentials = {
  username: '',
  password: ''
};

// 사용자 인증 정보를 관리하는 함수
function getCredentials() {
  return {
    username: global.authCredentials.username || process.env.D365_USERNAME || '',
    password: global.authCredentials.password || process.env.D365_PASSWORD || ''
  };
}

// 크레덴셜 설정 함수 (main.js에서 호출용)
function setCredentials(username, password) {
  if (!global.authCredentials) {
    global.authCredentials = {};
  }
  global.authCredentials.username = username;
  global.authCredentials.password = password;
  logger.info('EZVoucher 크레덴셜이 설정되었습니다:', username);
}

// 실행 예시
if (require.main === module) {
  const credentials = getCredentials();
  let globalBrowser; // 브라우저 인스턴스 전역 유지
  
  navigateToDynamics365(credentials)
    .then(result => {
      if (result && result.success) {
        logger.info('스크립트 실행 완료');
        logger.info('브라우저 창이 열린 상태로 유지됩니다. 종료하려면 Ctrl+C를 누르세요.');
        globalBrowser = result.browser; // 브라우저 인스턴스 저장
        
        // 프로세스 유지 (브라우저 창이 닫히지 않도록)
        process.stdin.resume();
        
        // 프로세스 종료 시 브라우저도 정리
        process.on('SIGINT', async () => {
          logger.info('프로세스 종료 요청이 감지되었습니다.');
          if (globalBrowser) {
            logger.info('브라우저 종료 중...');
            await globalBrowser.close();
          }
          process.exit(0);
        });
      } else if (result) {
        logger.warn(`스크립트 실행 중 오류 발생: ${result.error}`);
        logger.info('브라우저 창이 열린 상태로 유지됩니다. 종료하려면 Ctrl+C를 누르세요.');
        globalBrowser = result.browser;
        process.stdin.resume();
      }
    })
    .catch(err => {
      logger.error(`스크립트 실행 오류: ${err.message}`);
      // 오류가 발생해도 프로세스 유지
      logger.info('오류가 발생했지만 브라우저 창이 유지됩니다. 종료하려면 Ctrl+C를 누르세요.');
      process.stdin.resume();
    });
}

// interface 적용 전
// module.exports = { navigateToDynamics365 };

// interface 적용 후
const EventEmitter = require('events');
const taskEvents = new EventEmitter();


class EZVoucher {
  constructor() {
    // 필요한 속성 초기화
    this.browser = null;
    this.page = null;
    this.isLoggedIn = false;
    this.runningTasks = new Map();
  }

  // 이벤트 구독 함수
  onStatusUpdate(callback) {
    taskEvents.on('task-status-update', callback);
  }
  
  // 사용자 인증 정보 설정 메서드 추가
  setCredentials(username, password) {
    global.authCredentials = {
      username: username,
      password: password
    };
    logger.info(`사용자 인증 정보가 설정되었습니다. 사용자: ${username}`);
    return { 
      success: true,
      username: username,
      passwordSaved: password ? true : false
    };
  }

  // 저장된 인증 정보 확인 메서드 추가
  getCredentialsInfo() {
    return {
      username: global.authCredentials.username || '',
      hasSavedPassword: !!global.authCredentials.password,
      isValid: !!(global.authCredentials.username && global.authCredentials.password)
    };
  }

  // RPA 실행 함수
  async runAllTasks() {
    const credentials = getCredentials();
    const result = await processAllFiles(credentials);
    return result;
  }

  // 특정 작업 실행
  async runTask(taskName) {
    taskEvents.emit('task-status-update', {
      taskName,
      status: 'running',
      message: '작업 시작',
      timestamp: new Date().toISOString()
    });
    
    try {
      const credentials = getCredentials();
      const result = await navigateToDynamics365(credentials);
      
      if (result.success) {
        taskEvents.emit('task-status-update', {
          taskName,
          status: 'done',
          message: '작업 완료',
          timestamp: new Date().toISOString(),
          duration: '2분'
        });
        return result;
      } else {
        throw new Error(result.error || '알 수 없는 오류');
      }
    } catch (error) {      taskEvents.emit('task-status-update', {
        taskName,
        status: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      });
      throw error;
    }
  }

  // 파일 번호 범위 처리 메소드 추가
  async processFileRange(startFileNumber, endFileNumber) {
    taskEvents.emit('task-status-update', {
      taskName: `파일 ${startFileNumber}-${endFileNumber} 처리`,
      status: 'running',
      message: '범위 처리 시작',
      timestamp: new Date().toISOString()
    });
    
    try {
      const credentials = getCredentials();
      const result = await processFileRange(credentials, startFileNumber, endFileNumber);
      
      taskEvents.emit('task-status-update', {
        taskName: `파일 ${startFileNumber}-${endFileNumber} 처리`,
        status: 'done',
        message: result.message,
        timestamp: new Date().toISOString()
      });
      
      return result;
    } catch (error) {
      taskEvents.emit('task-status-update', {
        taskName: `파일 ${startFileNumber}-${endFileNumber} 처리`,
        status: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      });
      throw error;
    }
  }

  // 단일 파일 처리 메소드 추가
  async processSingleFile(fileNumber) {
    taskEvents.emit('task-status-update', {
      taskName: `파일 ${fileNumber} 처리`,
      status: 'running',
      message: '단일 파일 처리 시작',
      timestamp: new Date().toISOString()
    });
    
    try {
      const credentials = getCredentials();
      const result = await processSingleFile(credentials, fileNumber);
      
      taskEvents.emit('task-status-update', {
        taskName: `파일 ${fileNumber} 처리`,
        status: 'done',
        message: result.message,
        timestamp: new Date().toISOString()
      });
      
      return result;
    } catch (error) {
      taskEvents.emit('task-status-update', {
        taskName: `파일 ${fileNumber} 처리`,
        status: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      });
      throw error;
    }
  }
}

// 파일 번호 범위를 지정하여 처리하는 함수
async function processFileRange(credentials, startFileNumber, endFileNumber) {
  return await processAllFiles(credentials, startFileNumber, endFileNumber);
}

// 선택된 파일들을 처리하는 함수 (파일 범위 처리와 동일)
async function processSelectedFiles(credentials, startFileNumber, endFileNumber) {
  return await processAllFiles(credentials, startFileNumber, endFileNumber);
}

// 단일 파일을 처리하는 함수
async function processSingleFile(credentials, fileNumber) {
  return await processAllFiles(credentials, fileNumber, fileNumber);
}

// 싱글톤 인스턴스 생성 및 내보내기
const ezVoucher = new EZVoucher();

// 모듈에 setCredentials 함수 추가
ezVoucher.setCredentials = setCredentials;

module.exports = ezVoucher;


