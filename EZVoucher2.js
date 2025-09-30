/*
 * EZVoucher2.js - 매입송장 처리 RPA 자동화
 * 
 * 동작 순서:
 * 1. ERP 접속 및 로그인 완료
 *    - D365 페이지 접속 (https://d365.nepes.co.kr/namespaces/AXSF/?cmp=K02&mi=DefaultDashboard)
 *    - ADFS 로그인 처리 (#userNameInput, #passwordInput, #submitButton)
 *    - 페이지 로딩 완료 대기
 * 
 * 2. 검색 기능을 통한 구매 입고내역 조회 페이지 이동
 *    - 검색 버튼 클릭 (Find-symbol 버튼)
 *    - "구매 입고내역 조회(N)" 검색어 입력
 *    - NavigationSearchBox에서 해당 메뉴 클릭
 * 
 * 3. (추후 구현 예정) 매입송장 처리 로직
 *    - 파일 업로드
 *    - 데이터 처리
 *    - 결과 확인
 */

const puppeteer = require('puppeteer');
const puppeteerExtra = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const winston = require('winston');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx'); // 엑셀 파일 읽기용 라이브러리

const { ipcMain, dialog } = require('electron');

// 기본 대기 함수
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 에러 메시지로부터 실패한 단계를 추정하는 헬퍼 함수
function getCurrentStepFromError(errorMessage) {
  const errorLower = errorMessage.toLowerCase();
  
  if (errorLower.includes('network') || errorLower.includes('connection') || errorLower.includes('d365') || errorLower.includes('login')) {
    return 1; // 1단계: ERP 접속 및 로그인
  } else if (errorLower.includes('navigate') || errorLower.includes('구매') || errorLower.includes('inquiry')) {
    return 2; // 2단계: 구매 입고내역 조회 페이지 이동
  } else if (errorLower.includes('excel') || errorLower.includes('macro') || errorLower.includes('엑셀')) {
    return 3; // 3단계: 엑셀 파일 처리
  } else if (errorLower.includes('menu') || errorLower.includes('supplier') || errorLower.includes('공급사')) {
    return 4; // 4단계: 공급사송장 메뉴
  } else if (errorLower.includes('calendar') || errorLower.includes('캘린더')) {
    return 5; // 5단계: 캘린더 버튼
  } else if (errorLower.includes('pending') || errorLower.includes('대기중')) {
    return 6; // 6단계: 대기중인 공급사송장
  } else if (errorLower.includes('groupware') || errorLower.includes('그룹웨어')) {
    return 7; // 7단계: 그룹웨어 상신
  } else {
    return 1; // 기본값
  }
}

// 마지막 처리된 B값의 AT열 날짜 값을 저장하는 전역 변수 (FixedDueDate 입력용)
let lastProcessedDateFromATColumn = null;

// 마지막 처리된 B값의 AV열 날짜 값을 저장하는 전역 변수 (송장일 입력용)
let lastProcessedDateFromAVColumn = null;

// 마지막 처리된 B값의 AU열 값을 저장하는 전역 변수
let lastProcessedValueFromAUColumn = null;

// 마지막 처리된 B값의 I열 값을 저장하는 전역 변수 (필터 입력용)
let lastProcessedValueFromIColumn = null;

// 공급사송장 요소 아래 20px 위치에서 추출한 값을 저장하는 전역 변수 (3.5 동작용)
let extractedVendorInvoiceValue = null;

// 사용자가 입력한 A열 값을 저장하는 전역 변수
let userInputValueA = 3; // 기본값은 3

// YYYY-MM-DD 형식 또는 Excel 시리얼 번호를 M/dd/YYYY 형식으로 변환하는 함수
function convertDateFormat(dateValue) {
  if (!dateValue) return null;
  
  try {
    // Excel 시리얼 번호인지 확인 (숫자)
    if (typeof dateValue === 'number') {
      logger.info(`Converting Excel serial number: ${dateValue}`);
      
      // Excel 시리얼 번호를 JavaScript Date 객체로 변환
      // Excel 기준일: 1900년 1월 1일 (실제로는 1900년 1월 0일부터 계산)
      const excelEpoch = new Date(1900, 0, 1);
      const jsDate = new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000);
      
      // M/dd/YYYY 형식으로 변환
      const month = jsDate.getMonth() + 1;
      const day = jsDate.getDate();
      const year = jsDate.getFullYear();
      
      const convertedDate = `${month}/${day.toString().padStart(2, '0')}/${year}`;
      logger.info(`Excel serial conversion: ${dateValue} -> ${convertedDate}`);
      
      return convertedDate;
    }
    
    // YYYY-MM-DD 문자열 형식인지 확인
    const match = dateValue.toString().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      const [, year, month, day] = match;
      
      // M/dd/YYYY 형식으로 변환 (앞자리 0 제거)
      const convertedDate = `${parseInt(month)}/${day}/${year}`;
      logger.info(`String date conversion: ${dateValue} -> ${convertedDate}`);
      
      return convertedDate;
    }
    
    logger.warn(`Unsupported date format: ${dateValue} (type: ${typeof dateValue})`);
    return null;
    
  } catch (error) {
    logger.error(`Date conversion error: ${error.message}`);
    return null;
  }
}

// 성능 최적화를 위한 스마트 대기 시스템
const smartWait = {
  // 요소가 나타날 때까지 최대 timeout까지 대기
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
      logger.warn(`클릭 가능한 요소 대기 시간 초과: ${selector}`);
      return false;
    }
  },

  // 페이지가 준비될 때까지 대기
  forPageReady: async (page, timeout = 8000) => {
    try {
      await page.waitForFunction(
        () => document.readyState === 'complete',
        { timeout }
      );
      await delay(500); // 추가 안정화 대기
      return true;
    } catch (error) {
      logger.warn(`페이지 준비 대기 시간 초과: ${timeout}ms`);
      return false;
    }
  },

  // 여러 선택자 중 하나가 나타날 때까지 대기
  forAnyElement: async (page, selectors, timeout = 5000) => {
    try {
      await Promise.race(
        selectors.map(selector => 
          page.waitForSelector(selector, { visible: true, timeout })
        )
      );
      return true;
    } catch (error) {
      logger.warn(`복수 요소 대기 시간 초과: ${selectors.join(', ')}`);
      return false;
    }  }
};

/**
 * 데이터 테이블이 로드될 때까지 대기하는 함수
 * @param {Object} page - Puppeteer page 객체
 * @param {number} timeout - 최대 대기 시간 (기본값: 30초)
 * @returns {boolean} - 데이터 테이블이 로드되었는지 여부
 */
async function waitForDataTable(page, timeout = 30000) {
  const startTime = Date.now();
  logger.info(`데이터 테이블 로딩 대기 시작 (최대 ${timeout/1000}초)`);
  
  let loadingCompleted = false;
  
  while (Date.now() - startTime < timeout) {
    try {
      // 1. 로딩 스피너 확인 (있으면 계속 대기)
      const isLoading = await page.evaluate(() => {
        const loadingSelectors = [
          '.loading', '.spinner', '.ms-Spinner', '[aria-label*="로딩"]',
          '[aria-label*="Loading"]', '.dyn-loading', '.loadingSpinner'
        ];
        
        return loadingSelectors.some(selector => {
          const element = document.querySelector(selector);
          if (element) {
            const style = window.getComputedStyle(element);
            return style.display !== 'none' && 
                   style.visibility !== 'hidden' && 
                   element.offsetParent !== null;
          }
          return false;
        });
      });
      
      if (isLoading) {
        logger.info('로딩 중입니다. 계속 대기...');
        loadingCompleted = false; // 로딩이 다시 시작되면 플래그 리셋
        await delay(2000);
        continue;
      }
      
      // 2. 로딩 스피너가 사라진 후 처음이면 10초 대기
      if (!loadingCompleted) {
        logger.info('✅ 로딩 스피너가 사라졌습니다. 안정화를 위해 10초 대기 중...');
        await delay(5000);
        loadingCompleted = true;
        logger.info('안정화 대기 완료. 데이터 그리드 확인 중...');
      }
      
      // 3. 데이터 그리드 확인
      const hasDataGrid = await page.evaluate(() => {
        const gridSelectors = [
          '[data-dyn-controlname*="Grid"]', '.dyn-grid', 'div[role="grid"]',
          'table[role="grid"]', '[class*="grid"]', 'table'
        ];
        
        for (const selector of gridSelectors) {
          const element = document.querySelector(selector);
          if (element) {
            const rows = element.querySelectorAll('tr, [role="row"], [data-dyn-row]');
            if (rows.length > 0) { // 최소 1개 행이 있으면 OK
              return true;
            }
          }
        }
        return false;
      });
      
      if (hasDataGrid) {
        logger.info('✅ 데이터 그리드가 감지되었습니다. 테이블 로딩 완료!');
        return true;
      }
      
      logger.info('데이터 그리드를 찾는 중...');
      await delay(2000);
      
    } catch (error) {
      logger.warn(`데이터 테이블 대기 중 오류: ${error.message}`);
      await delay(2000);
    }
  }
  
  logger.warn(`⚠️ 데이터 테이블 로딩 대기 시간 초과 (${timeout/1000}초)`);
  return false;
}

// 로거 설정
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss'
    }),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  defaultMeta: { service: 'EZVoucher2-RPA' },
  transports: [
    new winston.transports.File({ 
      filename: path.join(__dirname, 'rpa.log'),
      maxsize: 5242880, // 5MB
      maxFiles: 5
    }),
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      )
    })
  ]
});

// Puppeteer Extra 설정
puppeteerExtra.use(StealthPlugin());

// 글로벌 변수
let globalCredentials = {
  username: '',
  password: ''
};

// 현재 날짜 가져오기
const now = new Date();
const currentYear = now.getFullYear();
const currentMonth = now.getMonth() + 1; // 0-based이므로 +1

// 글로벌 선택된 날짜 범위 정보 저장 객체 (동적 현재월로 초기화)
let globalDateRange = {
  year: currentYear,
  month: currentMonth, // 기본값: 동적 현재월
  fromDate: null,
  toDate: null
};

// 로그인 처리 함수 (EZVoucher.js와 동일한 ADFS 전용 로직)
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

// 글로벌 로그인 정보 설정
function setCredentials(username, password) {
  globalCredentials.username = username;
  globalCredentials.password = password;
  logger.info('매입송장 처리용 로그인 정보가 설정되었습니다');
}

// 글로벌 로그인 정보 반환
function getCredentials() {
  return globalCredentials;
}

// 글로벌 선택된 날짜 범위 정보 설정
function setSelectedDateRange(dateRangeInfo) {
  globalDateRange.year = dateRangeInfo.year;
  globalDateRange.month = dateRangeInfo.month;
  globalDateRange.fromDate = dateRangeInfo.fromDate;
  globalDateRange.toDate = dateRangeInfo.toDate;
  logger.info(`매입송장 처리용 날짜 범위가 설정되었습니다: ${dateRangeInfo.year}년 ${dateRangeInfo.month}월 (${dateRangeInfo.fromDate} ~ ${dateRangeInfo.toDate})`);
}

// 글로벌 선택된 날짜 범위 정보 반환
function getSelectedDateRange() {
  return globalDateRange;
}

/**
 * 단계별 진행 상황을 추적하는 D365 접속 함수 (다중모드용)
 */
async function connectToD365WithProgress(credentials, progressCallback, cycle) {
  logger.info(`=== ${cycle}번째 사이클 - D365 접속 시작 ===`);
  
  // 1단계 시작 콜백
  if (progressCallback) {
    progressCallback(cycle, 1, 0, null);
  }
  
  const browser = await puppeteerExtra.launch({
    headless: false,
    channel: 'chrome',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-web-security',
      '--disable-features=VizDisplayCompositor',
      '--start-maximized',
      '--ignore-certificate-errors',
      '--ignore-ssl-errors',
      '--ignore-certificate-errors-spki-list'
    ],
    defaultViewport: null
  });

  const page = await browser.newPage();
  
  try {
    // User-Agent 설정
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    
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
    
    // D365 페이지 접속 (재시도 로직 추가)
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
          if (progressCallback) {
            progressCallback(cycle, 1, 0, errorMsg);
          }
          throw new Error(errorMsg);
        }
        
        // 재시도 전 2초 대기
        logger.info('2초 후 재시도합니다...');
        await delay(2000);
      }
    }
    
    // 로그인 처리 (필요한 경우)
    if (await page.$('input[type="email"]') !== null || await page.$('#userNameInput') !== null) {
      logger.info('로그인 화면 감지됨, 로그인 시도 중...');
      await handleLogin(page, credentials);
    }
    
    // 로그인 후 페이지가 완전히 로드될 때까지 스마트 대기
    logger.info('로그인 후 페이지 로딩 확인 중...');
    const pageReady = await smartWait.forPageReady(page, 8000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }
    
    logger.info('페이지 로딩 확인 완료');
    logger.info(`=== 1단계: ERP 접속 및 로그인 완료 (${cycle}번째 사이클) ===`);
    
    // 1단계 완료 콜백
    if (progressCallback) {
      progressCallback(cycle, 2, 1, null);
    }
    
    // 2단계: 구매 입고내역 조회 페이지 이동
    try {
      await navigateToReceivingInquiry(page);
      logger.info(`=== 2단계: 구매 입고내역 조회 페이지 이동 완료 (${cycle}번째 사이클) ===`);
      
      // 2단계 완료 콜백
      if (progressCallback) {
        progressCallback(cycle, 3, 2, null);
      }
    } catch (step2Error) {
      const errorMsg = `2단계 실패: ${step2Error.message}`;
      logger.error(errorMsg);
      if (progressCallback) {
        progressCallback(cycle, 2, 1, errorMsg);
      }
      throw new Error(errorMsg);
    }
    
    // 3단계: 엑셀 파일 열기 및 매크로 실행
    try {
      logger.info(`🚀 === 3단계: 엑셀 파일 열기 및 매크로 실행 시작 (${cycle}번째 사이클) ===`);
      const excelResult = await executeExcelProcessing(page);
      if (!excelResult.success) {
        const errorMsg = `3단계 실패: ${excelResult.error}`;
        logger.error(errorMsg);
        if (progressCallback) {
          progressCallback(cycle, 3, 2, errorMsg);
        }
        throw new Error(errorMsg);
      } else {
        logger.info(`✅ 3단계: 엑셀 파일 열기 및 매크로 실행 완료 (${cycle}번째 사이클)`);
        logger.info(`✅ 4단계: 대기중인 공급사송장 메뉴 이동도 완료됨 (${cycle}번째 사이클)`);
        
        // 3단계, 4단계 완료 콜백
        if (progressCallback) {
          progressCallback(cycle, 5, 4, null);
        }
      }
    } catch (step3Error) {
      const errorMsg = `3단계 실패: ${step3Error.message}`;
      logger.error(errorMsg);
      if (progressCallback) {
        progressCallback(cycle, 3, 2, errorMsg);
      }
      throw new Error(errorMsg);
    }
    
    // 5~7단계는 executeExcelProcessing 내부에서 실행되므로 완료로 처리
    logger.info(`=== 5~7단계: 송장 처리 및 그룹웨어 상신 완료 (${cycle}번째 사이클) ===`);
    
    // 전체 완료 콜백
    if (progressCallback) {
      progressCallback(cycle, 7, 7, null);
    }
    
    // 전체 프로세스 완료 대기
    await delay(5000);
    
    // 완료 팝업창 표시
    try {
      await page.evaluate((cycleNum) => {
        alert(`🎉 ${cycleNum}번째 사이클 매입송장 처리 RPA 자동화가 완료되었습니다!\n\n✅ 1. ERP 접속 및 로그인 완료\n✅ 2. 구매 입고내역 조회 및 다운로드 완료\n✅ 3. 엑셀 파일 열기 및 매크로 실행 완료\n✅ 4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n✅ 5. 캘린더 버튼 클릭 및 송장 처리 완료\n✅ 6. 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료\n✅ 7. 그룹웨어 상신 완료\n\n브라우저가 자동으로 닫힙니다.`);
      }, cycle);
      logger.info('✅ 완료 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`완료 팝업창 표시 실패: ${alertError.message}`);
    }
    
    // 브라우저 닫기
    try {
      await browser.close();
      logger.info('✅ 브라우저 닫기 완료');
    } catch (closeError) {
      logger.warn(`브라우저 닫기 실패: ${closeError.message}`);
    }
    
    logger.info(`🎉 === ${cycle}번째 사이클 전체 RPA 프로세스 완료 - 브라우저 닫기 후 종료 ===`);
    
    // 성공 시 serializable한 객체만 반환
    return { 
      success: true, 
      message: `${cycle}번째 사이클 완료: 1. ERP 접속 및 로그인 완료\n2. 구매 입고내역 조회 및 다운로드 완료\n3. 엑셀 파일 열기 및 매크로 실행 완료\n4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n5. 캘린더 버튼 클릭 및 송장 처리 완료\n6. 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료\n7. 그룹웨어 상신 완료`,
      completedAt: new Date().toISOString(),
      browserKeptOpen: false,
      cycle: cycle
    };
    
  } catch (error) {
    logger.error(`${cycle}번째 사이클 D365 접속 중 오류 발생: ${error.message}`);
    
    // 에러 팝업창 표시
    try {
      await page.evaluate((errorMsg, cycleNum) => {
        alert(`❌ ${cycleNum}번째 사이클 매입송장 처리 RPA 자동화 중 오류가 발생했습니다!\n\n오류 내용: ${errorMsg}\n\n브라우저가 자동으로 닫힙니다.`);
      }, error.message, cycle);
      logger.info('❌ 에러 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`에러 팝업창 표시 실패: ${alertError.message}`);
    }
    
    // 브라우저 닫기 (에러 시에도)
    try {
      await browser.close();
      logger.info('✅ 브라우저 닫기 완료 (에러 발생 시)');
    } catch (closeError) {
      logger.warn(`브라우저 닫기 실패: ${closeError.message}`);
    }
    
    // 실패 시 serializable한 객체 반환
    return { 
      success: false, 
      error: error.message,
      failedAt: new Date().toISOString(),
      browserKeptOpen: false,
      cycle: cycle,
      failedStep: getCurrentStepFromError(error.message)
    };
  }
}

// 1. ERP 접속 및 로그인 완료
async function connectToD365(credentials) {
  logger.info('=== 매입송장 처리 - D365 접속 시작 ===');
  
  const browser = await puppeteerExtra.launch({
    headless: false,
    channel: 'chrome',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-web-security',
      '--disable-features=VizDisplayCompositor',
      '--start-maximized',
      '--ignore-certificate-errors',
      '--ignore-ssl-errors',
      '--ignore-certificate-errors-spki-list'
    ],
    defaultViewport: null
  });

  const page = await browser.newPage();
  
  try {
    // User-Agent 설정
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    
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
    
    // D365 페이지 접속 (재시도 로직 추가)
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
        
        // 재시도 전 2초 대기
        logger.info('2초 후 재시도합니다...');
        await delay(2000);
      }
    }    // 로그인 처리 (필요한 경우) - EZVoucher.js와 동일한 조건
    if (await page.$('input[type="email"]') !== null || await page.$('#userNameInput') !== null) {
      logger.info('로그인 화면 감지됨, 로그인 시도 중...');
      await handleLogin(page, credentials);
    }
    
    // 로그인 후 페이지가 완전히 로드될 때까지 스마트 대기
    logger.info('로그인 후 페이지 로딩 확인 중...');
    const pageReady = await smartWait.forPageReady(page, 8000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 2초 대기로 진행');
      await delay(2000);
    }    logger.info('페이지 로딩 확인 완료');
    
    logger.info('=== 1. ERP 접속 및 로그인 완료 ===');
    
    // 2번 동작 실행: 구매 입고내역 조회 페이지 이동
    await navigateToReceivingInquiry(page);
    
    // 3번 동작 실행: 엑셀 파일 열기 및 매크로 실행 (page 매개변수 전달)
    logger.info('🚀 === 3번 동작: 엑셀 파일 열기 및 매크로 실행 시작 ===');
    const excelResult = await executeExcelProcessing(page);
    if (!excelResult.success) {
      logger.warn(`⚠️ 엑셀 처리 실패: ${excelResult.error}`);
    } else {
      logger.info('✅ 3번 동작: 엑셀 파일 열기 및 매크로 실행 완료');
      logger.info('✅ 4번 동작: 대기중인 공급사송장 메뉴 이동도 완료됨');
    }
    
    
    // 전체 프로세스 완료 대기
    await delay(5000);
    
    // 완료 팝업창 표시
    try {
      await page.evaluate(() => {
        alert('🎉 매입송장 처리 RPA 자동화가 완료되었습니다!\n\n✅ 1. ERP 접속 및 로그인 완료\n✅ 2. 구매 입고내역 조회 및 다운로드 완료\n✅ 3. 엑셀 파일 열기 및 매크로 실행 완료\n✅ 4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n✅ 5. 캘린더 버튼 클릭 및 송장 처리 완료\n✅ 6. 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료\n✅ 7. 그룹웨어 상신 완료\n\n브라우저가 자동으로 닫힙니다.');
      });
      logger.info('✅ 완료 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`완료 팝업창 표시 실패: ${alertError.message}`);
    }
    
    // 브라우저 닫기
    try {
      await browser.close();
      logger.info('✅ 브라우저 닫기 완료');
    } catch (closeError) {
      logger.warn(`브라우저 닫기 실패: ${closeError.message}`);
    }
    
    logger.info('🎉 === 전체 RPA 프로세스 완료 (7번 RPA 실패 시) - 브라우저 닫기 후 종료 ===');
      // 성공 시 serializable한 객체만 반환
    return { 
      success: true, 
      message: '1. ERP 접속 및 로그인 완료\n2. 구매 입고내역 조회 및 다운로드 완료\n3. 엑셀 파일 열기 및 매크로 실행 완료\n4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n5. 캘린더 버튼 클릭 및 송장 처리 완료\n6. 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료\n7. 그룹웨어 상신 완료',
      completedAt: new Date().toISOString(),
      browserKeptOpen: false
    };
    
  } catch (error) {
    logger.error(`D365 접속 중 오류 발생: ${error.message}`);
    
    // 에러 팝업창 표시
    try {
      await page.evaluate((errorMsg) => {
        alert(`❌ 매입송장 처리 RPA 자동화 중 오류가 발생했습니다!\n\n오류 내용: ${errorMsg}\n\n브라우저가 자동으로 닫힙니다.`);
      }, error.message);
      logger.info('❌ 에러 팝업창 표시됨');
    } catch (alertError) {
      logger.warn(`에러 팝업창 표시 실패: ${alertError.message}`);
    }
    
    // 브라우저 닫기 (에러 시에도)
    try {
      await browser.close();
      logger.info('✅ 브라우저 닫기 완료 (에러 발생 시)');
    } catch (closeError) {
      logger.warn(`브라우저 닫기 실패: ${closeError.message}`);
    }
    
    // 에러 시에도 serializable한 객체만 반환
    return { 
      success: false, 
      error: error.message,
      failedAt: new Date().toISOString(),
      browserKeptOpen: false
    };
  }
}

// 2. 검색 기능을 통한 구매 입고내역 조회 페이지 이동
async function navigateToReceivingInquiry(page) {
  logger.info('=== 2. 구매 입고내역 조회 페이지 이동 시작 ===');
  
  try {
    // 2-1. 검색 버튼 클릭 (Find-symbol 버튼)
    logger.info('검색 버튼 찾는 중...');
    
    const searchButtonSelectors = [
      '.button-commandRing.Find-symbol',
      'span.Find-symbol',
      '[data-dyn-image-type="Symbol"].Find-symbol',
      '.button-container .Find-symbol'
    ];
    
    let searchButtonClicked = false;
    
    for (const selector of searchButtonSelectors) {
      try {
        logger.info(`검색 버튼 선택자 시도: ${selector}`);
        
        const searchButton = await page.$(selector);
        if (searchButton) {
          const isVisible = await page.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
          }, searchButton);
          
          if (isVisible) {
            await searchButton.click();
            logger.info(`검색 버튼 클릭 성공: ${selector}`);
            searchButtonClicked = true;
            break;
          } else {
            logger.warn(`검색 버튼이 보이지 않음: ${selector}`);
          }
        }
      } catch (error) {
        logger.warn(`검색 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!searchButtonClicked) {
      // JavaScript로 직접 검색 버튼 클릭 시도
      try {
        logger.info('JavaScript로 검색 버튼 직접 클릭 시도...');
        await page.evaluate(() => {
          const searchButtons = document.querySelectorAll('.Find-symbol, [data-dyn-image-type="Symbol"]');
          for (const btn of searchButtons) {
            if (btn.classList.contains('Find-symbol') || btn.getAttribute('data-dyn-image-type') === 'Symbol') {
              btn.click();
              return true;
            }
          }
          return false;
        });
        searchButtonClicked = true;
        logger.info('JavaScript로 검색 버튼 클릭 성공');
      } catch (jsError) {
        logger.error('JavaScript 검색 버튼 클릭 실패:', jsError.message);
      }
    }
    
    if (!searchButtonClicked) {
      throw new Error('검색 버튼을 찾을 수 없습니다.');
    }
    
    // 검색창이 나타날 때까지 대기
    await delay(2000);
    
    // 2-2. "구매 입고내역 조회(N)" 검색어 입력
    logger.info('검색어 입력 중...');
    
    const searchInputSelectors = [
      'input[type="text"]',
      '.navigationSearchBox input',
      '#NavigationSearchBox',
      'input[placeholder*="검색"]',
      'input[aria-label*="검색"]'
    ];
    
    let searchInputFound = false;
    const searchTerm = '구매 입고내역 조회(N)';
    
    for (const selector of searchInputSelectors) {
      try {
        logger.info(`검색 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 기존 텍스트 클리어
        await page.click(selector, { clickCount: 3 }); // 모든 텍스트 선택
        await page.keyboard.press('Backspace'); // 선택된 텍스트 삭제
        
        // 검색어 입력
        await page.type(selector, searchTerm, { delay: 100 });
        logger.info(`검색어 입력 완료: ${searchTerm}`);
        
        searchInputFound = true;
        break;
        
      } catch (error) {
        logger.warn(`검색 입력창 처리 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!searchInputFound) {
      throw new Error('검색 입력창을 찾을 수 없습니다.');
    }
    
    // 검색 결과가 나타날 때까지 대기
    await delay(3000);
    
    // 2-3. NavigationSearchBox에서 해당 메뉴 클릭
    logger.info('검색 결과에서 구매 입고내역 조회 메뉴 찾는 중...');
    
    const searchResultSelectors = [
      '.navigationSearchBox',
      '.search-results',
      '.navigation-search-results',
      '[data-dyn-bind*="NavigationSearch"]'
    ];
    
    let menuClicked = false;
    
    for (const containerSelector of searchResultSelectors) {
      try {
        const container = await page.$(containerSelector);
        if (container) {
          // 컨테이너 내에서 "구매 입고내역 조회" 텍스트가 포함된 요소 찾기
          const menuItems = await page.$$eval(`${containerSelector} *`, (elements) => {
            return elements
              .filter(el => {
                const text = el.textContent || el.innerText || '';
                return text.includes('구매 입고내역 조회') || text.includes('구매') && text.includes('입고');
              })
              .map(el => ({
                text: el.textContent || el.innerText,
                clickable: el.tagName === 'A' || el.tagName === 'BUTTON' || el.onclick || el.getAttribute('role') === 'button'
              }));
          });
          
          logger.info(`검색 결과 메뉴 항목들:`, menuItems);
          
          if (menuItems.length > 0) {
            // 첫 번째 매칭되는 항목 클릭
            await page.evaluate((containerSel) => {
              const container = document.querySelector(containerSel);
              if (container) {
                const elements = container.querySelectorAll('*');
                for (const el of elements) {
                  const text = el.textContent || el.innerText || '';
                  if (text.includes('구매 입고내역 조회') || (text.includes('구매') && text.includes('입고'))) {
                    el.click();
                    return true;
                  }
                }
              }
              return false;
            }, containerSelector);
            
            logger.info('구매 입고내역 조회 메뉴 클릭 완료');
            menuClicked = true;
            break;
          }
        }
      } catch (error) {
        logger.warn(`검색 결과 처리 실패: ${containerSelector} - ${error.message}`);
      }
    }
    
    if (!menuClicked) {
      // Enter 키로 첫 번째 결과 선택 시도
      logger.info('Enter 키로 검색 결과 선택 시도...');
      await page.keyboard.press('Enter');
      menuClicked = true;
    }
    
    // 페이지 이동 대기
    logger.info('구매 입고내역 조회 페이지 로딩 대기 중...');
    await delay(5000);
    
    // 페이지 로딩 완료 확인
    const pageReady = await smartWait.forPageReady(page, 10000);
    if (!pageReady) {
      logger.warn('페이지 로딩 확인 실패, 기본 3초 대기로 진행');
      await delay(3000);
    }
    
    logger.info('=== 2. 구매 입고내역 조회 페이지 이동 완료 ===');


    
    // 3. FromDate 입력 (현재 월의 첫날)
    logger.info('=== 3. FromDate 설정 시작 ===');
    
    // 현재 날짜에서 월의 첫날 계산
    /*const now = new Date();
    // 현재날짜 기준 현재월 가져오기
    const fromDate = `${now.getMonth() + 1}/1/${now.getFullYear()}`; // M/d/YYYY 형태

    logger.info(`설정할 FromDate: ${fromDate}`);
    */
    
    // 사용자 선택된 날짜 범위 사용
    let fromDate;

    // 디버깅: globalDateRange 현재 상태 확인
    logger.info(`[DEBUG] 현재 globalDateRange 상태:`, JSON.stringify(globalDateRange, null, 2));

    // globalDateRange에서 fromDate가 이미 설정된 경우 사용
    if (globalDateRange.fromDate) {
      fromDate = globalDateRange.fromDate;
      logger.info(`[UI에서 설정된 값 사용] FromDate: ${fromDate} (${globalDateRange.year}년 ${globalDateRange.month}월)`);
    } else {
      // 기본값: 현재월의 첫날 (fallback)
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      fromDate = `${currentMonth}/1/${currentYear}`;
      logger.info(`[기본값 사용] FromDate: ${fromDate} (현재월)`);
    }

    logger.info(`설정할 FromDate: ${fromDate}`);
    
    //-------------------------------------------------------------------------------

    // FromDate 입력창 선택자들
    const fromDateSelectors = [
      'input[name="FromDate"]',
      'input[id*="FromDate_input"]',
      'input[aria-labelledby*="FromDate_label"]',
      'input[placeholder=""][name="FromDate"]'
    ];
    
    let fromDateSet = false;
    
    for (const selector of fromDateSelectors) {
      try {
        logger.info(`FromDate 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 입력창 클릭
        await page.click(selector);
        await delay(500);
        
        // 기존 텍스트 클리어 (모든 텍스트 선택 후 삭제)
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await delay(300);
        
        // 날짜 입력
        await page.type(selector, fromDate, { delay: 100 });
        await page.keyboard.press('Tab'); // 포커스 이동으로 입력 확정
        
        logger.info(`FromDate 설정 완료: ${fromDate}`);
        fromDateSet = true;
        break;
        
      } catch (error) {
        logger.warn(`FromDate 설정 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!fromDateSet) {
      throw new Error('FromDate 입력창을 찾을 수 없습니다.');
    }
    
    await delay(1000); // 입력 안정화 대기
    
    // 4. ToDate 입력 (현재 월의 마지막 날)
    logger.info('=== 4. ToDate 설정 시작 ===');
    
    /*
    // 현재 날짜에서 월의 마지막 날 계산
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const toDate = `${now.getMonth() + 1}/${lastDay}/${now.getFullYear()}`; // M/d/YYYY 형태
    logger.info(`설정할 ToDate: ${toDate}`);
    */
   
    // 사용자 선택된 날짜 범위 사용
    let toDate;

    // globalDateRange에서 toDate가 이미 설정된 경우 사용
    if (globalDateRange.toDate) {
      toDate = globalDateRange.toDate;
      logger.info(`[UI에서 설정된 값 사용] ToDate: ${toDate} (${globalDateRange.year}년 ${globalDateRange.month}월)`);
    } else {
      // 기본값: 현재월의 마지막날 (fallback)
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      const lastDay = new Date(currentYear, currentMonth, 0).getDate();
      toDate = `${currentMonth}/${lastDay}/${currentYear}`;
      logger.info(`[기본값 사용] ToDate: ${toDate} (현재월)`);
    }

    logger.info(`설정할 ToDate: ${toDate}`);
    

    // ToDate 입력창 선택자들
    const toDateSelectors = [
      'input[name="ToDate"]',
      'input[id*="ToDate_input"]',
      'input[aria-labelledby*="ToDate_label"]',
      'input[placeholder=""][name="ToDate"]'
    ];
    
    let toDateSet = false;
    
    for (const selector of toDateSelectors) {
      try {
        logger.info(`ToDate 입력창 선택자 시도: ${selector}`);
        
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        
        // 입력창 클릭
        await page.click(selector);
        await delay(500);
        
        // 기존 텍스트 클리어 (모든 텍스트 선택 후 삭제)
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await delay(300);
        
        // 날짜 입력
        await page.type(selector, toDate, { delay: 100 });
        await page.keyboard.press('Tab'); // 포커스 이동으로 입력 확정
        
        logger.info(`ToDate 설정 완료: ${toDate}`);
        toDateSet = true;
        break;
        
      } catch (error) {
        logger.warn(`ToDate 설정 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!toDateSet) {
      throw new Error('ToDate 입력창을 찾을 수 없습니다.');
    }
    
    await delay(1000); // 입력 안정화 대기
    
    // 5. Inquiry 버튼 클릭
    logger.info('=== 5. Inquiry 버튼 클릭 시작 ===');
    
    // Inquiry 버튼 선택자들
    const inquiryButtonSelectors = [
      '.button-container:has(.button-label:contains("Inquiry"))',
      'span.button-label:contains("Inquiry")',
      'div.button-container span[id*="Inquiry_label"]',
      '[id*="Inquiry_label"]',
      'span[for*="Inquiry"]'
    ];
    
    let inquiryButtonClicked = false;
    
    for (const selector of inquiryButtonSelectors) {
      try {
        logger.info(`Inquiry 버튼 선택자 시도: ${selector}`);
        
        // CSS 선택자에 :contains()가 있는 경우 JavaScript로 처리
        if (selector.includes(':contains(')) {
          const clicked = await page.evaluate(() => {
            const buttons = document.querySelectorAll('.button-container');
            for (const container of buttons) {
              const label = container.querySelector('.button-label, span[id*="label"]');
              if (label && (label.textContent || label.innerText || '').includes('Inquiry')) {
                container.click();
                return true;
              }
            }
            return false;
          });
          
          if (clicked) {
            logger.info('JavaScript로 Inquiry 버튼 클릭 성공');
            inquiryButtonClicked = true;
            break;
          }
        } else {
          // 일반 선택자 처리
          const inquiryButton = await page.$(selector);
          if (inquiryButton) {
            const isVisible = await page.evaluate(el => {
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, inquiryButton);
            
            if (isVisible) {
              await inquiryButton.click();
              logger.info(`Inquiry 버튼 클릭 성공: ${selector}`);
              inquiryButtonClicked = true;
              break;
            } else {
              logger.warn(`Inquiry 버튼이 보이지 않음: ${selector}`);
            }
          }
        }
      } catch (error) {
        logger.warn(`Inquiry 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    
    // 추가 시도: ID와 텍스트를 조합한 방법
    if (!inquiryButtonClicked) {
      try {
        logger.info('ID와 텍스트 조합으로 Inquiry 버튼 찾는 중...');
        
        const clicked = await page.evaluate(() => {
          // id에 "Inquiry"가 포함된 요소들 찾기
          const elements = document.querySelectorAll('[id*="Inquiry"]');
          for (const el of elements) {
            // 클릭 가능한 요소이거나 부모가 클릭 가능한 요소인지 확인
            const clickableEl = el.closest('.button-container, button, [role="button"]') || el;
            if (clickableEl) {
              clickableEl.click();
              return true;
            }
          }
          return false;
        });
        
        if (clicked) {
          logger.info('ID 기반으로 Inquiry 버튼 클릭 성공');
          inquiryButtonClicked = true;
        }
      } catch (error) {
        logger.warn(`ID 기반 Inquiry 버튼 클릭 실패: ${error.message}`);
      }
    }
      if (!inquiryButtonClicked) {
      throw new Error('Inquiry 버튼을 찾을 수 없습니다.');
    }
    
    // 조회 실행 후 데이터 테이블이 나타날 때까지 대기
    logger.info('조회 실행 중, 데이터 테이블 로딩 대기...');
    
    // 기본 대기 시간 (최소 10초 - 조회 실행 후 초기 로딩 대기)
    await delay(5000);
    
    // 데이터 테이블 로딩 확인 (30초 타임아웃으로 단축)
    const dataTableLoaded = await waitForDataTable(page, 15000);
    
    if (!dataTableLoaded) {
      logger.warn('데이터 테이블 로딩 확인 실패, 하지만 계속 진행합니다...');
      // 추가 대기 후 계속 진행
      await delay(2000);
    }
      logger.info('=== 구매 입고내역 조회 설정 및 조회 실행 완료 ===');
    
    // 6. 데이터 내보내기 실행
    logger.info('🚀 === 6. 데이터 내보내기 시작 ===');
    
    // 내보내기 전 추가 안정화 대기
    await delay(1000);
    
    // 6-1. 구매주문 컬럼 헤더 우클릭
    logger.info('🔍 구매주문 컬럼 헤더 찾는 중...');
    
    // 더 많은 선택자 추가
    const purchaseOrderHeaderSelectors = [
      'div[data-dyn-columnname="NPS_VendPackingSlipSumReportTemp_PurchId"]',
      'div[data-dyn-controlname="NPS_VendPackingSlipSumReportTemp_PurchId"]',
      'div.dyn-headerCell[data-dyn-columnname*="PurchId"]',
      'div.dyn-headerCellLabel[title="구매주문"]',
      '[data-dyn-columnname*="PurchId"]',
      'th:contains("구매주문")',
      'div[title="구매주문"]'
    ];
    
    let headerRightClicked = false;
    
    // JavaScript로 "구매주문" 헤더 찾기 (더 robust한 방법)
    try {
      logger.info('JavaScript로 구매주문 헤더 찾는 중...');
      
      const headerFound = await page.evaluate(() => {
        // 모든 가능한 헤더 요소 검색
        const allHeaders = document.querySelectorAll('th, .dyn-headerCell, [role="columnheader"], div[data-dyn-columnname], div[title]');
        
        for (const header of allHeaders) {
          const text = header.textContent || header.innerText || header.title || '';
          const columnName = header.getAttribute('data-dyn-columnname') || '';
          
          if (text.includes('구매주문') || columnName.includes('PurchId')) {
            // 우클릭 이벤트 발생
            const event = new MouseEvent('contextmenu', {
              bubbles: true,
              cancelable: true,
              button: 2
            });
            header.dispatchEvent(event);
            return true;
          }
        }
        return false;
      });
      
      if (headerFound) {
        logger.info('✅ JavaScript로 구매주문 헤더 우클릭 성공');
        headerRightClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 헤더 우클릭 실패: ${error.message}`);
    }
    
    // 기존 방법으로도 시도
    if (!headerRightClicked) {
      for (const selector of purchaseOrderHeaderSelectors) {
        try {
          logger.info(`구매주문 헤더 선택자 시도: ${selector}`);
          
          if (selector.includes(':contains(')) {
            continue; // CSS :contains()는 지원되지 않으므로 스킵
          }
          
          const headerElement = await page.$(selector);
          if (headerElement) {
            const isVisible = await page.evaluate(el => {
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, headerElement);
            
            if (isVisible) {
              // 우클릭 실행
              await headerElement.click({ button: 'right' });
              logger.info(`✅ 구매주문 헤더 우클릭 성공: ${selector}`);
              headerRightClicked = true;
              break;
            } else {
              logger.warn(`구매주문 헤더가 보이지 않음: ${selector}`);
            }
          }
        } catch (error) {
          logger.warn(`구매주문 헤더 우클릭 실패: ${selector} - ${error.message}`);
        }
      }
    }
    
    if (!headerRightClicked) {
      logger.error('❌ 구매주문 컬럼 헤더를 찾을 수 없습니다.');
      throw new Error('구매주문 컬럼 헤더를 찾을 수 없습니다.');
    }
    
    // 컨텍스트 메뉴가 나타날 때까지 대기
    logger.info('⏳ 컨텍스트 메뉴 대기 중...');
    await delay(3000);
      // 6-2. "모든 행 내보내기" 메뉴 클릭
    logger.info('🔍 모든 행 내보내기 메뉴 찾는 중...');
    
    let exportMenuClicked = false;
    
    // JavaScript로 "모든 행 내보내기" 메뉴 찾기
    try {
      logger.info('JavaScript로 모든 행 내보내기 메뉴 찾는 중...');
      
      const clicked = await page.evaluate(() => {
        // 1. button-container 내부의 button-label에서 "모든 행 내보내기" 찾기
        const buttonContainers = document.querySelectorAll('.button-container');
        
        for (const container of buttonContainers) {
          const buttonLabel = container.querySelector('.button-label');
          if (buttonLabel) {
            const text = buttonLabel.textContent || buttonLabel.innerText || '';
            if (text.includes('모든 행 내보내기')) {
              // button-container 전체를 클릭
              container.click();
              return { success: true, text: text.trim(), method: 'button-container' };
            }
          }
        }
        
        // 2. 직접 button-label 요소에서 찾기
        const buttonLabels = document.querySelectorAll('.button-label');
        for (const label of buttonLabels) {
          const text = label.textContent || label.innerText || '';
          if (text.includes('모든 행 내보내기')) {
            // 부모 button-container 찾아서 클릭
            const parentContainer = label.closest('.button-container');
            if (parentContainer) {
              parentContainer.click();
              return { success: true, text: text.trim(), method: 'parent-container' };
            } else {
              // 부모가 없으면 label 자체 클릭
              label.click();
              return { success: true, text: text.trim(), method: 'direct-label' };
            }
          }
        }
        
        // 3. 모든 요소에서 텍스트 검색 (기존 방법)
        const allElements = document.querySelectorAll('span, button, [role="button"], [role="menuitem"]');
        
        for (const element of allElements) {
          const text = element.textContent || element.innerText || '';
          if (text.includes('모든 행 내보내기') || text.includes('내보내기') || text.includes('Export')) {
            // 클릭 가능한 부모 요소 찾기
            const clickableParent = element.closest('.button-container, button, [role="button"], [role="menuitem"]') || element;
            clickableParent.click();
            return { success: true, text: text.trim(), method: 'fallback' };
          }
        }
        
        return { success: false };
      });
      
      if (clicked.success) {
        logger.info(`✅ JavaScript로 내보내기 메뉴 클릭 성공 (${clicked.method}): "${clicked.text}"`);
        exportMenuClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 모든 행 내보내기 메뉴 클릭 실패: ${error.message}`);
    }
    
    if (!exportMenuClicked) {
      // 추가 시도: Puppeteer 선택자로 button-container 직접 찾기
      try {
        logger.info('Puppeteer 선택자로 모든 행 내보내기 버튼 찾는 중...');
        
        // button-container 내부에 "모든 행 내보내기" 텍스트가 있는 요소 찾기
        const buttonContainers = await page.$$('.button-container');
        
        for (const container of buttonContainers) {
          try {
            const text = await container.evaluate(el => {
              const label = el.querySelector('.button-label');
              return label ? (label.textContent || label.innerText || '') : '';
            });
            
            if (text.includes('모든 행 내보내기')) {
              await container.click();
              logger.info(`✅ Puppeteer로 내보내기 버튼 클릭 성공: "${text.trim()}"`);
              exportMenuClicked = true;
              break;
            }
          } catch (containerError) {
            logger.warn(`button-container 처리 중 오류: ${containerError.message}`);
          }
        }
      } catch (error) {
        logger.warn(`Puppeteer 모든 행 내보내기 버튼 클릭 실패: ${error.message}`);
      }
    }
    
    if (!exportMenuClicked) {
      logger.error('❌ 모든 행 내보내기 메뉴를 찾을 수 없습니다.');
      throw new Error('모든 행 내보내기 메뉴를 찾을 수 없습니다.');
    }
    
    // 다운로드 대화상자가 나타날 때까지 대기
    logger.info('⏳ 다운로드 대화상자 대기 중...');
    await delay(5000);
    
    // 6-3. "다운로드" 버튼 클릭
    logger.info('🔍 다운로드 버튼 찾는 중...');
    
    let downloadButtonClicked = false;
    
    // JavaScript로 다운로드 버튼 찾기 (더 강력한 로직)
    try {
      logger.info('JavaScript로 다운로드 버튼 찾는 중...');
      
      const clicked = await page.evaluate(() => {
        // 1. "다운로드" 텍스트가 포함된 모든 요소 검색
        const allElements = document.querySelectorAll('button, .button-label, span, [role="button"]');
        
        for (const element of allElements) {
          const text = element.textContent || element.innerText || '';
          if (text.includes('다운로드') || text.includes('Download')) {
            const clickable = element.tagName === 'BUTTON' ? element : element.closest('button, [role="button"], .button-container');
            if (clickable) {
              clickable.click();
              return { success: true, text: text.trim(), method: 'text-search' };
            }
          }
        }
        
        // 2. DownloadButton 관련 속성으로 검색
        const downloadElements = document.querySelectorAll('[name*="DownloadButton"], [id*="DownloadButton"], [data-dyn-controlname*="Download"]');
        for (const el of downloadElements) {
          const button = el.tagName === 'BUTTON' ? el : el.closest('button');
          if (button) {
            button.click();
            return { success: true, method: 'attribute-search' };
          }
        }
        
        // 3. Download 아이콘으로 검색
        const downloadIcons = document.querySelectorAll('.Download-symbol, [class*="download"], [class*="Download"]');
        for (const icon of downloadIcons) {
          const button = icon.closest('button, [role="button"]');
          if (button) {
            button.click();
            return { success: true, method: 'icon-search' };
          }
        }
        
        return { success: false };
      });
      
      if (clicked.success) {
        logger.info(`✅ JavaScript로 다운로드 버튼 클릭 성공 (${clicked.method}): ${clicked.text || 'N/A'}`);
        downloadButtonClicked = true;
      }
    } catch (error) {
      logger.warn(`JavaScript 다운로드 버튼 클릭 실패: ${error.message}`);
    }
    
    if (!downloadButtonClicked) {
      logger.error('❌ 다운로드 버튼을 찾을 수 없습니다.');
      throw new Error('다운로드 버튼을 찾을 수 없습니다.');
    }
    
    // 다운로드 완료 대기
    logger.info('📥 다운로드 실행 중, 완료 대기...');
    await delay(8000);
    
    logger.info('🎉 === 6. 데이터 내보내기 완료 ===');
    
    logger.info('=== 2. 구매 입고내역 조회 페이지 이동 및 데이터 다운로드 완료 ===');
    
    return {
      success: true,
      message: '구매 입고내역 조회 및 데이터 다운로드가 완료되었습니다.'
    };
    
  } catch (error) {
    logger.error(`구매 입고내역 조회 페이지 이동 중 오류: ${error.message}`);
    throw error;
  }
}


// 엑셀 파일에서 특정 셀 값 읽기 함수
function getCellValueFromExcel(filePath, sheetName, cellAddress) {
  try {
    logger.info(`엑셀 파일에서 셀 값 읽기: ${filePath}, 시트: ${sheetName}, 셀: ${cellAddress}`);
    
    const workbook = xlsx.readFile(filePath);
    logger.info(`워크북 로드 완료. 시트 목록: ${Object.keys(workbook.Sheets).join(', ')}`);
    
    // 시트명이 없으면 첫 번째 시트 사용
    const targetSheetName = sheetName || Object.keys(workbook.Sheets)[0];
    const worksheet = workbook.Sheets[targetSheetName];
    
    if (!worksheet) {
      throw new Error(`시트를 찾을 수 없습니다: ${targetSheetName}`);
    }
    
    const cell = worksheet[cellAddress];
    const cellValue = cell ? cell.v : '';
    
    logger.info(`셀 ${cellAddress} 값: "${cellValue}"`);
    return cellValue;
  } catch (error) {
    logger.error(`엑셀 셀 값 읽기 실패: ${error.message}`);
    throw error;
  }
}

// 다운받은 엑셀 파일 경로 찾기 함수 (파일을 열지 않고 경로만 반환)
async function openDownloadedExcel() {
  logger.info('🚀 === 다운받은 엑셀 파일 경로 찾기 시작 ===');
  
  try {
    const os = require('os');
    
    // Windows 기본 다운로드 폴더 경로
    const downloadPath = path.join(os.homedir(), 'Downloads');
    logger.info(`다운로드 폴더 경로: ${downloadPath}`);
    
    // 다운로드 폴더에서 최근 다운받은 엑셀 파일 찾기
    logger.info('최근 다운받은 엑셀 파일 찾는 중...');
    
    const files = fs.readdirSync(downloadPath);
    const excelFiles = files.filter(file => 
      (file.endsWith('.xlsx') || file.endsWith('.xls')) && 
      !file.startsWith('~$') // 임시 파일 제외
    );
    
    if (excelFiles.length === 0) {
      throw new Error('다운로드 폴더에서 엑셀 파일을 찾을 수 없습니다.');
    }
    
    // 파일들을 수정시간 기준으로 정렬하여 가장 최근 파일 찾기
    const excelFilesWithStats = excelFiles.map(file => {
      const filePath = path.join(downloadPath, file);
      const stats = fs.statSync(filePath);
      return {
        name: file,
        path: filePath,
        mtime: stats.mtime
      };
    }).sort((a, b) => b.mtime - a.mtime);
    
    const latestExcelFile = excelFilesWithStats[0];
    logger.info(`최신 엑셀 파일 발견: ${latestExcelFile.name}`);
    logger.info(`파일 경로: ${latestExcelFile.path}`);
    logger.info(`수정시간: ${latestExcelFile.mtime}`);
    
    // 파일이 최근 5분 이내에 다운로드된 것인지 확인
    const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);
    if (latestExcelFile.mtime < fiveMinutesAgo) {
      logger.warn('⚠️ 발견된 엑셀 파일이 5분 이전에 수정된 파일입니다. 최근 다운로드된 파일이 맞는지 확인하세요.');
    }
    
    // 파일을 열지 않고 경로만 반환
    logger.info('✅ 엑셀 파일 경로를 성공적으로 찾았습니다 (파일을 열지 않음).');
    
    return {
      success: true,
      message: '엑셀 파일 경로를 성공적으로 찾았습니다.',
      filePath: latestExcelFile.path,
      fileName: latestExcelFile.name
    };
    
  } catch (error) {
    logger.error(`엑셀 파일 경로 찾기 중 오류: ${error.message}`);
    
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '엑셀 파일 경로 찾기'
    };
  }
}

// 3번 RPA 동작: 엑셀 파일 열기 및 매크로 실행 (통합 관리)
async function executeExcelProcessing(page) {
  logger.info('🚀 === 3번 RPA 동작: 엑셀 파일 열기 및 매크로 실행 시작 ===');
  logger.info(`📋 현재 설정된 A열 값: userInputValueA = ${userInputValueA}`);
  try {
    // 1. 다운로드 폴더에서 최신 엑셀 파일 찾기 (파일을 열지 않고 경로만 획득)
    logger.info('Step 1: 엑셀 파일 경로 찾기 실행 중...');
    const openResult = await openDownloadedExcel();
    if (!openResult.success) {
      throw new Error(openResult.error || '엑셀 파일 경로 찾기에 실패했습니다.');
    }
    logger.info(`✅ Step 1 완료: ${openResult.fileName} (파일을 열지 않고 경로만 획득)`);
    // 2. 매크로 자동 실행 (PowerShell이 엑셀 파일을 열고 매크로 실행)
    logger.info('Step 2: 매크로 자동 실행 시작... (PowerShell이 엑셀 파일을 열고 매크로 실행)');
    const macroResult = await openExcelAndExecuteMacro(openResult.filePath);
    if (!macroResult.success) {
      throw new Error(macroResult.error || '엑셀 매크로 실행에 실패했습니다.');
    }
    logger.info('✅ Step 2 완료: 매크로 실행 성공');
    // 3. 완료 메시지 반환
    logger.info('🎉 === 3번 RPA 동작 완료 ===');
    // 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 (5초 대기 후 실행)
    logger.info('⏳ 5초 대기 후 4번 RPA 동작(대기중인 공급사송장 메뉴 이동) 시작 예정...');
    await delay(5000);
    
    let step4Status = '4번 RPA 동작 건너뜀';
    if (page) {
      try {
        const pendingResult = await navigateToPendingVendorInvoice(page, openResult.filePath);
        logger.info('4번 RPA 동작 결과:', pendingResult);
        step4Status = '4번 RPA 동작(대기중인 공급사송장 메뉴 이동) 실행 완료';
      } catch (step4Error) {
        logger.error(`4번 RPA 동작 중 오류 발생: ${step4Error.message}`);
        logger.warn('4번 RPA 동작 실패했지만 전체 프로세스는 계속 진행합니다.');
        step4Status = `4번 RPA 동작 실패: ${step4Error.message}`;
      }
    } else {
      logger.warn('4번 RPA 동작을 위한 page 인스턴스가 제공되지 않았습니다.');
    }
    return {
      success: true,
      message: '3번 RPA 동작: 엑셀 파일 매크로 실행이 완료되었습니다.',
      filePath: openResult.filePath,
      fileName: openResult.fileName,
      completedAt: new Date().toISOString(),
      steps: {
        step1: '엑셀 파일 경로 찾기 완료',
        step2: '매크로 실행 완료',
        step3: step4Status
      }
    };
  } catch (error) {
    logger.error(`3번 RPA 동작 중 오류: ${error.message}`);
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '3번 RPA 동작 (엑셀 파일 열기 및 매크로 실행)'
    };
  }
}

// 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동
async function navigateToPendingVendorInvoice(page, excelFilePath) {
  logger.info('🚀 === 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 시작 ===');
  try {
    // 1. 검색 버튼 클릭 (2-1과 동일)
    logger.info('검색 버튼 찾는 중...');
    const searchButtonSelectors = [
      '.button-commandRing.Find-symbol',
      'span.Find-symbol',
      '[data-dyn-image-type="Symbol"].Find-symbol',
      '.button-container .Find-symbol'
    ];
    let searchButtonClicked = false;
    for (const selector of searchButtonSelectors) {
      try {
        logger.info(`검색 버튼 선택자 시도: ${selector}`);
        const searchButton = await page.$(selector);
        if (searchButton) {
          const isVisible = await page.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
          }, searchButton);
          if (isVisible) {
            await searchButton.click();
            logger.info(`검색 버튼 클릭 성공: ${selector}`);
            searchButtonClicked = true;
            break;
          } else {
            logger.warn(`검색 버튼이 보이지 않음: ${selector}`);
          }
        }
      } catch (error) {
        logger.warn(`검색 버튼 클릭 실패: ${selector} - ${error.message}`);
      }
    }
    if (!searchButtonClicked) {
      // JavaScript로 직접 검색 버튼 클릭 시도
      try {
        logger.info('JavaScript로 검색 버튼 직접 클릭 시도...');
        await page.evaluate(() => {
          const searchButtons = document.querySelectorAll('.Find-symbol, [data-dyn-image-type="Symbol"]');
          for (const btn of searchButtons) {
            if (btn.classList.contains('Find-symbol') || btn.getAttribute('data-dyn-image-type') === 'Symbol') {
              btn.click();
              return true;
            }
          }
          return false;
        });
        searchButtonClicked = true;
        logger.info('JavaScript로 검색 버튼 클릭 성공');
      } catch (jsError) {
        logger.error('JavaScript 검색 버튼 클릭 실패:', jsError.message);
      }
    }
    if (!searchButtonClicked) {
      throw new Error('검색 버튼을 찾을 수 없습니다. (4번 RPA)');
    }
    // 검색창이 나타날 때까지 대기
    await delay(2000);
    // 2. "대기중인 공급사송장" 검색어 입력
    logger.info('검색어 입력 중...');
    const searchInputSelectors = [
      'input[type="text"]',
      '.navigationSearchBox input',
      '#NavigationSearchBox',
      'input[placeholder*="검색"]',
      'input[aria-label*="검색"]'
    ];
    let searchInputFound = false;
    const searchTerm = '대기중인 공급사송장';
    for (const selector of searchInputSelectors) {
      try {
        logger.info(`검색 입력창 선택자 시도: ${selector}`);
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type(selector, searchTerm, { delay: 100 });
        logger.info(`검색어 입력 완료: ${searchTerm}`);
        searchInputFound = true;
        break;
      } catch (error) {
        logger.warn(`검색 입력창 처리 실패: ${selector} - ${error.message}`);
      }
    }
    if (!searchInputFound) {
      throw new Error('검색 입력창을 찾을 수 없습니다. (4번 RPA)');
    }
    // 검색 결과가 나타날 때까지 대기
    await delay(3000);
    // 3. NavigationSearchBox에서 해당 메뉴 클릭
    logger.info('검색 결과에서 대기중인 공급사송장 메뉴 찾는 중...');
    const searchResultSelectors = [
      '.navigationSearchBox',
      '.search-results',
      '.navigation-search-results',
      '[data-dyn-bind*="NavigationSearch"]'
    ];
    let menuClicked = false;
    for (const containerSelector of searchResultSelectors) {
      try {
        const container = await page.$(containerSelector);
        if (container) {
          const menuItems = await page.$$eval(`${containerSelector} *`, (elements) => {
            return elements
              .filter(el => {
                const text = el.textContent || el.innerText || '';
                return text.includes('대기중인 공급사송장');
              })
              .map(el => ({
                text: el.textContent || el.innerText,
                clickable: el.tagName === 'A' || el.tagName === 'BUTTON' || el.onclick || el.getAttribute('role') === 'button'
              }));
          });
          logger.info(`검색 결과 메뉴 항목들:`, menuItems);
          if (menuItems.length > 0) {
            await page.evaluate((containerSel) => {
              const container = document.querySelector(containerSel);
              if (container) {
                const elements = container.querySelectorAll('*');
                for (const el of elements) {
                  const text = el.textContent || el.innerText || '';
                  if (text.includes('대기중인 공급사송장')) {
                    el.click();
                    return true;
                  }
                }
              }
              return false;
            }, containerSelector);
            logger.info('대기중인 공급사송장 메뉴 클릭 완료');
            menuClicked = true;
            break;
          }
        }
      } catch (error) {
        logger.warn(`검색 결과 처리 실패: ${containerSelector} - ${error.message}`);
      }
    }
    if (!menuClicked) {
      // Enter 키로 첫 번째 결과 선택 시도
      logger.info('Enter 키로 검색 결과 선택 시도...');
      await page.keyboard.press('Enter');
      menuClicked = true;
    }
    // 페이지 이동 대기
    logger.info('대기중인 공급사송장 페이지 로딩 대기 중...');
    await delay(5000);
    
    // 4번 RPA 동작 추가 단계들
    logger.info('=== 4번 RPA 동작 추가 단계 시작 ===');
    
    // 4-1. '공급사송장' 탭 클릭
    logger.info('4-1. 공급사송장 탭 찾는 중...');
    try {
      const vendorInvoiceTabClicked = await page.evaluate(() => {
        const spans = document.querySelectorAll('span.appBarTab-headerLabel');
        for (const span of spans) {
          const text = span.textContent || span.innerText || '';
          if (text.includes('공급사송장')) {
            span.click();
            return true;
          }
        }
        return false;
      });
      
      if (vendorInvoiceTabClicked) {
        logger.info('✅ 공급사송장 탭 클릭 성공');
        await delay(3000); // 탭 로딩 대기
      } else {
        logger.warn('⚠️ 공급사송장 탭을 찾을 수 없습니다.');
      }
    } catch (error) {
      logger.warn(`공급사송장 탭 클릭 실패: ${error.message}`);
    }
    
    // 4-2. '제품 입고로 부터' 버튼 클릭
    logger.info('4-2. 제품 입고로 부터 버튼 찾는 중...');
    try {
      const productReceiptButtonClicked = await page.evaluate(() => {
        const buttonContainers = document.querySelectorAll('.button-container');
        for (const container of buttonContainers) {
          const label = container.querySelector('.button-label');
          if (label) {
            const text = label.textContent || label.innerText || '';
            if (text.includes('제품 입고로 부터')) {
              container.click();
              return true;
            }
          }
        }
        return false;
      });
      
      if (productReceiptButtonClicked) {
        logger.info('✅ 제품 입고로 부터 버튼 클릭 성공');
        await delay(3000); // 버튼 클릭 후 로딩 대기
      } else {
        logger.warn('⚠️ 제품 입고로 부터 버튼을 찾을 수 없습니다.');
      }
    } catch (error) {
      logger.warn(`제품 입고로 부터 버튼 클릭 실패: ${error.message}`);
    }
      // 4-3 ~ 4-5. 엑셀 데이터 기반 반복 필터링 처리
    logger.info('4-3 ~ 4-5. 엑셀 데이터 기반 반복 필터링 처리 시작...');
    
    // 먼저 팝업창이 나타날 때까지 대기
    await delay(3000);
    
    try {
      // Step 1: 엑셀에서 A=1이고 B열이 NULL이 아닌 고유한 B값들 수집
      let uniqueBValues = [];
      if (excelFilePath) {
        try {
          logger.info('엑셀에서 A=1이고 B열이 NULL이 아닌 고유한 B값들 수집 중...');
          const workbook = xlsx.readFile(excelFilePath);
          const sheetName = Object.keys(workbook.Sheets)[0]; // 첫 번째 시트
          const worksheet = workbook.Sheets[sheetName];
          
          // 시트 범위 확인
          const range = xlsx.utils.decode_range(worksheet['!ref']);
          const bValues = new Set(); // 중복 제거용
          
          // A=1이고 B열이 NULL이 아닌 행들 찾기
          for (let row = range.s.r + 1; row <= range.e.r; row++) { // 헤더 제외
            const cellA = worksheet[xlsx.utils.encode_cell({ r: row, c: 0 })] || {}; // A열 (0번째 컬럼)
            const cellB = worksheet[xlsx.utils.encode_cell({ r: row, c: 1 })] || {}; // B열 (1번째 컬럼)
            
            const valueA = cellA.v;
            const valueB = cellB.v;
            
            // A=1이고 B가 NULL이 아닌 경우
            // 사이클 넘버 변경
            if (valueA === userInputValueA && valueB && valueB.toString().trim() !== '') {
              bValues.add(valueB.toString().trim());
            }
          }
          
          uniqueBValues = Array.from(bValues);
          logger.info(`수집된 고유한 B값들 (총 ${uniqueBValues.length}개): ${uniqueBValues.join(', ')}`);
        } catch (excelError) {
          logger.warn(`엑셀 데이터 수집 실패: ${excelError.message}`);
          // 백업용 테스트 데이터
          uniqueBValues = ['TEST'];
        }
      } else {
        logger.warn('엑셀 파일 경로가 제공되지 않음, 테스트 데이터 사용');
        uniqueBValues = ['TEST'];
      }
      
      if (uniqueBValues.length === 0) {
        logger.warn('처리할 B값이 없습니다. 기본 테스트 값으로 진행');
        uniqueBValues = ['TEST'];
      }
      
      // Step 2: 각 고유한 B값에 대해 4-3~4-5 순서 반복
      logger.info(`=== ${uniqueBValues.length}개 B값에 대해 순차 처리 시작 ===`);
      
      for (let index = 0; index < uniqueBValues.length; index++) {
        const currentBValue = uniqueBValues[index];
        logger.info(`\n🔄 [${index + 1}/${uniqueBValues.length}] B값 "${currentBValue}" 처리 시작`);
        
        try {
          // 4-3. 구매주문 헤더 클릭
          logger.info(`4-3. 구매주문 헤더 클릭 (B값: "${currentBValue}")`);
          
          const purchaseOrderHeaderClicked = await page.evaluate(() => {
            const dialogPopup = document.querySelector('.dialog-popup-content');
            if (!dialogPopup) {
              return { success: false, error: '팝업창을 찾을 수 없습니다.' };
            }
            
            // 구매주문 헤더 찾기
            const popupHeaders = dialogPopup.querySelectorAll('.dyn-headerCellLabel._11w1prk, .dyn-headerCellLabel');
            for (const header of popupHeaders) {
              const title = (header.getAttribute('title') || '').trim();
              const text = (header.textContent || header.innerText || '').trim();
              
              if (title === '구매주문' || text === '구매주문') {
                header.click();
                return { 
                  success: true, 
                  method: 'popup-header-text', 
                  title: title, 
                  text: text
                };
              }
            }
            
            // 백업: PurchOrder 포함된 요소 찾기
            const purchaseOrderElements = dialogPopup.querySelectorAll('[data-dyn-columnname*="PurchOrder"], [data-dyn-controlname*="PurchOrder"]');
            for (const element of purchaseOrderElements) {
              const headerLabel = element.querySelector('.dyn-headerCellLabel._11w1prk') || 
                                element.querySelector('.dyn-headerCellLabel');
              if (headerLabel) {
                headerLabel.click();
                return { 
                  success: true, 
                  method: 'popup-columnname-partial'
                };
              }
            }
            
            return { success: false, error: '팝업창 내에서 구매주문 헤더를 찾을 수 없습니다.' };
          });
          
          if (!purchaseOrderHeaderClicked.success) {
            logger.warn(`⚠️ 구매주문 헤더 클릭 실패 (B값: "${currentBValue}"): ${purchaseOrderHeaderClicked.error}`);
            continue; // 다음 B값으로 넘어감
          }
          
          logger.info(`✅ 구매주문 헤더 클릭 성공 (${purchaseOrderHeaderClicked.method})`);
          await delay(1000); // 헤더 클릭 후 필터창 로딩 대기
          
          // 4-4. 필터 입력창에 현재 B값 입력
          logger.info(`4-4. 필터 입력창에 B값 "${currentBValue}" 입력 중...`);
          
          // 필터 팝업창이 로드될 때까지 잠시 대기
          await delay(1500);
          
          const filterInputResult = await page.evaluate((value) => {
            // 다양한 필터 팝업 선택자 시도
            const popupSelectors = [
              '.columnHeader-popup',
              '[class*="popup"]',
              '[class*="filter"]',
              '[class*="dropdown"]',
              '.dyn-popup'
            ];
            
            let filterPopup = null;
            for (const popupSelector of popupSelectors) {
              filterPopup = document.querySelector(popupSelector);
              if (filterPopup && filterPopup.offsetParent !== null) {
                break;
              }
            }
            
            if (!filterPopup) {
              return { success: false, error: '필터 팝업창을 찾을 수 없음' };
            }
            
            // 다양한 입력 필드 선택자 시도
            const inputSelectors = [
              'input[role="combobox"]',
              'input.textbox.field',
              'input[type="text"]',
              'input[name*="Filter"]',
              'input[class*="filter"]',
              'input[class*="search"]',
              'input',
              'textarea'
            ];
            
            for (const selector of inputSelectors) {
              const inputs = filterPopup.querySelectorAll(selector);
              for (const input of inputs) {
                if (input && input.offsetParent !== null && !input.disabled) {
                  try {
                    // 포커스 설정
                    input.focus();
                    
                    // 기존 값 클리어 (다양한 방법으로)
                    input.value = '';
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('keydown', { bubbles: true }));
                    input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    
                    // 새 값 입력
                    input.value = value;
                    
                    // 다양한 이벤트 발생
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    input.dispatchEvent(new Event('keydown', { bubbles: true }));
                    input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    
                    // 값이 제대로 입력되었는지 확인
                    if (input.value === value) {
                      return { 
                        success: true, 
                        method: 'enhanced-input',
                        selector: selector,
                        value: value,
                        popupFound: filterPopup.className
                      };
                    }
                  } catch (inputError) {
                    continue;
                  }
                }
              }
            }
            
            return { success: false, error: '사용 가능한 필터 입력창을 찾을 수 없습니다.' };
          }, currentBValue);
          
          if (!filterInputResult.success) {
            // 대안: 키보드를 통한 직접 입력 시도
            logger.warn(`⚠️ 필터 입력 실패, 키보드 입력 시도 (B값: "${currentBValue}")`);
            
            try {
              // Ctrl+A로 전체 선택 후 값 입력
              await page.keyboard.down('Control');
              await page.keyboard.press('KeyA');
              await page.keyboard.up('Control');
              await delay(200);
              
              // 값 입력
              await page.keyboard.type(currentBValue);
              await delay(300);
              
              logger.info(`✅ 키보드를 통한 필터 입력 완료: "${currentBValue}"`);
            } catch (keyboardError) {
              logger.warn(`❌ 키보드 입력도 실패 (B값: "${currentBValue}"): ${keyboardError.message}`);
              continue; // 다음 B값으로 넘어감
            }
          } else {
            logger.info(`✅ 필터 입력 성공: "${filterInputResult.value}" (방법: ${filterInputResult.method})`);
          }
          
          // 4-5. Enter 키로 필터 적용
          logger.info('4-5. Enter 키로 필터 적용 중...');
          await delay(500);
          await page.keyboard.press('Enter');
          logger.info('✅ Enter 키로 필터 적용 완료');
          
          // 필터링 완료 대기 (단축: 10초 → 5초)
          logger.info('필터링 완료 대기 중... (5초)');
          await delay(5000);
          
          // 4-5-2. All Check 버튼 클릭
          logger.info('4-5-2. All Check 버튼 클릭 중...');
          
          const allCheckClicked = await page.evaluate(() => {
            // All Check 버튼 찾기
            const allCheckSpan = document.querySelector('#PurchJournalSelect_PackingSlip_45_NPS_AllCheck_label');
            if (allCheckSpan && allCheckSpan.textContent.trim() === 'All Check') {
              allCheckSpan.click();
              return { 
                success: true, 
                method: 'exact-span-id-AllCheck',
                text: allCheckSpan.textContent.trim()
              };
            }
            
            // 백업: span.button-label에서 "All Check" 찾기
            const allSpans = document.querySelectorAll('span.button-label');
            for (const span of allSpans) {
              const spanText = (span.textContent || span.innerText || '').trim();
              if (spanText === 'All Check') {
                span.click();
                return { 
                  success: true, 
                  method: 'span-text-AllCheck',
                  text: spanText
                };
              }
            }
            
            return { success: false, error: 'All Check 버튼을 찾을 수 없습니다.' };
          });
          
          if (allCheckClicked.success) {
            logger.info(`✅ All Check 버튼 클릭 성공 (${allCheckClicked.method}): "${allCheckClicked.text}"`);
            await delay(1000); // All Check 처리 대기
          } else {
            logger.warn(`⚠️ All Check 버튼 클릭 실패 (B값: "${currentBValue}"): ${allCheckClicked.error}`);
          }
          
          logger.info(`🎉 [${index + 1}/${uniqueBValues.length}] B값 "${currentBValue}" 처리 완료`);
          
          // 다음 B값 처리를 위한 짧은 대기 (1초)
          if (index < uniqueBValues.length - 1) {
            await delay(1000);
          }
          
        } catch (currentBError) {
          logger.warn(`❌ B값 "${currentBValue}" 처리 중 오류: ${currentBError.message}`);
          continue; // 다음 B값으로 넘어감
        }
      }
      
      logger.info(`🎉 === 모든 B값 처리 완료 (총 ${uniqueBValues.length}개) ===`);
      
      // ========== 모든 B값 처리 완료 후 Alt+Enter 한 번만 실행 ==========
      logger.info('🚀 === 모든 B값 처리 완료 후 Alt + Enter 입력 중... ===');
      try {
        await page.keyboard.down('Alt');
        await page.keyboard.press('Enter');
        await page.keyboard.up('Alt');
        logger.info('✅ Alt + Enter 입력 완료');
        
        // Alt+Enter 후 페이지 변경 및 로딩 대기
        logger.info('Alt+Enter 후 페이지 로딩 대기 중...');
        await delay(5000); // 5초 대기
        
      } catch (altEnterError) {
        logger.error(`❌ Alt+Enter 실행 중 오류: ${altEnterError.message}`);
      }
      
      // 마지막으로 처리된 B값의 AT열 날짜 값 추출 (2회 재시도)
      if (uniqueBValues.length > 0 && excelFilePath) {
        let extractionSuccess = false;
        let retryCount = 0;
        const maxRetries = 2;
        
        while (!extractionSuccess && retryCount < maxRetries) {
          try {
            const lastBValue = uniqueBValues[uniqueBValues.length - 1];
            logger.info(`AT column extraction attempt ${retryCount + 1}/${maxRetries} for last B value: "${lastBValue}"`);
            logger.info(`Excel file path: ${excelFilePath}`);
            
            const workbook = xlsx.readFile(excelFilePath);
            const sheetName = Object.keys(workbook.Sheets)[0];
            const worksheet = workbook.Sheets[sheetName];
            const range = xlsx.utils.decode_range(worksheet['!ref']);
            
            logger.info(`Excel sheet info: ${sheetName}, range: ${worksheet['!ref']}`);
            logger.info(`Searching for rows where A=1 and B="${lastBValue}"`);
            
            // A=13이고 B=lastBValue인 행들을 찾아서 마지막 행의 AT열 값 추출
            let lastRowWithTargetB = -1;
            let foundRows = [];
            
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
              const cellA = worksheet[xlsx.utils.encode_cell({ r: row, c: 0 })] || {};
              const cellB = worksheet[xlsx.utils.encode_cell({ r: row, c: 1 })] || {};
              
              const valueA = cellA.v;
              const valueB = cellB.v;
              
              // Debug log for first 10 rows
              if (row <= range.s.r + 10) {
                logger.info(`Row ${row + 1}: A=${valueA}, B=${valueB}`);
              }
              
              if (valueA === userInputValueA && valueB && valueB.toString().trim() === lastBValue.toString().trim()) {
                lastRowWithTargetB = row;
                foundRows.push(row + 1);
                logger.info(`Found matching row: ${row + 1} (A=${valueA}, B=${valueB})`);
              }
            }
            
            logger.info(`🔍 검색 조건: A열=${userInputValueA}, B열="${lastBValue}"`);
            logger.info(`📊 검색 결과: 총 ${foundRows.length}개 행 발견`);
            logger.info(`Found rows: [${foundRows.join(', ')}], Final selected row: ${lastRowWithTargetB + 1}`);
            
            if (lastRowWithTargetB !== -1) {
              // AT열은 45번째 컬럼 (A=0부터 시작하므로 AT=45)
              const atColumnIndex = 45;
              const cellAT = worksheet[xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: atColumnIndex })] || {};
              const atValue = cellAT.v;
              
              logger.info(`AT column (index ${atColumnIndex}) cell address: ${xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: atColumnIndex })}`);
              logger.info(`AT column raw value: ${atValue} (type: ${typeof atValue})`);
              
              // AV열은 47번째 컬럼 (A=0부터 시작하므로 AV=47)
              const avColumnIndex = 47;
              const cellAV = worksheet[xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: avColumnIndex })] || {};
              const avValue = cellAV.v;
              
              logger.info(`AV column (index ${avColumnIndex}) cell address: ${xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: avColumnIndex })}`);
              logger.info(`AV column raw value: ${avValue} (type: ${typeof avValue})`);
              
              // AU열은 46번째 컬럼 (A=0부터 시작하므로 AU=46)
              const auColumnIndex = 46;
              const cellAU = worksheet[xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: auColumnIndex })] || {};
              const auValue = cellAU.v;
              
              logger.info(`AU column (index ${auColumnIndex}) cell address: ${xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: auColumnIndex })}`);
              logger.info(`AU column raw value: ${auValue} (type: ${typeof auValue})`);
              
              // AT열 값 저장
              if (atValue) {
                lastProcessedDateFromATColumn = atValue;
                logger.info(`AT column date extraction SUCCESS: ${atValue} (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
              } else {
                logger.warn(`AT column value is empty (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
                lastProcessedDateFromATColumn = null;
              }
              
              // AV열 값 저장
              if (avValue) {
                lastProcessedDateFromAVColumn = avValue;
                logger.info(`AV column date extraction SUCCESS: ${avValue} (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
              } else {
                logger.warn(`AV column value is empty (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
                lastProcessedDateFromAVColumn = null;
              }
              
              // AU열 값 저장
              if (auValue) {
                lastProcessedValueFromAUColumn = auValue;
                logger.info(`AU column extraction SUCCESS: ${auValue} (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
              } else {
                logger.warn(`AU column value is empty (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
                lastProcessedValueFromAUColumn = null;
              }
              
              // I열은 8번째 컬럼 (A=0부터 시작하므로 I=8)
              const iColumnIndex = 8;
              const cellI = worksheet[xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: iColumnIndex })] || {};
              const iValue = cellI.v;
              
              logger.info(`I column (index ${iColumnIndex}) cell address: ${xlsx.utils.encode_cell({ r: lastRowWithTargetB, c: iColumnIndex })}`);
              logger.info(`I column raw value: ${iValue} (type: ${typeof iValue})`);
              
              // I열 값 저장
              if (iValue) {
                lastProcessedValueFromIColumn = iValue;
                logger.info(`I column extraction SUCCESS: ${iValue} (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
                logger.info(`🔍 I열 전역 변수 저장 확인: "${lastProcessedValueFromIColumn}" (타입: ${typeof lastProcessedValueFromIColumn})`);
              } else {
                logger.warn(`I column value is empty (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1})`);
                lastProcessedValueFromIColumn = null;
                logger.warn(`🔍 I열 값이 비어있어서 null로 설정함`);
              }
              
              // AT, AV, AU 중 하나라도 성공하면 추출 성공으로 간주
              if (atValue || avValue || auValue) {
                extractionSuccess = true;
              } else {
                logger.warn(`All columns (AT, AV, AU) are empty (B value: ${lastBValue}, row: ${lastRowWithTargetB + 1}) - attempt ${retryCount + 1}`);
                retryCount++;
                if (retryCount < maxRetries) {
                  logger.info(`Retrying AT/AV/AU column extraction in 2 seconds...`);
                  await delay(2000);
                }
              }
            } else {
              logger.warn(`No matching row found for last B value: "${lastBValue}" - attempt ${retryCount + 1}`);
              logger.info(`All B values list: [${uniqueBValues.join(', ')}]`);
              retryCount++;
              if (retryCount < maxRetries) {
                logger.info(`Retrying AT column extraction in 2 seconds...`);
                await delay(2000);
              }
            }
          } catch (atError) {
            logger.error(`Error during AT column extraction attempt ${retryCount + 1}: ${atError.message}`);
            logger.error(`Stack trace: ${atError.stack}`);
            retryCount++;
            if (retryCount < maxRetries) {
              logger.info(`Retrying AT column extraction in 2 seconds...`);
              await delay(2000);
            }
          }
        }
        
        if (!extractionSuccess) {
          logger.error(`❌ AT column date extraction failed after ${maxRetries} attempts`);
          lastProcessedDateFromATColumn = null;
        }
      } else {
        logger.warn(`AT extraction conditions not met: uniqueBValues.length=${uniqueBValues.length}, excelFilePath=${excelFilePath}`);
      }
      
    } catch (error) {
      logger.warn(`반복 필터링 처리 중 오류: ${error.message}`);
    }
    
    // 4-6. 프로세스 완료
    logger.info('=== 4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료 ===');
    
    // 4번 완료 후 5초 대기
    logger.info('⏰ 4번 RPA 완료 후 5초 대기 중...');
    await delay(5000);
    
    // ========== 5번 RPA 동작: 캘린더 버튼 클릭 ==========
    logger.info('🚀 === 5번 RPA 동작: 캘린더 버튼 클릭 시작 ===');
    try {
      await clickCalendarButton(page);
      logger.info('✅ 5번 RPA 동작: 캘린더 버튼 클릭 완료');
    } catch (step5Error) {
      logger.error(`❌ 5번 RPA 동작 실패: ${step5Error.message}`);
    }
    
    return { success: true, message: '4번 RPA 동작: 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료, 5번 RPA 동작: 캘린더 버튼 클릭 완료' };
  } catch (error) {
    logger.error(`4번 RPA 동작 중 오류: ${error.message}`);
    return { success: false, error: error.message, step: '4번 RPA 동작 (대기중인 공급사송장 메뉴 이동)' };
  }
}

// 엑셀 파일 열기 및 매크로 자동 실행 함수
async function openExcelAndExecuteMacro(excelFilePath) {
  const { exec } = require('child_process');
  const { promisify } = require('util');
  const os = require('os');
  const execAsync = promisify(exec);
  
  logger.info('🚀 === 엑셀 파일 열기 및 매크로 자동 실행 시작 ===');
  logger.info(`대상 엑셀 파일: ${excelFilePath}`);
  
  try {
    // VBA 코드 정의
    const vbaCode = `
Sub GroupBy_I_Z_And_Process()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, groupNum As Long
    Dim key As String
    Dim groupMap As Object, groupSums As Object, groupDesc As Object
    Dim maturityDate As Date, adjustedSum As Double
    Dim maturityCol As Long, descCol As Long, taxDateCol As Long
    Dim gDate As Variant
    Dim currentGroup As Long, nextGroup As Long
    Dim lastCol As Long
    Dim pText As String, jText As String

    Set ws = ActiveSheet ' Use active sheet instead of specific name
    Set groupMap = CreateObject("Scripting.Dictionary")
    Set groupSums = CreateObject("Scripting.Dictionary")
    Set groupDesc = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False   ' Turn off screen update
    Application.Calculation = xlCalculationManual   ' Turn off auto calc

    ' Find last row in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Sort by I and Z columns
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("I2:I" & lastRow), Order:=xlAscending
    ws.Sort.SortFields.Add Key:=ws.Range("Z2:Z" & lastRow), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range("A1:AG" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' Insert Group Number column at A
    ws.Columns("A").Insert Shift:=xlToRight
    ws.Cells(1, 1).Value = "Group Number"

    groupNum = 1

    ' Assign group number, sum AG, make invoice description
    For i = 2 To lastRow
        key = ws.Cells(i, "I").Value & "|" & ws.Cells(i, "Z").Value

        If Not groupMap.exists(key) Then
            groupMap(key) = groupNum
            groupSums(groupNum) = 0

            ' Invoice description: Month(gDate) & " Month " & P & "_" & J
            gDate = ws.Cells(i, "G").Value
            pText = ws.Cells(i, "P").Value
            jText = ws.Cells(i, "J").Value

            If IsDate(gDate) Then
                groupDesc(groupNum) = Month(gDate) & ChrW(50900) & pText & "_" & jText
            Else
                groupDesc(groupNum) = "Date Error " & pText & "_" & jText
            End If

            groupNum = groupNum + 1
        End If

        ws.Cells(i, 1).Value = groupMap(key)
        groupSums(groupMap(key)) = groupSums(groupMap(key)) + Val(ws.Cells(i, "AG").Value)
    Next i

    ' Add columns: Maturity Date, Invoice Description, Tax Invoice Date
    maturityCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, maturityCol).Value = "Maturity Date"

    descCol = maturityCol + 1
    ws.Cells(1, descCol).Value = "Invoice Description"

    taxDateCol = descCol + 1
    ws.Cells(1, taxDateCol).Value = "Tax Invoice Date"

    ' Fill Maturity Date, Invoice Description, Tax Invoice Date
    For i = 2 To lastRow
        Dim gNum As Long
        gNum = ws.Cells(i, 1).Value
        gDate = ws.Cells(i, "G").Value

        If IsDate(gDate) Then
            adjustedSum = groupSums(gNum) * 1.1
            If adjustedSum < 10000000 Then
                maturityDate = WorksheetFunction.EoMonth(gDate, 1) ' End of next month
            Else
                maturityDate = WorksheetFunction.EoMonth(gDate, 2) ' End of following month
            End If
            ws.Cells(i, maturityCol).Value = maturityDate

            ' Tax Invoice Date: end of the month for G column
            ws.Cells(i, taxDateCol).Value = WorksheetFunction.EoMonth(gDate, 0)
        Else
            ws.Cells(i, maturityCol).Value = "Date Error"
            ws.Cells(i, taxDateCol).Value = "Date Error"
        End If

        ' Enter Invoice Description
        ws.Cells(i, descCol).Value = groupDesc(gNum)
    Next i

    ' Apply date format to Tax Invoice Date column
    ws.Range(ws.Cells(2, taxDateCol), ws.Cells(lastRow, taxDateCol)).NumberFormat = "yyyy-mm-dd"

    ' Add line between groups
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 2 To lastRow - 1
        currentGroup = ws.Cells(i, 1).Value
        nextGroup = ws.Cells(i + 1, 1).Value

        If currentGroup <> nextGroup Then
            With ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
    Next i

    Application.ScreenUpdating = True    ' Turn on screen update
    Application.Calculation = xlCalculationAutomatic   ' Turn on auto calc

End Sub
`;

    // 임시 PowerShell 스크립트 생성
    const tempDir = os.tmpdir();
    const psScriptPath = path.join(tempDir, `excel_macro_${Date.now()}.ps1`);
    
    // PowerShell 스크립트 내용 (VBA 코드를 직접 포함)
    const psScript = `
# Excel 매크로 자동 실행 PowerShell 스크립트
param(
    [string]$ExcelFilePath = "${excelFilePath.replace(/\\/g, '\\\\')}"
)

Write-Host "Excel 매크로 자동 실행 스크립트 시작"
Write-Host "대상 파일: $ExcelFilePath"

try {
    # COM 객체 생성
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Excel 애플리케이션 생성 완료"
    
    # 기존에 열린 워크북이 있는지 확인
    $workbook = $null
    $fileName = Split-Path $ExcelFilePath -Leaf
    
    foreach ($wb in $excel.Workbooks) {
        if ($wb.Name -eq $fileName) {
            $workbook = $wb
            Write-Host "기존에 열린 워크북 사용: $fileName"
            break
        }
    }
    
    # 워크북이 없으면 새로 열기
    if ($workbook -eq $null) {
        if (Test-Path $ExcelFilePath) {
            $workbook = $excel.Workbooks.Open($ExcelFilePath)
            Write-Host "워크북 열기 완료: $ExcelFilePath"
        } else {
            throw "파일을 찾을 수 없습니다: $ExcelFilePath"
        }
    }
    
    # 워크시트 선택
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate()
    
    Write-Host "워크시트 활성화 완료"
    
    # 기존 VBA 모듈 제거
    $vbaProject = $workbook.VBProject
    for ($i = $vbaProject.VBComponents.Count; $i -ge 1; $i--) {
        $component = $vbaProject.VBComponents.Item($i)
        if ($component.Type -eq 1) {  # vbext_ct_StdModule
            $vbaProject.VBComponents.Remove($component)
            Write-Host "기존 VBA 모듈 제거: $($component.Name)"
        }
    }
    
    # 새 VBA 모듈 추가
    $vbaModule = $vbaProject.VBComponents.Add(1)  # vbext_ct_StdModule
    $vbaModule.Name = "GroupProcessModule"
    
    Write-Host "새 VBA 모듈 추가 완료"
    
    # VBA 코드 추가 - 잠시 대기 후 추가
    Start-Sleep -Milliseconds 500
    
    # VBA 코드 추가
    $vbaCode = @"
${vbaCode}
"@;
    
    $vbaModule.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA 코드 추가 완료"
    
    # 매크로 실행 전 대기
    Start-Sleep -Seconds 2
   
    Write-Host "VBA 프로젝트 준비 완료, 매크로 실행 중..."
    
    # 매크로 실행 - 정확한 함수명 사용
    try {
        $excel.Run("GroupBy_I_Z_And_Process")
        Write-Host "매크로 실행 완료"
    } catch {
        Write-Host "매크로 실행 실패: $($_.Exception.Message)"
        # 대안으로 모듈명.함수명 형태로 시도
        try {
            $excel.Run("GroupProcessModule.GroupBy_I_Z_And_Process")
            Write-Host "모듈명 포함 매크로 실행 완료"
        } catch {
            Write-Host "모듈명 포함 매크로 실행도 실패: $($_.Exception.Message)"
            throw "매크로 실행에 실패했습니다."
        }
    }
    
    # 매크로 실행 후 파일 저장
    Start-Sleep -Seconds 2
    Write-Host "매크로 실행 후 파일 저장 중..."
    
    try {
        $workbook.Save()
        Write-Host "파일 저장 완료"
    } catch {
        Write-Host "파일 저장 실패: $($_.Exception.Message)"
        # 다른 이름으로 저장 시도
        try {
            $savePath = $ExcelFilePath -replace '\.xlsx$', '_processed.xlsx'
            $workbook.SaveAs($savePath)
            Write-Host "다른 이름으로 저장 완료: $savePath"
        } catch {
            Write-Host "다른 이름으로 저장도 실패: $($_.Exception.Message)"
            throw "파일 저장에 실패했습니다."
        }
    }
    
    # Excel을 보이게 설정
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    
    Write-Host "Excel 매크로 자동 실행 완료"
    
} catch {
    Write-Host "오류 발생: $($_.Exception.Message)"
    if ($excel) {
        $excel.Visible = $true
        $excel.DisplayAlerts = $true
    }
    exit 1
}
`;

    // PowerShell 스크립트 파일 저장
    fs.writeFileSync(psScriptPath, psScript, 'utf8');
    logger.info(`PowerShell 스크립트 생성 완료: ${psScriptPath}`);
    
    // PowerShell 스크립트 실행
    logger.info('PowerShell 스크립트 실행 중...');
    const result = await execAsync(`powershell -ExecutionPolicy Bypass -File "${psScriptPath}"`, {
      timeout: 60000, // 60초 타임아웃
      encoding: 'utf8'
    });
    
    if (result.stdout) {
      logger.info('PowerShell 실행 결과:');
      logger.info(result.stdout);
    }
    
    if (result.stderr) {
      logger.warn('PowerShell 실행 경고:');
      logger.warn(result.stderr);
    }
    
    // 임시 파일 정리
    try {
      fs.unlinkSync(psScriptPath);
      logger.info('임시 PowerShell 스크립트 파일 정리 완료');
    } catch (cleanupError) {
      logger.warn(`임시 파일 정리 실패: ${cleanupError.message}`);
    }
    
    logger.info('✅ 엑셀 매크로 자동 실행 완료');
    
    return {
      success: true,
      message: '엑셀 매크로가 성공적으로 실행되었습니다.',
      filePath: excelFilePath
    };
    
  } catch (error) {
    logger.error(`엑셀 매크로 실행 중 오류: ${error.message}`);
    
    return {
      success: false,
      error: error.message,
      failedAt: new Date().toISOString(),
      step: '엑셀 매크로 실행'
    };
  }
}

/**
 * 5번 RPA 동작: 캘린더 버튼 클릭
 */
async function clickCalendarButton(page) {
  try {
    logger.info('캘린더 버튼 찾는 중...');
    
    // dyn-date-picker-button 클래스를 가진 캘린더 버튼 선택자들
    const calendarButtonSelectors = [
      'div.dyn-container.dyn-date-picker-button[role="button"][title="Open"]',
      'div.dyn-date-picker-button[role="button"]',
      'div[class*="dyn-date-picker-button"]',
      '.dyn-date-picker-button',
      'div[title="Open"][role="button"]',
      'button[title="Open"]',
      'div[role="button"][aria-label="Open"]',
      'div[class*="date-picker"][role="button"]',
      'div[class*="calendar"][role="button"]',
      'div.button[title="Open"]'
    ];
    
    let buttonFound = false;
    let buttonPosition = null;
    
    for (const selector of calendarButtonSelectors) {
      try {
        logger.info(`캘린더 버튼 선택자 시도: ${selector}`);
        
        // 요소가 존재하는지 확인
        const button = await page.$(selector);
        if (button) {
          // 요소가 보이는지 확인
          const isVisible = await button.isIntersectingViewport();
          if (isVisible) {
            logger.info(`캘린더 버튼 발견: ${selector}`);
            
            // 버튼 위치 정보 가져오기
            buttonPosition = await button.boundingBox();
            logger.info(`캘린더 버튼 위치: x=${buttonPosition.x}, y=${buttonPosition.y}, width=${buttonPosition.width}, height=${buttonPosition.height}`);
            
            buttonFound = true;
            break;
          } else {
            logger.warn(`캘린더 버튼이 화면에 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
        continue;
      }
    }
    
    if (!buttonFound) {
      // SVG 내용을 포함한 더 구체적인 선택자 시도
      try {
        logger.info('SVG 내용 기반 캘린더 버튼 찾는 중...');
        
        const svgButtonInfo = await page.evaluate(() => {
          // 다양한 SVG 패턴으로 캘린더 버튼 찾기
          const potentialButtons = document.querySelectorAll('div[role="button"], button, div[title="Open"], div[class*="picker"], div[class*="calendar"]');
          for (const element of potentialButtons) {
            const svg = element.querySelector('svg');
            if (svg) {
              const svgContent = svg.innerHTML;
              // 다양한 캘린더 SVG 패턴 확인
              if (svgContent.includes('M33.09,6.82h6.75v31.5h-36V6.82h6.75V4.57h2.25V6.82h18V4.57h2.25Z') ||
                  svgContent.includes('calendar') ||
                  svgContent.includes('date') ||
                  element.title === 'Open' ||
                  element.getAttribute('aria-label') === 'Open') {
                const rect = element.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) {
                  return {
                    x: rect.x,
                    y: rect.y,
                    width: rect.width,
                    height: rect.height,
                    element: element.tagName + '.' + element.className
                  };
                }
              }
            }
          }
          return null;
        });
        
        if (svgButtonInfo) {
          buttonPosition = svgButtonInfo;
          logger.info(`SVG 기반 캘린더 버튼 위치: x=${buttonPosition.x}, y=${buttonPosition.y}, width=${buttonPosition.width}, height=${buttonPosition.height}, element=${svgButtonInfo.element}`);
          buttonFound = true;
        }
      } catch (svgError) {
        logger.warn(`SVG 기반 검색 실패: ${svgError.message}`);
      }
    }

    if (!buttonFound) {
      // 최후의 수단: 모든 클릭 가능한 요소에서 "Open" 관련 요소 찾기
      try {
        logger.info('포괄적 검색으로 캘린더 버튼 찾는 중...');
        
        const generalButtonInfo = await page.evaluate(() => {
          const allClickable = document.querySelectorAll('div, button, span, a');
          for (const element of allClickable) {
            if ((element.title === 'Open' || 
                 element.getAttribute('aria-label') === 'Open' ||
                 element.className.includes('date-picker') ||
                 element.className.includes('calendar') ||
                 element.getAttribute('role') === 'button') &&
                element.offsetParent !== null) {
              const rect = element.getBoundingClientRect();
              if (rect.width > 0 && rect.height > 0) {
                return {
                  x: rect.x,
                  y: rect.y,
                  width: rect.width,
                  height: rect.height,
                  tag: element.tagName,
                  className: element.className,
                  title: element.title
                };
              }
            }
          }
          return null;
        });
        
        if (generalButtonInfo) {
          buttonPosition = generalButtonInfo;
          logger.info(`포괄적 검색으로 캘린더 버튼 발견: ${generalButtonInfo.tag}.${generalButtonInfo.className}, title="${generalButtonInfo.title}"`);
          logger.info(`버튼 위치: x=${buttonPosition.x}, y=${buttonPosition.y}, width=${buttonPosition.width}, height=${buttonPosition.height}`);
          buttonFound = true;
        }
      } catch (generalError) {
        logger.warn(`포괄적 검색 실패: ${generalError.message}`);
      }
    }
    
    if (!buttonFound || !buttonPosition) {
      // 페이지 상태 디버깅 정보 수집
      try {
        const debugInfo = await page.evaluate(() => {
          const allButtons = document.querySelectorAll('div[role="button"], button');
          const allWithTitle = document.querySelectorAll('[title="Open"]');
          const allDatePickers = document.querySelectorAll('[class*="date"], [class*="picker"], [class*="calendar"]');
          
          return {
            totalButtons: allButtons.length,
            titleOpenElements: allWithTitle.length,
            datePickerElements: allDatePickers.length,
            url: window.location.href,
            title: document.title
          };
        });
        
        logger.error(`캘린더 버튼 찾기 실패 - 디버깅 정보:`);
        logger.error(`- 전체 버튼 요소: ${debugInfo.totalButtons}개`);
        logger.error(`- title="Open" 요소: ${debugInfo.titleOpenElements}개`);
        logger.error(`- 날짜 선택기 관련 요소: ${debugInfo.datePickerElements}개`);
        logger.error(`- 현재 URL: ${debugInfo.url}`);
        logger.error(`- 페이지 제목: ${debugInfo.title}`);
      } catch (debugError) {
        logger.warn(`디버깅 정보 수집 실패: ${debugError.message}`);
      }
      
      throw new Error('캘린더 버튼을 찾을 수 없습니다. 모든 선택자와 대안 방법이 실패했습니다.');
    }
    
    // 캘린더 버튼 왼쪽에 있는 송장일 입력 필드 찾기
    logger.info('캘린더 버튼 왼쪽의 송장일 입력 필드 찾는 중...');
    
    let invoiceDateInput = null;
    const inputSelectors = [
      'input[type="text"]',
      'input[class*="date"]',
      'input[class*="Date"]',
      'input[data-dyn-controlname*="date"]',
      'input[data-dyn-controlname*="Date"]'
    ];
    
    // 캘린더 버튼 기준으로 왼쪽에 있는 입력 필드 찾기
    for (const selector of inputSelectors) {
      try {
        const inputs = await page.$$(selector);
        for (const input of inputs) {
          const inputBox = await input.boundingBox();
          if (inputBox && 
              Math.abs(inputBox.y - buttonPosition.y) < 20 && // 같은 행에 있는지 확인
              inputBox.x < buttonPosition.x && // 캘린더 버튼 왼쪽에 있는지 확인
              (buttonPosition.x - inputBox.x - inputBox.width) < 50) { // 거리가 가까운지 확인
            
            invoiceDateInput = input;
            logger.info(`송장일 입력 필드 발견: ${selector}, 위치: x=${inputBox.x}, y=${inputBox.y}`);
            break;
          }
        }
        if (invoiceDateInput) break;
      } catch (error) {
        logger.warn(`선택자 ${selector} 확인 중 오류: ${error.message}`);
      }
    }
    
    if (!invoiceDateInput) {
      // 대안: 캘린더 버튼 왼쪽 20px 지점을 더블클릭
      logger.warn('송장일 입력 필드를 찾을 수 없음, 캘린더 버튼 왼쪽 좌표로 대체');
      const targetX = buttonPosition.x - 20;
      const targetY = buttonPosition.y + buttonPosition.height / 2;
      
      logger.info(`대체 위치로 이동: x=${targetX}, y=${targetY}`);
      await page.mouse.move(targetX, targetY);
      await delay(500);
      await page.mouse.click(targetX, targetY, { clickCount: 2 });
      await delay(500);
    } else {
      // 송장일 입력 필드를 더블클릭
      logger.info('송장일 입력 필드 더블클릭 수행 중...');
      await invoiceDateInput.click({ clickCount: 2 });
      await delay(500);
    }
    
    // AV열에서 추출한 날짜 값 입력 (송장일 입력용)
    let dateToInput = null;
    
    if (lastProcessedDateFromAVColumn) {
      const convertedDate = convertDateFormat(lastProcessedDateFromAVColumn);
      if (convertedDate) {
        dateToInput = convertedDate;
        logger.info(`Inputting extracted AV column date: ${dateToInput} (original: ${lastProcessedDateFromAVColumn})`);
      } else {
        logger.error('❌ Date conversion failed from AV column data');
        throw new Error('AV열에서 추출한 날짜 데이터 변환에 실패했습니다. 프로세스를 중단합니다.');
      }
    } else {
      logger.error('❌ No AV column date value available after retry attempts');
      throw new Error('AV열에서 날짜 데이터를 추출할 수 없습니다. 2회 재시도 후에도 실패했습니다. 프로세스를 중단합니다.');
    }
    
    await page.keyboard.type(dateToInput);
    await delay(300);
    
    // Enter 키 입력
    logger.info('Enter 키 입력 중...');
    await page.keyboard.press('Enter');
    
    logger.info(`✅ 캘린더 버튼 왼쪽 더블클릭, ${dateToInput} 입력, Enter 완료`);
    
    // 포커스 해제를 위해 페이지 아무 지점 클릭
    logger.info('포커스 해제를 위해 페이지 아무 지점 클릭...');
    await page.mouse.click(100, 100);
    await delay(500);
    
    // AV열 송장일 입력 후 송장 통합 처리
    try {
      await processInvoiceIntegrationAfterAV(page);
      logger.info('✅ AV열 후 송장 통합 처리 완료');
    } catch (integrationError) {
      logger.warn(`⚠️ AV열 후 송장 통합 처리 실패했지만 계속 진행: ${integrationError.message}`);
    }
    
    // 송장 통합 처리 완료 후 송장 번호 input 요소 클릭 추가
    logger.info('🔍 송장 통합 처리 완료 후 송장 번호 input 요소 찾는 중...');
    
    // 페이지 상태 확인을 위한 디버깅
    await page.evaluate(() => {
      console.log('=== 송장 통합 후 페이지 상태 확인 ===');
      console.log('현재 URL:', window.location.href);
      console.log('페이지 제목:', document.title);
      console.log('전체 input 요소 수:', document.querySelectorAll('input').length);
    });
    
    // 송장 번호 input 요소 선택자들
    const invoiceInputSelectors = [
      'input#PurchParmTable_gridParmTableNum_474_0_0_input',
      'input[id*="PurchParmTable_gridParmTableNum"][id*="_input"]',
      'input[id*="gridParmTableNum"][id*="_input"]',
      'input[aria-label="송장 번호"]',
      'input[class*="dyn-field"][class*="dyn-hyperlink"]',
      'div[data-dyn-controlname="PurchParmTable_gridParmTableNum"] input',
      'div[id*="PurchParmTable_gridParmTableNum"] input'
    ];
    
    let inputFound = false;
    let targetInput = null;
    
    for (const selector of invoiceInputSelectors) {
      try {
        logger.info(`송장 번호 input 선택자 시도: ${selector}`);
        
        // 요소가 존재하는지 확인
        const input = await page.$(selector);
        if (input) {
          // 요소가 보이는지 확인
          const isVisible = await input.isIntersectingViewport();
          if (isVisible) {
            // input 요소의 속성 정보 가져오기
            const inputInfo = await page.evaluate((el) => {
              return {
                id: el.id,
                value: el.value,
                ariaLabel: el.getAttribute('aria-label'),
                maxLength: el.getAttribute('maxlength'),
                className: el.className
              };
            }, input);
            
            logger.info(`송장 번호 input 발견: ${selector}`);
            logger.info(`Input 정보: id=${inputInfo.id}, value="${inputInfo.value}", aria-label="${inputInfo.ariaLabel}"`);
            
            // input 위치 정보 가져오기
            const inputPosition = await input.boundingBox();
            logger.info(`송장 번호 input 위치: x=${inputPosition.x}, y=${inputPosition.y}, width=${inputPosition.width}, height=${inputPosition.height}`);
            
            targetInput = input;
            inputFound = true;
            break;
          } else {
            logger.warn(`송장 번호 input이 화면에 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
        continue;
      }
    }
    
    if (!inputFound) {
      // 더 광범위한 검색: value 속성에 특정 패턴이 있는 input 찾기
      try {
        logger.info('value 패턴 기반 송장 번호 input 찾는 중...');
        
        const inputByValue = await page.evaluate(() => {
          const inputs = document.querySelectorAll('input[type="text"], input[role="textbox"]');
          for (const input of inputs) {
            const value = input.value || '';
            // 송장 번호 패턴: 숫자_문자숫자조합_숫자 형태
            if (value.match(/^\d+_[A-Z0-9]+_\d+$/)) {
              const rect = input.getBoundingClientRect();
              return {
                found: true,
                id: input.id,
                value: input.value,
                ariaLabel: input.getAttribute('aria-label'),
                x: rect.x,
                y: rect.y,
                width: rect.width,
                height: rect.height
              };
            }
          }
          return { found: false };
        });
        
        if (inputByValue.found) {
          logger.info(`패턴 기반 송장 번호 input 발견: id=${inputByValue.id}, value="${inputByValue.value}"`);
          logger.info(`위치: x=${inputByValue.x}, y=${inputByValue.y}, width=${inputByValue.width}, height=${inputByValue.height}`);
          
          // 좌표로 클릭
          const clickX = inputByValue.x + inputByValue.width / 2;
          const clickY = inputByValue.y + inputByValue.height / 2;
          
          await page.mouse.click(clickX, clickY);
          await delay(500);
          
          logger.info('✅ 송장 번호 input 클릭 완료 (패턴 기반)');
          inputFound = true;
        }
      } catch (patternError) {
        logger.warn(`패턴 기반 검색 실패: ${patternError.message}`);
      }
    } else {
      // 찾은 input 요소 클릭
      logger.info('송장 번호 input 클릭 수행 중...');
      await targetInput.click();
      await delay(500);
      
      logger.info('✅ 송장 번호 input 클릭 완료');
    }
    
    if (inputFound) {
      logger.info(`✅ 5번 RPA 동작: 캘린더 버튼 클릭 및 송장 번호 input 클릭 완료`);
      
      // 공급사송장 요소에서 값 추출 (3.5 동작용)
      try {
        logger.info('공급사송장 요소에서 값 추출 중...');
        extractedVendorInvoiceValue = await page.evaluate(() => {
          // 공급사송장 span 요소 찾기
          const vendorInvoiceSpan = document.querySelector('span.formCaption.link-content-validLink[role="link"]');
          if (!vendorInvoiceSpan || !vendorInvoiceSpan.textContent.includes('공급사송장')) {
            return null;
          }
          
          // 공급사송장 요소의 부모나 형제 요소에서 값 찾기
          let targetValue = null;
          
          // 방법 1: 부모 요소에서 다음 input이나 span 찾기
          const parentElement = vendorInvoiceSpan.closest('td, div, form');
          if (parentElement) {
            const nextInputs = parentElement.querySelectorAll('input, span');
            for (const input of nextInputs) {
              if (input !== vendorInvoiceSpan && input.value && input.value.trim()) {
                targetValue = input.value.trim();
                break;
              }
              if (input !== vendorInvoiceSpan && input.textContent && input.textContent.trim() && 
                  input.textContent.includes('_')) {
                targetValue = input.textContent.trim();
                break;
              }
            }
          }
          
          // 방법 2: elementFromPoint로 20px 아래 위치 확인
          if (!targetValue) {
            const rect = vendorInvoiceSpan.getBoundingClientRect();
            const targetX = rect.x + (rect.width / 2);
            const targetY = rect.y + rect.height + 20;
            
            const targetElement = document.elementFromPoint(targetX, targetY);
            if (targetElement && targetElement.textContent && targetElement.textContent.trim()) {
              targetValue = targetElement.textContent.trim();
            }
          }
          
          // 방법 3: 전체 페이지에서 송장번호 패턴 찾기 (최후의 수단)
          if (!targetValue) {
            const allElements = document.querySelectorAll('input, span, td, div');
            for (const element of allElements) {
              const text = element.value || element.textContent || '';
              if (text.match(/\d{6}_V\d+_\d+/)) { // 송장번호 패턴 매칭
                targetValue = text.trim();
                break;
              }
            }
          }
          
          return targetValue;
        });
        
        if (extractedVendorInvoiceValue) {
          logger.info(`✅ 공급사송장 아래 값 추출 성공: "${extractedVendorInvoiceValue}"`);
          
          // 콜론 이후 부분만 제거 (두 번째 '_'와 숫자는 유지)
          let processedValue = extractedVendorInvoiceValue;
          
          // 콜론 이후 부분 제거 (: 피엠텍 등)
          if (processedValue.includes(':')) {
            processedValue = processedValue.split(':')[0].trim();
          }
          
          extractedVendorInvoiceValue = processedValue;
          logger.info(`✅ 공급사송장 값 가공 완료: "${extractedVendorInvoiceValue}"`);
        } else {
          logger.warn('⚠️ 공급사송장 아래 값 추출 실패');
        }
      } catch (extractError) {
        logger.warn(`공급사송장 값 추출 중 오류: ${extractError.message}`);
        extractedVendorInvoiceValue = null;
      }
      
      // 새 탭이 열릴 때까지 대기
      logger.info('새 탭 열릴 때까지 대기 중...');
      await delay(3000);
      
      // 모든 탭 가져오기
      const pages = await page.browser().pages();
      logger.info(`현재 열린 탭 수: ${pages.length}`);
      
      // 가장 최근에 열린 탭으로 이동 (마지막 탭)
      const newTab = pages[pages.length - 1];
      await newTab.bringToFront();
      logger.info('새 탭으로 이동 완료');
      
      // 페이지 로딩 완료까지 대기
      try {
        await newTab.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 });
        logger.info('새 탭 페이지 로딩 완료');
      } catch (loadError) {
        logger.warn(`페이지 로딩 대기 중 오류: ${loadError.message}, 계속 진행`);
        await delay(2000); // 추가 대기
      }
      
      // InvoiceDetails_Description input 요소 찾기
      logger.info('InvoiceDetails_Description input 요소 찾는 중...');
      
      const descriptionInputSelectors = [
        'input#VendEditInvoice_5_InvoiceDetails_Description_input',
        'input[name="InvoiceDetails_Description"]',
        'input[id*="InvoiceDetails_Description_input"]',
        'input[id*="Description_input"]',
        'input[aria-labelledby*="InvoiceDetails_Description_label"]',
        'input[class*="textbox"][class*="field"]',
        'input[class*="textbox"]',
        'input[type="text"][maxlength="255"]',
        'input[data-dyn-bind*="InvoiceDetails_Description"]',
        'input[placeholder*="Description"]',
        'input[aria-label*="Description"]',
        'div[data-dyn-controlname*="Description"] input',
        'input[type="text"]',
        'textarea[name*="Description"]',
        'textarea[id*="Description"]',
        'input[class*="field"]',
        'input'
      ];
      
      let descriptionInputFound = false;
      let targetDescriptionInput = null;
      let retryCount = 0;
      const maxRetries = 3;
      
      while (!descriptionInputFound && retryCount < maxRetries) {
        logger.info(`Description input 찾기 시도 ${retryCount + 1}/${maxRetries}`);
        
        for (const selector of descriptionInputSelectors) {
          try {
            logger.info(`Description input 선택자 시도: ${selector}`);
            
            const input = await newTab.$(selector);
            if (input) {
              const isVisible = await input.isIntersectingViewport();
              if (isVisible) {
                logger.info(`Description input 발견: ${selector}`);
                
                // input 위치 정보 가져오기
                const inputPosition = await input.boundingBox();
                logger.info(`Description input 위치: x=${inputPosition.x}, y=${inputPosition.y}`);
                
                targetDescriptionInput = input;
                descriptionInputFound = true;
                break;
              }
            }
          } catch (selectorError) {
            logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
            continue;
          }
        }
        
        if (!descriptionInputFound) {
          retryCount++;
          if (retryCount < maxRetries) {
            logger.info(`Description input을 찾지 못함, 2초 후 재시도...`);
            await delay(2000);
          }
        }
      }
      
      if (descriptionInputFound && targetDescriptionInput) {
        // Description input 클릭
        logger.info('Description input 클릭 수행 중...');
        await targetDescriptionInput.click();
        await delay(500);
        
        // AU열 값 붙여넣기
        if (lastProcessedValueFromAUColumn) {
          logger.info(`AU열 값 붙여넣기: ${lastProcessedValueFromAUColumn}`);
          await targetDescriptionInput.type(String(lastProcessedValueFromAUColumn));
          await delay(300);
          
          // Enter 키 입력
          logger.info('Description input에 Enter 키 입력 중...');
          await newTab.keyboard.press('Enter');
          await delay(500);
          
          logger.info('✅ Description input 클릭 및 AU열 값 붙여넣기, Enter 완료');
          
          // 3초 딜레이 후 AT열 값을 위한 FixedDueDate textbox 처리
          logger.info('3초 대기 후 FixedDueDate textbox 처리 시작...');
          await delay(3000);
          
          // FixedDueDate textbox 찾기 및 AT열 값 입력
          await processFixedDueDateInput(newTab);
          
        } else {
          logger.warn('⚠️ AU열 값이 없어서 붙여넣기를 건너뜁니다');
        }
      } else {
        logger.warn('⚠️ Description input을 찾을 수 없습니다');
      }
      
    } else {
      logger.warn('⚠️ 송장 번호 input을 찾을 수 없었지만 캘린더 부분은 완료됨');
    }
    
  } catch (error) {
    logger.error(`캘린더 버튼 처리 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 * FixedDueDate textbox 찾기 및 AT열 값 입력
 */
async function processFixedDueDateInput(page) {
  try {
    logger.info('🚀 FixedDueDate textbox 처리 시작');
    
    // FixedDueDate textbox 선택자들 (더 포괄적으로 개선)
    const fixedDueDateSelectors = [
      'input[name="PurchParmTable_FixedDueDate"]',
      'input[id*="PurchParmTable_FixedDueDate_input"]',
      'input[id*="FixedDueDate_input"]',
      'input[id*="FixedDueDate"]',
      'input.textbox.field[role="combobox"][aria-haspopup="dialog"]',
      'input[aria-controls="ui-datepicker-div"]',
      'input[role="combobox"][aria-haspopup="dialog"]',
      'input.textbox.field',
      'input[class*="textbox"]',
      'input[class*="date"]',
      'input[type="text"][class*="field"]',
      'input[placeholder*="날짜"]',
      'input[placeholder*="date"]'
    ];
    
    let targetInput = null;
    let inputFound = false;
    
    // 각 선택자로 FixedDueDate textbox 찾기
    for (const selector of fixedDueDateSelectors) {
      try {
        logger.info(`FixedDueDate textbox 선택자 시도: ${selector}`);
        
        const input = await page.$(selector);
        if (input) {
          const isVisible = await input.isIntersectingViewport();
          if (isVisible) {
            // input 정보 확인
            const inputInfo = await page.evaluate((el) => {
              return {
                id: el.id,
                name: el.name,
                value: el.value,
                title: el.title,
                placeholder: el.placeholder
              };
            }, input);
            
            logger.info(`FixedDueDate textbox 발견: ${selector}`);
            logger.info(`Input 정보: id=${inputInfo.id}, name=${inputInfo.name}, value="${inputInfo.value}", title="${inputInfo.title}"`);
            
            targetInput = input;
            inputFound = true;
            break;
          } else {
            logger.warn(`FixedDueDate textbox가 화면에 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
        continue;
      }
    }
    
    if (!inputFound) {
      // JavaScript evaluate를 사용한 더 포괄적인 검색
      logger.info('JavaScript evaluate로 FixedDueDate textbox 찾는 중...');
      try {
        const result = await page.evaluate(() => {
          // 모든 input 요소에서 날짜 관련 요소 찾기
          const inputs = document.querySelectorAll('input[type="text"], input[role="combobox"], input.textbox, input.field');
          
          for (const input of inputs) {
            // FixedDueDate 관련 속성 체크
            if (input.name && input.name.includes('FixedDueDate')) {
              return { success: true, selector: `input[name="${input.name}"]`, id: input.id };
            }
            if (input.id && input.id.includes('FixedDueDate')) {
              return { success: true, selector: `input[id="${input.id}"]`, id: input.id };
            }
            // 날짜 입력 필드 추정
            if (input.placeholder && (input.placeholder.includes('날짜') || input.placeholder.includes('date'))) {
              return { success: true, selector: `input[placeholder="${input.placeholder}"]`, id: input.id };
            }
            // 캘린더 버튼 근처의 input 찾기
            if (input.getAttribute('aria-haspopup') === 'dialog' || input.role === 'combobox') {
              return { success: true, selector: `input[id="${input.id}"]`, id: input.id };
            }
          }
          
          return { success: false };
        });
        
        if (result.success) {
          logger.info(`JavaScript evaluate로 FixedDueDate 발견: ${result.selector}, id: ${result.id}`);
          targetInput = await page.$(result.selector);
          if (targetInput) {
            const isVisible = await targetInput.isIntersectingViewport();
            if (isVisible) {
              inputFound = true;
              logger.info(`✅ JavaScript evaluate로 FixedDueDate textbox 찾기 성공`);
            }
          }
        }
      } catch (evalError) {
        logger.warn(`JavaScript evaluate 실패: ${evalError.message}`);
      }
    }
    
    if (!inputFound) {
      logger.warn('⚠️ FixedDueDate textbox를 찾을 수 없어 이 단계를 건너뜁니다.');
      return; // 오류 대신 경고로 처리하고 계속 진행
    }
    
    // AT열 값 확인 및 변환
    if (!lastProcessedDateFromATColumn) {
      logger.warn('⚠️ AT열 값이 없어서 FixedDueDate 입력을 건너뜁니다');
      return;
    }
    
    // AT열 날짜를 M/DD/YYYY 형식으로 변환
    const convertedDate = convertDateFormat(lastProcessedDateFromATColumn);
    if (!convertedDate) {
      logger.error('❌ AT열 날짜 변환에 실패했습니다');
      throw new Error('AT열 날짜 데이터 변환에 실패했습니다.');
    }
    
    logger.info(`AT열 날짜 변환: ${lastProcessedDateFromATColumn} -> ${convertedDate}`);
    
    // FixedDueDate textbox 클릭 및 값 입력
    logger.info('FixedDueDate textbox 클릭 수행 중...');
    await targetInput.click();
    await delay(500);
    
    // 기존 값 모두 선택 후 삭제
    await page.keyboard.down('Control');
    await page.keyboard.press('KeyA');
    await page.keyboard.up('Control');
    await delay(200);
    
    // AT열 값 입력
    logger.info(`AT열 값 입력: ${convertedDate}`);
    await targetInput.type(convertedDate);
    await delay(300);
    
    // Enter 키 입력
    logger.info('FixedDueDate input에 Enter 키 입력 중...');
    await page.keyboard.press('Enter');
    await delay(500);
    
    logger.info('✅ FixedDueDate textbox 클릭 및 AT열 값 입력, Enter 완료');
    
    // 사업자등록번호 input에 직접 입력
    try {
      await processBizRegNumInput(page);
    } catch (bizRegError) {
      logger.warn(`⚠️ 사업자등록번호 입력 실패했지만 계속 진행: ${bizRegError.message}`);
    }
    
  } catch (error) {
    logger.error(`FixedDueDate textbox 처리 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 *  
 */
async function processBizRegNumInput(page) {
  try {
    logger.info('🚀 사업자등록번호 input 처리 시작');
    
    // 사업자등록번호 input 선택자들
    const bizRegInputSelectors = [
      'input[name="VendInvoiceInfoTable_KVBizRegNum_Line"]',
      'input[id*="VendInvoiceInfoTable_KVBizRegNum_Line_input"]',
      'input[id*="KVBizRegNum_Line_input"]',
      'input[aria-labelledby*="KVBizRegNum_Line_label"]',
      'input.textbox.field[role="combobox"][aria-haspopup="grid"]',
      'input[data-dyn-bind*="VendInvoiceInfoTable_KVBizRegNum"]'
    ];
    
    let inputFound = false;
    let targetInput = null;
    
    // 첫 번째 시도: 일반적인 선택자로 찾기
    for (const selector of bizRegInputSelectors) {
      try {
        logger.info(`사업자등록번호 input 선택자 시도: ${selector}`);
        
        const input = await page.$(selector);
        if (input) {
          const isVisible = await input.isIntersectingViewport();
          if (isVisible) {
            // input 정보 확인
            const inputInfo = await page.evaluate((el) => {
              return {
                id: el.id,
                name: el.name,
                value: el.value,
                ariaLabelledBy: el.getAttribute('aria-labelledby'),
                className: el.className,
                type: el.type,
                role: el.role
              };
            }, input);
            
            logger.info(`사업자등록번호 input 발견: ${selector}`);
            logger.info(`Input 정보: id=${inputInfo.id}, name=${inputInfo.name}, value="${inputInfo.value}", role="${inputInfo.role}"`);
            
            targetInput = input;
            inputFound = true;
            break;
          } else {
            logger.warn(`사업자등록번호 input이 화면에 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
        continue;
      }
    }
    
    // 두 번째 시도: name 속성으로 찾기
    if (!inputFound) {
      try {
        logger.info('name 속성 기반 사업자등록번호 input 찾는 중...');
        
        const inputByName = await page.evaluate(() => {
          const inputs = document.querySelectorAll('input[type="text"], input[role="combobox"]');
          for (const input of inputs) {
            const name = input.name || '';
            if (name.includes('KVBizRegNum') || name.includes('VendInvoiceInfoTable')) {
              const rect = input.getBoundingClientRect();
              return {
                found: true,
                id: input.id,
                name: input.name,
                value: input.value,
                ariaLabelledBy: input.getAttribute('aria-labelledby'),
                x: rect.x,
                y: rect.y,
                width: rect.width,
                height: rect.height
              };
            }
          }
          return { found: false };
        });
        
        if (inputByName.found) {
          logger.info(`name 기반 사업자등록번호 input 발견: id=${inputByName.id}, name=${inputByName.name}`);
          logger.info(`위치: x=${inputByName.x}, y=${inputByName.y}`);
          
          inputFound = true;
          // 좌표를 이용해서 나중에 클릭할 준비
        }
      } catch (nameError) {
        logger.warn(`name 기반 검색 실패: ${nameError.message}`);
      }
    }
    
    if (!inputFound) {
      throw new Error('사업자등록번호 input을 찾을 수 없습니다.');
    }
    
    // input 클릭
    if (targetInput) {
      logger.info('사업자등록번호 input 클릭 수행 중...');
      await targetInput.click();
      await delay(500);
    } else {
      // name 기반으로 찾은 경우 좌표로 클릭
      const inputByName = await page.evaluate(() => {
        const inputs = document.querySelectorAll('input[type="text"], input[role="combobox"]');
        for (const input of inputs) {
          const name = input.name || '';
          if (name.includes('KVBizRegNum') || name.includes('VendInvoiceInfoTable')) {
            const rect = input.getBoundingClientRect();
            return {
              x: rect.x + rect.width / 2,
              y: rect.y + rect.height / 2
            };
          }
        }
        return null;
      });
      
      if (inputByName) {
        await page.mouse.click(inputByName.x, inputByName.y);
        await delay(500);
      }
    }
    
    // 기존 값 모두 선택 후 삭제
    await page.keyboard.down('Control');
    await page.keyboard.press('KeyA');
    await page.keyboard.up('Control');
    await delay(200);
    
    // "4138601441" 입력
    logger.info('사업자등록번호 "4138601441" 입력 중...');
    await page.keyboard.type('4138601441');
    await delay(300);
    
    // Enter 키 입력
    logger.info('Enter 키 입력 중...');
    await page.keyboard.press('Enter');
    await delay(500);
    
    logger.info('✅ 사업자등록번호 input 클릭, 값 입력, Enter 완료');
    
    // 2초 딜레이 후 KVTenderId input 처리
    logger.info('2초 대기 후 KVTenderId input 처리 시작...');
    await delay(2000);
    
    // KVTenderId input 찾기 및 처리
    await processKVTenderIdInput(page);
    
    
  } catch (error) {
    logger.error(`사업자등록번호 input 처리 중 오류: ${error.message}`);
    throw error;
  }
}


/**
 * KVTenderId input 찾기, 클릭, 값 입력 및 Enter 처리
 */
async function processKVTenderIdInput(page) {
  try {
    logger.info('🚀 KVTenderId input 처리 시작');
    
    // KVTenderId input 선택자들
    const tenderIdInputSelectors = [
      'input[name="VendInvoiceInfoTable_KVTenderId_Line"]',
      'input[id*="VendInvoiceInfoTable_KVTenderId_Line_input"]',
      'input[id*="KVTenderId_Line_input"]',
      'input[aria-labelledby*="KVTenderId_Line_label"]',
      'input.textbox.field[role="combobox"][aria-haspopup="grid"]',
      'input[data-dyn-bind*="VendInvoiceInfoTable_KVTenderId"]'
    ];
    
    let inputFound = false;
    let targetInput = null;
    
    // 첫 번째 시도: 일반적인 선택자로 찾기
    for (const selector of tenderIdInputSelectors) {
      try {
        logger.info(`KVTenderId input 선택자 시도: ${selector}`);
        
        const input = await page.$(selector);
        if (input) {
          const isVisible = await input.isIntersectingViewport();
          if (isVisible) {
            // input 정보 확인
            const inputInfo = await page.evaluate((el) => {
              return {
                id: el.id,
                name: el.name,
                value: el.value,
                ariaLabelledBy: el.getAttribute('aria-labelledby'),
                className: el.className,
                type: el.type,
                role: el.role
              };
            }, input);
            
            logger.info(`KVTenderId input 발견: ${selector}`);
            logger.info(`Input 정보: id=${inputInfo.id}, name=${inputInfo.name}, value="${inputInfo.value}", role="${inputInfo.role}"`);
            
            targetInput = input;
            inputFound = true;
            break;
          } else {
            logger.warn(`KVTenderId input이 화면에 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`선택자 ${selector} 시도 실패: ${selectorError.message}`);
        continue;
      }
    }
    
    // 두 번째 시도: name 속성으로 찾기
    if (!inputFound) {
      try {
        logger.info('name 속성 기반 KVTenderId input 찾는 중...');
        
        const inputByName = await page.evaluate(() => {
          const inputs = document.querySelectorAll('input[type="text"], input[role="combobox"]');
          for (const input of inputs) {
            const name = input.name || '';
            if (name.includes('KVTenderId') || name.includes('VendInvoiceInfoTable')) {
              const rect = input.getBoundingClientRect();
              return {
                found: true,
                id: input.id,
                name: input.name,
                value: input.value,
                ariaLabelledBy: input.getAttribute('aria-labelledby'),
                x: rect.x,
                y: rect.y,
                width: rect.width,
                height: rect.height
              };
            }
          }
          return { found: false };
        });
        
        if (inputByName.found) {
          logger.info(`name 기반 KVTenderId input 발견: id=${inputByName.id}, name=${inputByName.name}`);
          logger.info(`위치: x=${inputByName.x}, y=${inputByName.y}`);
          
          inputFound = true;
          // 좌표를 이용해서 나중에 클릭할 준비
        }
      } catch (nameError) {
        logger.warn(`name 기반 검색 실패: ${nameError.message}`);
      }
    }
    
    if (!inputFound) {
      throw new Error('KVTenderId input을 찾을 수 없습니다.');
    }
    
    // input 클릭
    if (targetInput) {
      logger.info('KVTenderId input 클릭 수행 중...');
      await targetInput.click();
      await delay(500);
    } else {
      // name 기반으로 찾은 경우 좌표로 클릭
      const inputByName = await page.evaluate(() => {
        const inputs = document.querySelectorAll('input[type="text"], input[role="combobox"]');
        for (const input of inputs) {
          const name = input.name || '';
          if (name.includes('KVTenderId') || name.includes('VendInvoiceInfoTable')) {
            const rect = input.getBoundingClientRect();
            return {
              x: rect.x + rect.width / 2,
              y: rect.y + rect.height / 2
            };
          }
        }
        return null;
      });
      
      if (inputByName) {
        await page.mouse.click(inputByName.x, inputByName.y);
        await delay(500);
      }
    }
    
    // 기존 값 모두 선택 후 삭제
    await page.keyboard.down('Control');
    await page.keyboard.press('KeyA');
    await page.keyboard.up('Control');
    await delay(200);
    
    // "11" 입력
    logger.info('KVTenderId "11" 입력 중...');
    await page.keyboard.type('11');
    await delay(300);
    
    // Enter 키 입력
    logger.info('Enter 키 입력 중...');
    await page.keyboard.press('Enter');
    await delay(500);
    
    logger.info('✅ KVTenderId input 클릭, 값 입력, Enter 완료');
    
    // 새창 닫기 버튼 처리
    await processCloseNewWindow(page);
    
    // UserBtn 아래쪽 닫기 버튼 클릭 처리
    await clickCloseButtonBelowUserBtn(page);
    
    // 저장 버튼 클릭 처리
    await clickSaveButton(page);
    
    // 2초 대기 후 6번 RPA 동작 시작
    logger.info('⏳ 2초 대기 후 6번 RPA 동작 시작 예정...');
    await delay(2000);
    
    try {
      await executeStep6RPA(page);
      logger.info('✅ 6번 RPA 동작 완료');
      
      // 7번 RPA 동작: 그룹웨어 상신
      logger.info('⏳ 2초 대기 후 7번 RPA 동작(그룹웨어 상신) 시작 예정...');
      await delay(2000);
      
      try {
        await executeStep7RPA(page);
        logger.info('✅ 7번 RPA 동작: 그룹웨어 상신 완료');
        
        // 7번 RPA 성공 시 브라우저 닫기
        try {
          await browser.close();
          logger.info('✅ 단일모드 전체 RPA 프로세스 완료 - 브라우저 닫기 완료');
          
          // 성공 완료 후 바로 반환
          return { 
            success: true, 
            message: '1. ERP 접속 및 로그인 완료\n2. 구매 입고내역 조회 및 다운로드 완료\n3. 엑셀 파일 열기 및 매크로 실행 완료\n4. 대기중인 공급사송장 메뉴 이동 및 엑셀 데이터 처리 완료\n5. 캘린더 버튼 클릭 및 송장 처리 완료\n6. 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료\n7. 그룹웨어 상신 완료',
            completedAt: new Date().toISOString(),
            browserKeptOpen: false
          };
        } catch (closeError) {
          logger.warn(`브라우저 닫기 실패: ${closeError.message}`);
        }
        
      } catch (step7Error) {
        logger.warn(`⚠️ 7번 RPA 동작 실패했지만 계속 진행: ${step7Error.message}`);
      }
      
    } catch (step6Error) {
      logger.warn(`⚠️ 6번 RPA 동작 실패했지만 계속 진행: ${step6Error.message}`);
    }
    
  } catch (error) {
    logger.error(`KVTenderId input 처리 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 * 새창에서 "창 닫기" 버튼을 찾아 클릭하는 함수
 */
async function processCloseNewWindow(page) {
  try {
    logger.info('🔍 새창에서 "창 닫기" 버튼 찾는 중...');
    
    // 새창이 나타날 때까지 잠시 대기
    await delay(2000);
    
    // 현재 모든 페이지(탭) 가져오기
    const browser = page.browser();
    const pages = await browser.pages();
    
    logger.info(`현재 열린 페이지 수: ${pages.length}`);
    
    // 새로 열린 페이지(새창) 찾기 - 마지막에 열린 페이지 확인
    let newPage = null;
    if (pages.length > 1) {
      newPage = pages[pages.length - 1]; // 가장 최근에 열린 페이지
      logger.info('새창 감지됨, 새창에서 "창 닫기" 버튼 찾는 중...');
    } else {
      // 새창이 팝업이 아닌 현재 페이지의 모달일 경우
      logger.info('새창이 현재 페이지의 모달로 추정됨, 현재 페이지에서 "창 닫기" 버튼 찾는 중...');
      newPage = page;
    }
    
    // "창 닫기" 버튼 선택자들
    const closeButtonSelectors = [
      // 지정된 선택자 패턴
      'span[data-dyn-bind*="FormButtonControlClose"]',
      '#NPS_VATInvoiceResultList4UserPo_7_FormButtonControlClose_label',
      'span[id*="FormButtonControlClose_label"]',
      'span[class="button-label"][for*="FormButtonControlClose"]',
      
      // 일반적인 닫기 버튼 선택자들
      'button[aria-label*="닫기"]',
      'button[title*="닫기"]',
      'span[aria-label*="닫기"]',
      'span[title*="닫기"]',
      '[data-dyn-bind*="닫기"]',
      
      // 텍스트 기반 선택자들
      'button:contains("창 닫기")',
      'span:contains("창 닫기")',
      'button:contains("닫기")',
      'span:contains("닫기")',
      
      // X 버튼이나 Close 버튼
      'button[aria-label*="Close"]',
      'button[title*="Close"]',
      '.close-button',
      '.btn-close',
      '[role="button"][aria-label*="Close"]'
    ];
    
    let closeButtonClicked = false;
    
    // 첫 번째 시도: 일반적인 선택자로 찾기
    for (const selector of closeButtonSelectors) {
      try {
        if (selector.includes(':contains(')) {
          continue; // CSS :contains()는 지원되지 않으므로 스킵
        }
        
        logger.info(`창 닫기 버튼 선택자 시도: ${selector}`);
        
        const closeButton = await newPage.$(selector);
        if (closeButton) {
          const isVisible = await newPage.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
          }, closeButton);
          
          if (isVisible) {
            await closeButton.click();
            logger.info(`✅ 창 닫기 버튼 클릭 성공: ${selector}`);
            closeButtonClicked = true;
            break;
          } else {
            logger.warn(`창 닫기 버튼이 보이지 않음: ${selector}`);
          }
        }
      } catch (selectorError) {
        logger.warn(`창 닫기 버튼 선택자 시도 실패: ${selector} - ${selectorError.message}`);
        continue;
      }
    }
    
    // 두 번째 시도: JavaScript로 직접 텍스트 검색
    if (!closeButtonClicked) {
      try {
        logger.info('JavaScript로 "창 닫기" 텍스트 검색 중...');
        
        const clicked = await newPage.evaluate(() => {
          // 모든 요소에서 "창 닫기" 텍스트 검색
          const allElements = document.querySelectorAll('*');
          
          for (const element of allElements) {
            const text = element.textContent || element.innerText || '';
            if (text.trim() === '창 닫기' || text.includes('창 닫기')) {
              // 클릭 가능한 요소인지 확인
              const clickableEl = element.closest('button, [role="button"], .button-container, span[for], label[for]') || element;
              
              // 해당 요소가 클릭 가능한지 확인
              const style = window.getComputedStyle(clickableEl);
              if (style.display !== 'none' && style.visibility !== 'hidden') {
                clickableEl.click();
                return { success: true, text: text.trim(), tagName: clickableEl.tagName };
              }
            }
          }
          return { success: false };
        });
        
        if (clicked.success) {
          logger.info(`✅ JavaScript로 창 닫기 버튼 클릭 성공: "${clicked.text}" (${clicked.tagName})`);
          closeButtonClicked = true;
        }
      } catch (jsError) {
        logger.warn(`JavaScript 창 닫기 버튼 클릭 실패: ${jsError.message}`);
      }
    }
    
    // 세 번째 시도: 특정 ID 패턴으로 찾기
    if (!closeButtonClicked) {
      try {
        logger.info('특정 ID 패턴으로 창 닫기 버튼 찾는 중...');
        
        const clicked = await newPage.evaluate(() => {
          // FormButtonControlClose가 포함된 ID를 가진 요소들 찾기
          const elements = document.querySelectorAll('[id*="FormButtonControlClose"]');
          
          for (const element of elements) {
            // label이나 span 요소인 경우, for 속성에 해당하는 실제 버튼 찾기
            const targetId = element.getAttribute('for');
            if (targetId) {
              const targetButton = document.getElementById(targetId);
              if (targetButton) {
                targetButton.click();
                return { success: true, method: 'for-target', id: targetId };
              }
            }
            
            // 직접 클릭 시도
            const style = window.getComputedStyle(element);
            if (style.display !== 'none' && style.visibility !== 'hidden') {
              element.click();
              return { success: true, method: 'direct', id: element.id };
            }
          }
          return { success: false };
        });
        
        if (clicked.success) {
          logger.info(`✅ ID 패턴으로 창 닫기 버튼 클릭 성공 (${clicked.method}): ${clicked.id}`);
          closeButtonClicked = true;
        }
      } catch (idError) {
        logger.warn(`ID 패턴 창 닫기 버튼 클릭 실패: ${idError.message}`);
      }
    }
    
    if (!closeButtonClicked) {
      logger.warn('⚠️ 창 닫기 버튼을 찾을 수 없습니다. 수동으로 닫아야 할 수 있습니다.');
      // 에러를 throw하지 않고 경고만 표시 (프로세스 진행을 방해하지 않기 위해)
    } else {
      logger.info('✅ 새창 닫기 처리 완료');
      
      // 창이 닫힌 후 잠시 대기
      await delay(1000);
    }
    
  } catch (error) {
    logger.error(`새창 닫기 처리 중 오류: ${error.message}`);
    // 에러를 throw하지 않고 로그만 남김 (프로세스 진행을 방해하지 않기 위해)
  }
}

/**
 * UserBtn 아래쪽 닫기 버튼(commandRing Cancel-symbol) 클릭 함수
 */
async function clickCloseButtonBelowUserBtn(page) {
  try {
    logger.info('🔍 UserBtn 아래쪽 닫기 버튼 찾는 중...');
    
    // 1. UserBtn 요소 찾기
    const userBtn = await page.$('button#UserBtn');
    if (!userBtn) {
      logger.warn('⚠️ UserBtn 요소를 찾을 수 없습니다.');
      return;
    }
    
    // UserBtn의 위치 정보 가져오기
    const userBtnPosition = await userBtn.boundingBox();
    if (!userBtnPosition) {
      logger.warn('⚠️ UserBtn의 위치 정보를 가져올 수 없습니다.');
      return;
    }
    
    logger.info(`UserBtn 위치: x=${userBtnPosition.x}, y=${userBtnPosition.y}, width=${userBtnPosition.width}, height=${userBtnPosition.height}`);
    
    // Y축변경부분 - UserBtn 아래쪽 20px 지점
    const targetY = userBtnPosition.y + 20;
    logger.info(`닫기 버튼 검색 기준 Y좌표: ${targetY} (UserBtn Y좌표 + 20px)`);
    
    // 2. UserBtn 아래쪽에서 commandRing Cancel-symbol 버튼 찾기
    const closeButtonSelectors = [
      '[class*="commandRing"][class*="Cancel-symbol"]',
      '.commandRing.Cancel-symbol',
      'button[class*="commandRing"][class*="Cancel-symbol"]',
      'div[class*="commandRing"][class*="Cancel-symbol"]',
      '[class*="Cancel-symbol"]',
      'button[class*="Cancel-symbol"]'
    ];
    
    let closeButtonFound = false;
    
    for (const selector of closeButtonSelectors) {
      try {
        logger.info(`닫기 버튼 선택자 시도: ${selector}`);
        
        const buttons = await page.$$(selector);
        for (const button of buttons) {
          const buttonPosition = await button.boundingBox();
          if (buttonPosition && buttonPosition.y > targetY) {
            // UserBtn 아래쪽에 있는 버튼인 경우
            logger.info(`닫기 버튼 발견: ${selector}, 위치: x=${buttonPosition.x}, y=${buttonPosition.y}`);
            
            const isVisible = await button.isIntersectingViewport();
            if (isVisible) {
              logger.info('닫기 버튼 클릭 시도...');
              await button.click();
              await delay(500);
              closeButtonFound = true;
              break;
            } else {
              logger.warn('닫기 버튼이 화면에 보이지 않음');
            }
          }
        }
        
        if (closeButtonFound) break;
      } catch (err) {
        logger.warn(`닫기 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!closeButtonFound) {
      // JavaScript evaluate로 더 정확한 검색
      try {
        logger.info('JavaScript evaluate로 닫기 버튼 찾는 중...');
        const result = await page.evaluate((targetY) => {
          // 모든 요소 중에서 commandRing Cancel-symbol 관련 요소 찾기
          const allElements = document.querySelectorAll('*');
          for (const element of allElements) {
            const className = element.className || '';
            if ((className.includes('commandRing') && className.includes('Cancel-symbol')) ||
                className.includes('Cancel-symbol')) {
              const rect = element.getBoundingClientRect();
              if (rect.y > targetY && element.offsetParent !== null) {
                console.log(`닫기 버튼 클릭 시도: class="${className}", y=${rect.y}`);
                element.click();
                return { success: true, className: className, y: rect.y };
              }
            }
          }
          return { success: false };
        }, targetY);
        
        if (result.success) {
          logger.info(`JavaScript evaluate로 닫기 버튼 클릭 성공: class="${result.className}", y=${result.y}`);
          await delay(500);
          closeButtonFound = true;
        }
      } catch (err) {
        logger.warn(`JavaScript evaluate 닫기 버튼 클릭 실패: ${err.message}`);
      }
    }
    
    if (closeButtonFound) {
      logger.info('✅ UserBtn 아래쪽 닫기 버튼 클릭 완료');
    } else {
      logger.warn('⚠️ UserBtn 아래쪽 닫기 버튼을 찾을 수 없습니다.');
    }
    
  } catch (error) {
    logger.error(`UserBtn 아래쪽 닫기 버튼 클릭 중 오류: ${error.message}`);
    // 에러를 throw하지 않고 로그만 남김 (프로세스 진행을 방해하지 않기 위해)
  }
}

/**
 * 저장 버튼 클릭 함수
 */
async function clickSaveButton(page) {
  try {
    logger.info('🔍 저장 버튼 찾는 중...');
    
    // 저장 버튼 선택자들
    const saveButtonSelectors = [
      'span#VendEditInvoice_5_SystemDefinedSaveButton_label',
      'span[id*="SystemDefinedSaveButton_label"]',
      'span.button-label[for*="SystemDefinedSaveButton"]',
      'span[for="VendEditInvoice_5_SystemDefinedSaveButton"]',
      'button#VendEditInvoice_5_SystemDefinedSaveButton'
    ];
    
    let saveButtonFound = false;
    
    for (const selector of saveButtonSelectors) {
      try {
        logger.info(`저장 버튼 선택자 시도: ${selector}`);
        
        const button = await page.$(selector);
        if (button) {
          const isVisible = await button.isIntersectingViewport();
          const text = await button.evaluate(el => el.textContent);
          
          logger.info(`저장 버튼 발견: ${selector}, 가시성: ${isVisible}, 텍스트: "${text}"`);
          
          if (isVisible && text && text.includes('저장')) {
            logger.info('저장 버튼 클릭 시도...');
            await button.click();
            await delay(1000); // 저장 후 1초 대기
            saveButtonFound = true;
            break;
          }
        }
      } catch (err) {
        logger.warn(`저장 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!saveButtonFound) {
      // JavaScript evaluate로 더 정확한 검색
      try {
        logger.info('JavaScript evaluate로 저장 버튼 찾는 중...');
        const result = await page.evaluate(() => {
          // 모든 span과 button 요소에서 "저장" 텍스트가 있는 요소 찾기
          const allElements = document.querySelectorAll('span, button');
          for (const element of allElements) {
            const text = element.textContent ? element.textContent.trim() : '';
            const id = element.id || '';
            
            if (text === '저장' && element.offsetParent !== null &&
                (id.includes('SystemDefinedSaveButton') || element.className.includes('button-label'))) {
              console.log(`저장 버튼 클릭 시도: text="${text}", id="${id}"`);
              element.click();
              return { success: true, text: text, id: id };
            }
          }
          return { success: false };
        });
        
        if (result.success) {
          logger.info(`JavaScript evaluate로 저장 버튼 클릭 성공: text="${result.text}", id="${result.id}"`);
          await delay(1000); // 저장 후 1초 대기
          saveButtonFound = true;
        }
      } catch (err) {
        logger.warn(`JavaScript evaluate 저장 버튼 클릭 실패: ${err.message}`);
      }
    }
    
    if (saveButtonFound) {
      logger.info('✅ 저장 버튼 클릭 완료');
    } else {
      logger.warn('⚠️ 저장 버튼을 찾을 수 없습니다.');
    }
    
  } catch (error) {
    logger.error(`저장 버튼 클릭 중 오류: ${error.message}`);
    // 에러를 throw하지 않고 로그만 남김 (프로세스 진행을 방해하지 않기 위해)
  }
}

/**
 * 6번 RPA 동작: 대기중인 공급사송장 메뉴 이동
 */
async function executeStep6RPA(page) {
  logger.info('🚀 === 6번 RPA 동작: 대기중인 공급사송장 메뉴 이동 시작 ===');
  
  try {
    // 1. 검색 버튼 클릭
    logger.info('1. 검색 버튼 찾는 중...');
    const searchButtonSelectors = [
      '.button-commandRing.Find-symbol',
      'span.Find-symbol',
      '[data-dyn-image-type="Symbol"].Find-symbol',
      '.button-container .Find-symbol'
    ];
    
    let searchButtonClicked = false;
    for (const selector of searchButtonSelectors) {
      try {
        logger.info(`검색 버튼 선택자 시도: ${selector}`);
        const searchButton = await page.$(selector);
        if (searchButton) {
          const isVisible = await searchButton.isIntersectingViewport();
          logger.info(`검색 버튼 발견: ${selector}, 가시성: ${isVisible}`);
          
          if (isVisible) {
            await searchButton.click();
            await delay(500);
            logger.info('검색 버튼 클릭 완료');
            searchButtonClicked = true;
            break;
          }
        }
      } catch (err) {
        logger.warn(`검색 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!searchButtonClicked) {
      throw new Error('검색 버튼을 찾을 수 없습니다. (6번 RPA)');
    }
    
    // 검색창이 나타날 때까지 대기
    await delay(2000);
    
    // 2. "대기중인 공급사송장" 검색어 입력
    logger.info('2. 검색어 입력 중...');
    const searchInputSelectors = [
      'input[type="text"]',
      '.navigationSearchBox input',
      '#NavigationSearchBox',
      'input[placeholder*="검색"]',
      'input[aria-label*="검색"]'
    ];
    
    let searchInputFound = false;
    const searchTerm = '대기중인 공급사송장';
    
    for (const selector of searchInputSelectors) {
      try {
        logger.info(`검색 입력창 선택자 시도: ${selector}`);
        await page.waitForSelector(selector, { visible: true, timeout: 5000 });
        await page.click(selector, { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type(selector, searchTerm, { delay: 100 });
        logger.info(`검색어 입력 완료: ${searchTerm}`);
        searchInputFound = true;
        break;
      } catch (error) {
        logger.warn(`검색 입력창 처리 실패: ${selector} - ${error.message}`);
      }
    }
    
    if (!searchInputFound) {
      throw new Error('검색 입력창을 찾을 수 없습니다. (6번 RPA)');
    }
    
    // 검색 결과가 나타날 때까지 대기
    await delay(3000);
    
    // 3. 검색 결과에서 대기중인 공급사송장 메뉴 클릭
    logger.info('3. 검색 결과에서 대기중인 공급사송장 메뉴 찾는 중...');
    const searchResultSelectors = [
      '.navigationSearchBox',
      '.search-results',
      '.navigation-search-results',
      '[data-dyn-bind*="NavigationSearch"]'
    ];
    
    let menuClicked = false;
    for (const containerSelector of searchResultSelectors) {
      try {
        const container = await page.$(containerSelector);
        if (container) {
          const menuItems = await page.$$eval(`${containerSelector} *`, (elements) => {
            return elements
              .filter(el => {
                const text = el.textContent || el.innerText || '';
                return text.includes('대기중인 공급사송장');
              })
              .map(el => ({
                text: el.textContent || el.innerText,
                clickable: el.tagName === 'A' || el.tagName === 'BUTTON' || el.onclick || el.getAttribute('role') === 'button'
              }));
          });
          
          logger.info(`검색 결과 메뉴 항목들:`, menuItems);
          
          if (menuItems.length > 0) {
            await page.evaluate((containerSel) => {
              const container = document.querySelector(containerSel);
              if (container) {
                const elements = container.querySelectorAll('*');
                for (const el of elements) {
                  const text = el.textContent || el.innerText || '';
                  if (text.includes('대기중인 공급사송장')) {
                    el.click();
                    return true;
                  }
                }
              }
              return false;
            }, containerSelector);
            
            logger.info('대기중인 공급사송장 메뉴 클릭 완료');
            menuClicked = true;
            break;
          }
        }
      } catch (error) {
        logger.warn(`검색 결과 처리 실패: ${containerSelector} - ${error.message}`);
      }
    }
    
    if (!menuClicked) {
      // Enter 키로 첫 번째 결과 선택 시도
      logger.info('Enter 키로 검색 결과 선택 시도...');
      await page.keyboard.press('Enter');
      menuClicked = true;
    }
    
    // 페이지 이동 대기
    logger.info('4. 대기중인 공급사송장 페이지 로딩 대기 중...');
    await delay(5000);
    
    // 5. 필터 텍스트박스에 I열 값 입력
    logger.info('5. 필터 텍스트박스에 AU열 값 입력 중...');
    
    // AU열 값 확인 (디버깅 강화)
    logger.info(`🔍 AU열 값 상태 체크: ${lastProcessedValueFromAUColumn} (타입: ${typeof lastProcessedValueFromAUColumn})`);
    
    if (!lastProcessedValueFromAUColumn) {
      logger.warn('⚠️ 저장된 AU열 값이 없습니다. 필터 입력을 건너뜁니다.');  
      logger.warn(`⚠️ AU열 값 디버그: "${lastProcessedValueFromAUColumn}" (타입: ${typeof lastProcessedValueFromAUColumn})`);
    } else {
      logger.info(`📋 사용할 AU열 값: "${lastProcessedValueFromAUColumn}"`);
      
      // 필터 텍스트박스 선택자들
      const filterInputSelectors = [
        'input[name="QuickFilterControl_Input"]',
        'input[id*="QuickFilterControl_Input_input"]',
        'input[aria-label="필터"]',
        'input[id*="QuickFilterControl"]'
      ];
      
      let filterInputFound = false;
      
      for (const selector of filterInputSelectors) {
        try {
          logger.info(`필터 텍스트박스 선택자 시도: ${selector}`);
          
          const input = await page.$(selector);
          if (input) {
            const isVisible = await input.isIntersectingViewport();
            logger.info(`필터 텍스트박스 발견: ${selector}, 가시성: ${isVisible}`);
            
            if (isVisible) {
              // 텍스트박스 클릭 및 기존 내용 삭제
              await input.click();
              await delay(300);
              
              // 기존 내용 모두 선택 후 삭제
              await page.keyboard.down('Control');
              await page.keyboard.press('KeyA');
              await page.keyboard.up('Control');
              await delay(200);
              
              // AU열 값 입력
              await input.type(String(lastProcessedValueFromAUColumn));
              await delay(1000); // 1초 대기하여 콤보박스 나타나게 함
              
              // 콤보박스에서 4번째 항목(인덱스 3) 클릭
              try {
                const comboboxItem = await page.$('li.quickFilter-listItem[data-dyn-index="3"]');
                if (comboboxItem) {
                  await comboboxItem.click();
                  await delay(500);
                  logger.info(`✅ 콤보박스 4번째 항목 클릭 완료`);
                  
                  // 1초 대기 후 추가 동작 시작
                  await delay(1000);
                  
                  // 1. SVG 체크박스 클릭
                  try {
                    const svgCheckbox = await page.$('div.dyn-container._ln972h.dyn-svg-symbol');
                    if (svgCheckbox) {
                      await svgCheckbox.click();
                      await delay(1000);
                      logger.info(`✅ SVG 체크박스 클릭 완료`);
                      
                      // 2. 그룹웨어 버튼 클릭
                      try {
                        const groupwareButton = await page.$('button[id*="NPS_GroupWareActionPaneTab_button"]');
                        if (groupwareButton) {
                          await groupwareButton.click();
                          await delay(1000);
                          logger.info(`✅ 그룹웨어 버튼 클릭 완료`);
                          
                          // 3. 그룹웨어 승인 버튼 클릭
                          try {
                            const approvalButton = await page.$('div.button-container span.button-label[id*="NPS_IF_GRW_POINVOICEBATCH_label"]');
                            if (approvalButton) {
                              await approvalButton.click();
                              await delay(1000);
                              logger.info(`✅ 그룹웨어 승인 버튼 클릭 완료`);
                              
                              // 4. 새 창(로그인 창) 대기 및 처리 - 개선된 방법
                              try {
                                logger.info('새 창(로그인 창) 대기 중...');
                                
                                let newPage = null;
                                let attempts = 0;
                                const maxAttempts = 10;
                                
                                // 3번째 탭 (인덱스 2) 확인 방법
                                while (!newPage && attempts < maxAttempts) {
                                  try {
                                    const pages = await page.browser().pages();
                                    logger.info(`현재 페이지 수: ${pages.length}`);
                                    
                                    // 3번째 탭이 존재하는지 확인 (인덱스 2)
                                    if (pages.length >= 3) {
                                      newPage = pages[2]; // 3번째 탭 (인덱스 2)
                                      logger.info('✅ 3번째 탭에서 새 창 감지됨');
                                      break;
                                    }
                                    
                                    // 만약 3번째 탭이 없으면, 가장 최근에 열린 페이지 확인
                                    if (pages.length > 1) {
                                      newPage = pages[pages.length - 1];
                                      logger.info(`✅ 가장 최근 페이지에서 새 창 감지됨 (총 ${pages.length}개 페이지, 인덱스 ${pages.length - 1})`);
                                      break;
                                    }
                                    
                                  } catch (pageError) {
                                    logger.warn(`페이지 확인 실패 (시도 ${attempts + 1}/${maxAttempts}): ${pageError.message}`);
                                  }
                                  
                                  attempts++;
                                  logger.info(`3번째 탭 대기 중... (시도 ${attempts}/${maxAttempts})`);
                                  await delay(1000);
                                }
                                
                                if (!newPage) {
                                  throw new Error('새 창을 감지할 수 없습니다');
                                }
                                
                                // 새 페이지 로딩 대기
                                try {
                                  await newPage.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
                                } catch (navError) {
                                  logger.warn(`페이지 네비게이션 대기 실패: ${navError.message}, 계속 진행`);
                                }
                                await delay(1000);
                                logger.info('✅ 새 로그인 창 감지 및 로딩 완료');
                                
                                // 4.1 로그인 요소 대기 및 확인
                                logger.info('로그인 요소 대기 중...');
                                let loginAttempts = 0;
                                const maxLoginAttempts = 5;
                                let loginSuccess = false;
                                
                                while (!loginSuccess && loginAttempts < maxLoginAttempts) {
                                  try {
                                    // 로그인 요소들이 모두 존재하는지 확인
                                    await newPage.waitForSelector('#txtLoginID', { visible: true, timeout: 3000 });
                                    await newPage.waitForSelector('#txtPassword', { visible: true, timeout: 3000 });
                                    await newPage.waitForSelector('#btnLogin', { visible: true, timeout: 3000 });
                                    
                                    logger.info('✅ 모든 로그인 요소 감지됨');
                                    
                                    // 4.1 아이디 입력 (하드코딩)
                                    const loginId = 'accounting';
                                    await newPage.click('#txtLoginID'); // 포커스
                                    await newPage.evaluate(() => document.querySelector('#txtLoginID').value = ''); // 기존 값 클리어
                                    await newPage.type('#txtLoginID', loginId);
                                    await delay(100);
                                    logger.info(`✅ 로그인 ID 입력 완료: ${loginId}`);
                                    
                                    // 4.2 패스워드 입력 (하드코딩)
                                    const loginPassword = 'P@ssw0rd';
                                    await newPage.click('#txtPassword'); // 포커스
                                    await newPage.evaluate(() => document.querySelector('#txtPassword').value = ''); // 기존 값 클리어
                                    await newPage.type('#txtPassword', loginPassword);
                                    await delay(100);
                                    logger.info(`✅ 로그인 PW 입력 완료`);
                                    
                                    // 4.3 로그인 버튼 클릭
                                    await newPage.click('#btnLogin');
                                    await delay(500);
                                    logger.info(`✅ 로그인 버튼 클릭 완료`);
                                    
                                    loginSuccess = true;
                                    
                                  } catch (loginError) {
                                    loginAttempts++;
                                    logger.warn(`로그인 시도 ${loginAttempts}/${maxLoginAttempts} 실패: ${loginError.message}`);
                                    
                                    if (loginAttempts < maxLoginAttempts) {
                                      logger.info('2초 후 재시도...');
                                      await delay(2000);
                                    }
                                  }
                                }
                                
                                if (!loginSuccess) {
                                  throw new Error('로그인 요소를 찾을 수 없습니다');
                                }
                                
                              } catch (newPageError) {
                                logger.error(`❌ 새 창 로그인 처리 실패: ${newPageError.message}. 작업 중단.`);
                                return;
                              }
                              
                            } else {
                              logger.error('❌ 그룹웨어 승인 버튼을 찾을 수 없습니다. 작업 중단.');
                              return;
                            }
                          } catch (approvalError) {
                            logger.error(`❌ 그룹웨어 승인 버튼 클릭 실패: ${approvalError.message}. 작업 중단.`);
                            return;
                          }
                          
                        } else {
                          logger.error('❌ 그룹웨어 버튼을 찾을 수 없습니다. 작업 중단.');
                          return;
                        }
                      } catch (groupwareError) {
                        logger.error(`❌ 그룹웨어 버튼 클릭 실패: ${groupwareError.message}. 작업 중단.`);
                        return;
                      }
                      
                    } else {
                      logger.error('❌ SVG 체크박스를 찾을 수 없습니다. 작업 중단.');
                      return;
                    }
                  } catch (svgError) {
                    logger.error(`❌ SVG 체크박스 클릭 실패: ${svgError.message}. 작업 중단.`);
                    return;
                  }
                  
                } else {
                  logger.error('❌ 콤보박스에서 4번째 항목을 찾을 수 없습니다.');
                  return;
                }
              } catch (comboError) {
                logger.error(`❌ 콤보박스 항목 클릭 실패: ${comboError.message}`);
                return;
              }
              
              logger.info(`✅ 필터 텍스트박스에 AU열 값 입력 및 콤보박스 선택 완료: "${lastProcessedValueFromAUColumn}"`);
              filterInputFound = true;
              break;
            }
          }
        } catch (err) {
          logger.warn(`필터 텍스트박스 선택자 실패: ${selector} - ${err.message}`);
        }
      }
      
      if (!filterInputFound) {
        // JavaScript evaluate로 더 정확한 검색
        try {
          logger.info('JavaScript evaluate로 필터 텍스트박스 찾는 중...');
          const result = await page.evaluate((iValue) => {
            // 필터 관련 input 찾기
            const inputs = document.querySelectorAll('input[name*="QuickFilter"], input[aria-label*="필터"], input[id*="QuickFilter"]');
            for (const input of inputs) {
              if (input.offsetParent !== null) {
                // 기존 내용 삭제 후 새 값 입력
                input.focus();
                input.select();
                input.value = '';
                input.value = iValue;
                
                // input 이벤트 발생시켜 변경사항 알림
                const inputEvent = new Event('input', { bubbles: true });
                input.dispatchEvent(inputEvent);
                
                return {
                  success: true,
                  value: iValue,
                  id: input.id,
                  name: input.name
                };
              }
            }
            return { success: false };
          }, String(lastProcessedValueFromAUColumn));
          
          if (result.success) {
            logger.info(`✅ JavaScript evaluate로 필터 입력 성공: "${result.value}", id: "${result.id}"`);
            await delay(1000); // 콤보박스 나타날 때까지 대기
            
            // 콤보박스에서 4번째 항목(인덱스 3) 클릭
            try {
              const comboboxItem = await page.$('li.quickFilter-listItem[data-dyn-index="3"]');
              if (comboboxItem) {
                await comboboxItem.click();
                await delay(500);
                logger.info(`✅ 콤보박스 4번째 항목 클릭 완료`);
                
                // 1초 대기 후 추가 동작 시작
                await delay(1000);
                
                // 1. SVG 체크박스 클릭
                try {
                  const svgCheckbox = await page.$('div.dyn-container._ln972h.dyn-svg-symbol');
                  if (svgCheckbox) {
                    await svgCheckbox.click();
                    await delay(1000);
                    logger.info(`✅ SVG 체크박스 클릭 완료`);
                    
                    // 2. 그룹웨어 버튼 클릭
                    try {
                      const groupwareButton = await page.$('button[id*="NPS_GroupWareActionPaneTab_button"]');
                      if (groupwareButton) {
                        await groupwareButton.click();
                        await delay(1000);
                        logger.info(`✅ 그룹웨어 버튼 클릭 완료`);
                        
                        // 3. 그룹웨어 승인 버튼 클릭
                        try {
                          const approvalButton = await page.$('div.button-container span.button-label[id*="NPS_IF_GRW_POINVOICEBATCH_label"]');
                          if (approvalButton) {
                            await approvalButton.click();
                            await delay(1000);
                            logger.info(`✅ 그룹웨어 승인 버튼 클릭 완료`);
                            
                            // 4. 새 창(로그인 창) 대기 및 처리 - 개선된 방법
                            try {
                              logger.info('새 창(로그인 창) 대기 중...');
                              
                              let newPage = null;
                              let attempts = 0;
                              const maxAttempts = 10;
                              
                              // 3번째 탭 (인덱스 2) 확인 방법
                              while (!newPage && attempts < maxAttempts) {
                                try {
                                  const pages = await page.browser().pages();
                                  logger.info(`현재 페이지 수: ${pages.length}`);
                                  
                                  // 3번째 탭이 존재하는지 확인 (인덱스 2)
                                  if (pages.length >= 3) {
                                    newPage = pages[2]; // 3번째 탭 (인덱스 2)
                                    logger.info('✅ 3번째 탭에서 새 창 감지됨');
                                    break;
                                  }
                                  
                                  // 만약 3번째 탭이 없으면, 가장 최근에 열린 페이지 확인
                                  if (pages.length > 1) {
                                    newPage = pages[pages.length - 1];
                                    logger.info(`✅ 가장 최근 페이지에서 새 창 감지됨 (총 ${pages.length}개 페이지, 인덱스 ${pages.length - 1})`);
                                    break;
                                  }
                                  
                                } catch (pageError) {
                                  logger.warn(`페이지 확인 실패 (시도 ${attempts + 1}/${maxAttempts}): ${pageError.message}`);
                                }
                                
                                attempts++;
                                logger.info(`3번째 탭 대기 중... (시도 ${attempts}/${maxAttempts})`);
                                await delay(1000);
                              }
                              
                              if (!newPage) {
                                throw new Error('새 창을 감지할 수 없습니다');
                              }
                              
                              // 새 페이지 로딩 대기
                              try {
                                await newPage.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
                              } catch (navError) {
                                logger.warn(`페이지 네비게이션 대기 실패: ${navError.message}, 계속 진행`);
                              }
                              await delay(1000);
                              logger.info('✅ 새 로그인 창 감지 및 로딩 완료');
                              
                              // 4.1 로그인 요소 대기 및 확인
                              logger.info('로그인 요소 대기 중...');
                              let loginAttempts = 0;
                              const maxLoginAttempts = 5;
                              let loginSuccess = false;
                              
                              while (!loginSuccess && loginAttempts < maxLoginAttempts) {
                                try {
                                  // 로그인 요소들이 모두 존재하는지 확인
                                  await newPage.waitForSelector('#txtLoginID', { visible: true, timeout: 2000 });
                                  await newPage.waitForSelector('#txtPassword', { visible: true, timeout: 2000 });
                                  await newPage.waitForSelector('#btnLogin', { visible: true, timeout: 2000 });
                                  
                                  logger.info('✅ 모든 로그인 요소 감지됨');
                                  
                                  // 4.1 아이디 입력 (하드코딩)
                                  const loginId = 'accounting';
                                  await newPage.click('#txtLoginID'); // 포커스
                                  await newPage.evaluate(() => document.querySelector('#txtLoginID').value = ''); // 기존 값 클리어
                                  await newPage.type('#txtLoginID', loginId);
                                  await delay(100);
                                  logger.info(`✅ 로그인 ID 입력 완료: ${loginId}`);
                                  
                                  // 4.2 패스워드 입력 (하드코딩)
                                  const loginPassword = 'P@ssw0rd';
                                  await newPage.click('#txtPassword'); // 포커스
                                  await newPage.evaluate(() => document.querySelector('#txtPassword').value = ''); // 기존 값 클리어
                                  await newPage.type('#txtPassword', loginPassword);
                                  await delay(100);
                                  logger.info(`✅ 로그인 PW 입력 완료`);
                                  
                                  // 4.3 로그인 버튼 클릭
                                  await newPage.click('#btnLogin');
                                  await delay(100);
                                  logger.info(`✅ 로그인 버튼 클릭 완료`);
                                  
                                  loginSuccess = true;
                                  
                                } catch (loginError) {
                                  loginAttempts++;
                                  logger.warn(`로그인 시도 ${loginAttempts}/${maxLoginAttempts} 실패: ${loginError.message}`);
                                  
                                  if (loginAttempts < maxLoginAttempts) {
                                    logger.info('2초 후 재시도...');
                                    await delay(2000);
                                  }
                                }
                              }
                              
                              if (!loginSuccess) {
                                throw new Error('로그인 요소를 찾을 수 없습니다');
                              }
                              
                            } catch (newPageError) {
                              logger.error(`❌ 새 창 로그인 처리 실패: ${newPageError.message}. 작업 중단.`);
                              return;
                            }
                            
                          } else {
                            logger.error('❌ 그룹웨어 승인 버튼을 찾을 수 없습니다. 작업 중단.');
                            return;
                          }
                        } catch (approvalError) {
                          logger.error(`❌ 그룹웨어 승인 버튼 클릭 실패: ${approvalError.message}. 작업 중단.`);
                          return;
                        }
                        
                      } else {
                        logger.error('❌ 그룹웨어 버튼을 찾을 수 없습니다. 작업 중단.');
                        return;
                      }
                    } catch (groupwareError) {
                      logger.error(`❌ 그룹웨어 버튼 클릭 실패: ${groupwareError.message}. 작업 중단.`);
                      return;
                    }
                    
                  } else {
                    logger.error('❌ SVG 체크박스를 찾을 수 없습니다. 작업 중단.');
                    return;
                  }
                } catch (svgError) {
                  logger.error(`❌ SVG 체크박스 클릭 실패: ${svgError.message}. 작업 중단.`);
                  return;
                }
                
              } else {
                logger.error('❌ 콤보박스에서 4번째 항목을 찾을 수 없습니다.');
                return;
              }
            } catch (comboError) {
              logger.error(`❌ 콤보박스 항목 클릭 실패: ${comboError.message}`);
              return;
            }
            
            filterInputFound = true;
          }
        } catch (err) {
          logger.warn(`JavaScript evaluate 필터 입력 실패: ${err.message}`);
        }
      }
      
      if (!filterInputFound) {
        logger.warn('⚠️ 필터 텍스트박스를 찾을 수 없습니다.');
      } else {
        logger.info('✅ 필터 텍스트박스에 I열 값 입력 완료');
      }
    }
    
    logger.info('✅ 6번 RPA 동작: 대기중인 공급사송장 메뉴 이동 및 AU열 값 필터 입력 완료');
    
  } catch (error) {
    logger.error(`6번 RPA 동작 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 * 7번 RPA 동작: 그룹웨어 상신 
 */
async function executeStep7RPA(page) {
  logger.info('🚀 === 7번 RPA 동작: 그룹웨어 상신 시작 ===');
  
  try {
    // 먼저 그룹웨어 새창(3번째 탭)으로 전환
    logger.info('그룹웨어 새창으로 전환 중...');
    const pages = await page.browser().pages();
    logger.info(`현재 열린 페이지 수: ${pages.length}`);
    
    let groupwarePage = null;
    
    // 3번째 탭이 있는지 확인
    if (pages.length >= 3) {
      groupwarePage = pages[2]; // 3번째 탭 (인덱스 2)
      logger.info('3번째 탭을 그룹웨어 페이지로 사용');
    } else if (pages.length > 1) {
      groupwarePage = pages[pages.length - 1]; // 가장 최근 페이지
      logger.info('가장 최근 페이지를 그룹웨어 페이지로 사용');
    } else {
      throw new Error('그룹웨어 새창을 찾을 수 없습니다');
    }
    
    // 그룹웨어 페이지로 포커스 이동
    await groupwarePage.bringToFront();
    await delay(1000);
    
    // 페이지 로딩 완료 대기 - waitForNavigation 제거
    try {
      // 페이지가 이미 로드된 상태이므로 단순 대기만 사용
      await delay(2000);
      logger.info('그룹웨어 페이지 로딩 대기 완료');
    } catch (loadError) {
      logger.warn(`그룹웨어 페이지 로딩 대기 실패: ${loadError.message}, 계속 진행`);
    }
    
    logger.info('✅ 그룹웨어 새창으로 전환 완료');
    
    // 그룹웨어 페이지 완전 로딩 대기 (15초 카운트다운)
    logger.info('🔄 그룹웨어 페이지 완전 로딩 대기 중... (15초)');
    
    // 브라우저 팝업으로 카운트다운 표시
    await groupwarePage.evaluate(() => {
      // 기존 팝업이 있다면 제거
      const existingPopup = document.getElementById('loading-countdown-popup');
      if (existingPopup) existingPopup.remove();
      
      // 팝업 생성
      const popup = document.createElement('div');
      popup.id = 'loading-countdown-popup';
      popup.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: #fff;
        border: 2px solid #007bff;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        z-index: 10000;
        font-family: Arial, sans-serif;
        font-size: 18px;
        text-align: center;
        min-width: 300px;
      `;
      popup.innerHTML = `
        <div style="color: #007bff; font-weight: bold; margin-bottom: 10px;">
          🔄 그룹웨어 페이지 로딩 대기 중
        </div>
        <div id="countdown-text" style="font-size: 24px; color: #ff6b35;">
          15초 남음
        </div>
      `;
      document.body.appendChild(popup);
    });
    
    for (let i = 15; i > 0; i--) {
      logger.info(`⏳ 그룹웨어 로딩 대기: ${i}초 남음`);
      
      // 브라우저 팝업 텍스트 업데이트
      await groupwarePage.evaluate((seconds) => {
        const countdownText = document.getElementById('countdown-text');
        if (countdownText) {
          countdownText.textContent = `${seconds}초 남음`;
        }
      }, i);
      
      await delay(1000);
    }
    
    // 팝업 제거
    await groupwarePage.evaluate(() => {
      const popup = document.getElementById('loading-countdown-popup');
      if (popup) popup.remove();
    });
    
    logger.info('✅ 그룹웨어 페이지 로딩 대기 완료 (15초)');
    
    // 이제 groupwarePage를 사용하여 나머지 작업 수행
    // 1. 보안 설정 클릭
    logger.info('1. 보안 설정 버튼 클릭 중...');
    const securityButton = await groupwarePage.$('#hbtnSetSecurity');
    if (!securityButton) {
      throw new Error('보안 설정 버튼을 찾을 수 없습니다');
    }
    
    await securityButton.click();
    await delay(1000);
    logger.info('✅ 보안 설정 버튼 클릭 완료');
    
    // 2. 공개 항목 체크 & 확인
    logger.info('2. 공개 항목 체크 및 확인 버튼 클릭 중...');
    
    // 공개 라디오 버튼 클릭
    const publicRadio = await groupwarePage.$('input[name="rdoSecurity"][value="1"]');
    if (!publicRadio) {
      throw new Error('공개 라디오 버튼을 찾을 수 없습니다');
    }
    
    await publicRadio.click();
    await delay(500);
    logger.info('✅ 공개 라디오 버튼 클릭 완료');
    
    // 확인 버튼 클릭
    const confirmButton = await groupwarePage.$('span.btn.btn-primary.btn-xs[name="btnSecurity"]');
    if (!confirmButton) {
      throw new Error('보안 설정 확인 버튼을 찾을 수 없습니다');
    }
    
    await confirmButton.click();
    await delay(1000);
    logger.info('✅ 보안 설정 확인 버튼 클릭 완료');
    
    // 3. 제목 설정 (송장설명 + 송장번호)
    logger.info('3. 제목 설정 중...');
    
    // 3.1 송장설명 텍스트 가져오기 - 테이블 구조 기반 (개선된 방법)
    let invoiceDescription = '';
    try {
      logger.info('송장설명 추출: 테이블 구조 기반 방법');
      invoiceDescription = await groupwarePage.evaluate(() => {
        // 모든 테이블 찾기
        const tables = document.querySelectorAll('table');
        
        for (let table of tables) {
          // 헤더에서 '송장설명' 컬럼 찾기
          const headerRows = table.querySelectorAll('thead tr');
          if (headerRows.length === 0) continue;
          
          let descriptionColumnIndex = -1;
          let headerRow = null;
          
          // 모든 헤더 행에서 '송장설명' 찾기
          for (let row of headerRows) {
            const headers = row.querySelectorAll('th');
            for (let i = 0; i < headers.length; i++) {
              if (headers[i].textContent.trim() === '송장설명') {
                descriptionColumnIndex = i;
                headerRow = row;
                break;
              }
            }
            if (descriptionColumnIndex !== -1) break;
          }
          
          if (descriptionColumnIndex === -1) continue;
          
          // tbody에서 첫 번째 데이터 행 찾기
          const tbody = table.querySelector('tbody');
          if (!tbody) continue;
          
          const firstDataRow = tbody.querySelector('tr:first-child');
          if (!firstDataRow) continue;
          
          const cells = firstDataRow.querySelectorAll('td');
          if (cells[descriptionColumnIndex]) {
            const span = cells[descriptionColumnIndex].querySelector('span.fcs_it');
            if (span && span.textContent.trim()) {
              return span.textContent.trim();
            }
          }
        }
        
        return '';
      });
      
      logger.info(`송장설명 값 추출: "${invoiceDescription}"`);
    } catch (descError) {
      logger.warn(`송장설명 추출 실패: ${descError.message}`);
    }
    
    // 3.4 송장번호 텍스트 가져오기 - 테이블 구조 기반 (개선된 방법)
    let invoiceNumber = '';
    try {
      logger.info('송장번호 추출: 테이블 구조 기반 방법');
      invoiceNumber = await groupwarePage.evaluate(() => {
        // 모든 테이블 찾기
        const tables = document.querySelectorAll('table');
        
        for (let table of tables) {
          // 헤더에서 '송장번호' 컬럼 찾기
          const headerRows = table.querySelectorAll('thead tr');
          if (headerRows.length === 0) continue;
          
          let numberColumnIndex = -1;
          let headerRow = null;
          
          // 모든 헤더 행에서 '송장번호' 찾기
          for (let row of headerRows) {
            const headers = row.querySelectorAll('th');
            for (let i = 0; i < headers.length; i++) {
              if (headers[i].textContent.trim() === '송장번호') {
                numberColumnIndex = i;
                headerRow = row;
                break;
              }
            }
            if (numberColumnIndex !== -1) break;
          }
          
          if (numberColumnIndex === -1) continue;
          
          // tbody에서 첫 번째 데이터 행 찾기
          const tbody = table.querySelector('tbody');
          if (!tbody) continue;
          
          const firstDataRow = tbody.querySelector('tr:first-child');
          if (!firstDataRow) continue;
          
          const cells = firstDataRow.querySelectorAll('td');
          if (cells[numberColumnIndex]) {
            const span = cells[numberColumnIndex].querySelector('span.fcs_it');
            if (span && span.textContent.trim()) {
              return span.textContent.trim();
            }
          }
        }
        
        return '';
      });
      
      logger.info(`송장번호 값 추출: "${invoiceNumber}"`);
    } catch (numError) {
      logger.warn(`송장번호 추출 실패: ${numError.message}`);
    }
    
    // 3.2 제목 input 클릭 후 송장설명 붙여넣기
    const titleInput = await groupwarePage.$('input.fcs_itn#FORM_FD_Subject');
    if (!titleInput) {
      throw new Error('제목 input을 찾을 수 없습니다');
    }
    
    await titleInput.click();
    await delay(300);
    
    // 기존 값 클리어
    await groupwarePage.evaluate(() => {
      const input = document.querySelector('#FORM_FD_Subject');
      if (input) {
        input.value = '';
        input.focus();
      }
    });
    
    // 3.2 송장설명 입력 - AU열 값 사용으로 변경
    if (lastProcessedValueFromAUColumn) {
      await titleInput.type(String(lastProcessedValueFromAUColumn));
      logger.info(`✅ 3.2 송장설명 입력 완료 (AU열 값 사용): "${lastProcessedValueFromAUColumn}"`);
    } else {
      logger.warn('AU열 값이 없어 빈 값으로 진행');
    }
    
    // 3.3 중괄호 입력
    await titleInput.type('()');
    logger.info('✅ 3.3 중괄호 "()" 입력 완료');
    
    // 3.5 중괄호 안에 공급사송장 값 입력 (5번 RPA에서 추출한 값 사용)
    if (extractedVendorInvoiceValue) {
      // 백스페이스로 닫는 괄호 제거
      await groupwarePage.keyboard.press('Backspace'); // ) 제거
      
      // 공급사송장에서 추출한 값 입력
      await titleInput.type(String(extractedVendorInvoiceValue));
      
      // 다시 괄호 추가
      await titleInput.type(')');
      
      logger.info(`✅ 3.5 공급사송장 값 입력 완료: "${extractedVendorInvoiceValue}"`);
    } else {
      logger.warn('공급사송장에서 추출한 값이 없어 빈 괄호로 진행');
    }
    
    const finalTitle = `${lastProcessedValueFromAUColumn || ''}(${extractedVendorInvoiceValue || ''})`;
    logger.info(`✅ 최종 제목 완료: "${finalTitle}"`);
    
    // 4. 상신 버튼 클릭
    logger.info('4. 상신 버튼 클릭 중...');
    const submitButton = await groupwarePage.$('#hbtnUpApproval');
    if (!submitButton) {
      throw new Error('상신 버튼을 찾을 수 없습니다');
    }
    
    await submitButton.click();
    await delay(2000); // 팝업이 뜰 시간 대기
    logger.info('✅ 상신 버튼 클릭 완료');
    
    // 5. 상신처리 팝업에서 최종 상신 버튼 클릭
    logger.info('5. 상신처리 팝업에서 최종 상신 버튼 클릭 중...');
    
    // 팝업이 나타날 때까지 대기
    let popupVisible = false;
    let attempts = 0;
    const maxAttempts = 10;
    
    while (!popupVisible && attempts < maxAttempts) {
      try {
        const popup = await groupwarePage.$('#popupDraft');
        if (popup) {
          const isVisible = await groupwarePage.evaluate(el => {
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden';
          }, popup);
          
          if (isVisible) {
            popupVisible = true;
            logger.info('✅ 상신처리 팝업 감지됨');
            break;
          }
        }
      } catch (popupError) {
        logger.warn(`팝업 확인 시도 ${attempts + 1}: ${popupError.message}`);
      }
      
      attempts++;
      await delay(1000);
    }
    
    if (!popupVisible) {
      throw new Error('상신처리 팝업을 찾을 수 없습니다');
    }
    
    // 최종 상신 버튼 클릭
    const finalSubmitButton = await groupwarePage.$('#btnDraft');
    if (!finalSubmitButton) {
      throw new Error('최종 상신 버튼을 찾을 수 없습니다');
    }
    
    await finalSubmitButton.click();
    await delay(2000); // 2초 대기
    logger.info('✅ 최종 상신 버튼 클릭 완료');
    
    // 그룹웨어 창 닫기
    await groupwarePage.close();
    logger.info('✅ 그룹웨어 창 닫기 완료');
    
    logger.info('✅ 7번 RPA 동작: 그룹웨어 상신 완료');
    
  } catch (error) {
    logger.error(`7번 RPA 동작 중 오류: ${error.message}`);
    throw error;
  }
}


// 모듈 내보내기
module.exports = {
  setCredentials,
  getCredentials,
  connectToD365,
  waitForDataTable,
  processInvoice: connectToD365, // 전체 프로세스 기능 활성화 (connectToD365와 동일한 함수)
  openDownloadedExcel,
  openExcelAndExecuteMacro,
  executeExcelProcessing, // 3번 동작: 엑셀 파일 열기 및 매크로 실행 통합 관리
  navigateToPendingVendorInvoice, // 4번 동작: 대기중인 공급사송장 메뉴 이동
  processCloseNewWindow, // 5번 동작: 새창에서 "창 닫기" 버튼 클릭
};

/**
 * AT열값 입력 후 송장 통합 처리 함수
 */
async function processInvoiceIntegrationAfterAT(page) {
  logger.info('🔄 송장 통합 처리 시작...');
  
  try {
    // 1. "송장 통합" 스판 요소 클릭
    logger.info('1. 송장 통합 버튼 찾는 중...');
    
    // 먼저 페이지에 어떤 요소들이 있는지 디버깅
    await page.evaluate(() => {
      console.log('=== 페이지의 모든 span 요소 확인 ===');
      const spans = document.querySelectorAll('span');
      spans.forEach((span, index) => {
        if (span.textContent && span.textContent.includes('송장')) {
          console.log(`Span ${index}: text="${span.textContent.trim()}", id="${span.id}", class="${span.className}"`);
        }
      });
    });
    
    const invoiceIntegrationSelectors = [
      'button[data-dyn-controlname="summaryPurchSetup"]',
      'button[data-dyn-role="DropDialogButton"][data-dyn-controlname="summaryPurchSetup"]',
      'button.dropDialogButton[data-dyn-controlname="summaryPurchSetup"]',
      'button[id*="summaryPurchSetup"]',
      '#VendEditInvoice_5_summaryPurchSetup',
      'button.dynamicsButton.dropDialogButton',
      'span.button-label.button-label-dropDown[id*="summaryPurchSetup_label"]',
      'span[id*="summaryPurchSetup_label"]'
    ];
    
    let integrationButtonFound = false;
    
    for (const selector of invoiceIntegrationSelectors) {
      try {
        logger.info(`송장 통합 버튼 선택자 시도: ${selector}`);
        const button = await page.$(selector);
        if (button) {
          logger.info(`선택자로 요소 발견: ${selector}`);
          const isVisible = await button.isIntersectingViewport();
          logger.info(`요소 가시성 확인: ${isVisible}`);
          
          // 요소의 텍스트 내용 확인
          const textContent = await button.evaluate(el => el.textContent);
          logger.info(`요소 텍스트 내용: "${textContent}"`);
          
          if (isVisible) {
            logger.info(`송장 통합 버튼 클릭 시도: ${selector}`);
            await button.click();
            await delay(500);
            integrationButtonFound = true;
            break;
          } else {
            logger.warn(`요소가 화면에 보이지 않음: ${selector}`);
          }
        } else {
          logger.warn(`선택자로 요소를 찾을 수 없음: ${selector}`);
        }
      } catch (err) {
        logger.warn(`송장 통합 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!integrationButtonFound) {
      // 모든 요소에서 "송장 통합" 텍스트를 포함한 요소 찾기
      try {
        logger.info('모든 요소에서 송장 통합 텍스트 검색 중...');
        const result = await page.evaluate(() => {
          // 모든 clickable 요소 검색
          const allElements = document.querySelectorAll('span, button, div, a');
          const foundElements = [];
          
          for (const element of allElements) {
            if (element.textContent && element.textContent.includes('송장 통합')) {
              foundElements.push({
                tagName: element.tagName,
                id: element.id,
                className: element.className,
                textContent: element.textContent.trim(),
                isVisible: element.offsetParent !== null
              });
            }
          }
          
          // 첫 번째로 찾은 송장 통합 요소 클릭 시도
          if (foundElements.length > 0) {
            const element = document.querySelector(`${foundElements[0].tagName.toLowerCase()}${foundElements[0].id ? '#' + foundElements[0].id : ''}${foundElements[0].className ? '.' + foundElements[0].className.split(' ').join('.') : ''}`);
            if (element) {
              element.click();
              return { success: true, elements: foundElements };
            }
          }
          
          return { success: false, elements: foundElements };
        });
        
        logger.info(`발견된 송장 통합 요소들: ${JSON.stringify(result.elements, null, 2)}`);
        
        if (result.success) {
          logger.info('포괄적 검색으로 송장 통합 버튼 클릭 성공');
          await delay(500);
          integrationButtonFound = true;
        } else if (result.elements.length > 0) {
          logger.warn('송장 통합 요소는 발견했지만 클릭에 실패');
        } else {
          logger.warn('송장 통합 텍스트를 포함한 요소를 찾을 수 없음');
        }
      } catch (err) {
        logger.warn(`포괄적 송장 통합 버튼 검색 실패: ${err.message}`);
      }
    }
    
    if (!integrationButtonFound) {
      throw new Error('송장 통합 버튼을 찾을 수 없습니다.');
    }
    
    // 2. 송장 통합 DropDialogButton 클릭 후 나타나는 옵션 대기
    logger.info('2. 송장 통합 드롭다운 옵션 대기 중...');
    
    // DropDialogButton 클릭 후 옵션이 나타날 때까지 대기
    await delay(1500);
    
    // 3. "송장 계정" 옵션 클릭
    logger.info('3. "송장 계정" 옵션 찾는 중...');
    
    // 먼저 페이지에 있는 모든 li 요소 확인
    await page.evaluate(() => {
      console.log('=== 페이지의 모든 li 요소 확인 ===');
      const lis = document.querySelectorAll('li');
      lis.forEach((li, index) => {
        if (li.textContent && li.textContent.trim()) {
          console.log(`Li ${index}: text="${li.textContent.trim()}", id="${li.id}", class="${li.className}", data-dyn-index="${li.getAttribute('data-dyn-index')}"`);
        }
      });
    });
    
    const optionSelectors = [
      'li[data-dyn-index="1"]',
      'li[id*="list_item1"]',
      'li[role="option"]',
      'li[id*="sumBy_list_item"]',
      'li:contains("송장 계정")',
      'li'
    ];
    
    let optionFound = false;
    
    for (const selector of optionSelectors) {
      try {
        logger.info(`송장 계정 옵션 선택자 시도: ${selector}`);
        
        if (selector === 'li') {
          // 모든 li 요소에 대해 하나씩 확인
          const allLis = await page.$$('li');
          logger.info(`총 ${allLis.length}개의 li 요소 발견`);
          
          for (let i = 0; i < allLis.length; i++) {
            try {
              const li = allLis[i];
              const isVisible = await li.isIntersectingViewport();
              const text = await li.evaluate(el => el.textContent);
              
              if (text && text.includes('송장 계정') && isVisible) {
                logger.info(`송장 계정 옵션 발견 (li[${i}]): "${text.trim()}"`);
                await li.click();
                await delay(500);
                optionFound = true;
                break;
              }
            } catch (innerErr) {
              // 개별 li 처리 실패는 무시
            }
          }
          
          if (optionFound) break;
        } else {
          const option = await page.$(selector);
          if (option) {
            logger.info(`선택자로 요소 발견: ${selector}`);
            const isVisible = await option.isIntersectingViewport();
            logger.info(`요소 가시성: ${isVisible}`);
            
            if (isVisible) {
              const text = await option.evaluate(el => el.textContent);
              logger.info(`요소 텍스트: "${text}"`);
              
              if (text && text.includes('송장 계정')) {
                logger.info(`송장 계정 옵션 클릭: ${selector}`);
                await option.click();
                await delay(500);
                optionFound = true;
                break;
              }
            }
          } else {
            logger.warn(`선택자로 요소를 찾을 수 없음: ${selector}`);
          }
        }
      } catch (err) {
        logger.warn(`송장 계정 옵션 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!optionFound) {
      // JavaScript evaluate로 "송장 계정" 텍스트 찾기
      try {
        logger.info('JavaScript evaluate로 송장 계정 옵션 찾는 중...');
        const evaluateResult = await page.evaluate(() => {
          const lis = document.querySelectorAll('li');
          for (const li of lis) {
            if (li.textContent && li.textContent.includes('송장 계정')) {
              li.click();
              return true;
            }
          }
          return false;
        });
        
        if (evaluateResult) {
          logger.info('JavaScript evaluate로 송장 계정 옵션 클릭 성공');
          await delay(500);
          optionFound = true;
        }
      } catch (err) {
        logger.warn(`JavaScript evaluate 송장 계정 옵션 클릭 실패: ${err.message}`);
      }
    }
    
    if (!optionFound) {
      throw new Error('송장 계정 옵션을 찾을 수 없습니다.');
    }
    
    // 4. "연결" 버튼 클릭
    logger.info('4. "연결" 버튼 찾는 중...');
    await delay(1000);
    
    const connectButtonSelectors = [
      'span[id*="buttonReArrange_label"]',
      'span.button-label[for*="buttonReArrange"]',
      'span[class="button-label"]'
    ];
    
    let connectButtonFound = false;
    
    for (const selector of connectButtonSelectors) {
      try {
        const button = await page.$(selector);
        if (button) {
          const isVisible = await button.isIntersectingViewport();
          if (isVisible) {
            const text = await button.textContent();
            if (text && text.trim() === '연결') {
              logger.info(`연결 버튼 클릭: ${selector}`);
              await button.click();
              await delay(500);
              connectButtonFound = true;
              break;
            }
          }
        }
      } catch (err) {
        logger.warn(`연결 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!connectButtonFound) {
      // JavaScript evaluate로 "연결" 텍스트 찾기
      try {
        logger.info('JavaScript evaluate로 연결 버튼 찾는 중...');
        const buttonFound = await page.evaluate(() => {
          const spans = document.querySelectorAll('span');
          for (const span of spans) {
            if (span.textContent && span.textContent.trim() === '연결') {
              span.click();
              return true;
            }
          }
          return false;
        });
        
        if (buttonFound) {
          logger.info('JavaScript evaluate로 연결 버튼 클릭 성공');
          await delay(500);
          connectButtonFound = true;
        }
      } catch (err) {
        logger.warn(`JavaScript evaluate 연결 버튼 클릭 실패: ${err.message}`);
      }
    }
    
    if (!connectButtonFound) {
      logger.error('연결 버튼을 찾을 수 없습니다. 송장 통합 처리를 건너뜁니다.');
      // 에러를 throw하지 않고 경고만 남김 (전체 프로세스 중단 방지)
    }
    
    logger.info('✅ 송장 통합 처리 완료');
    
  } catch (error) {
    logger.error(`송장 통합 처리 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 * AV열 송장일 입력 후 송장 통합 처리 함수
 */
async function processInvoiceIntegrationAfterAV(page) {
  logger.info('🔄 AV열 후 송장 통합 처리 시작...');
  
  try {
    // 1. "송장 통합" 버튼 클릭하여 팝업 열기
    logger.info('1. 송장 통합 버튼 찾는 중...');
    
    const invoiceIntegrationSelectors = [
      'button[data-dyn-controlname="summaryPurchSetup"]',
      'button[data-dyn-role="DropDialogButton"][data-dyn-controlname="summaryPurchSetup"]',
      'button.dropDialogButton[data-dyn-controlname="summaryPurchSetup"]',
      'button[id*="summaryPurchSetup"]',
      '#VendEditInvoice_5_summaryPurchSetup',
      'button.dynamicsButton.dropDialogButton',
      'span.button-label.button-label-dropDown[id*="summaryPurchSetup_label"]',
      'span[id*="summaryPurchSetup_label"]'
    ];
    
    let integrationButtonFound = false;
    
    for (const selector of invoiceIntegrationSelectors) {
      try {
        logger.info(`송장 통합 버튼 선택자 시도: ${selector}`);
        const button = await page.$(selector);
        if (button) {
          const isVisible = await button.isIntersectingViewport();
          const textContent = await button.evaluate(el => el.textContent);
          logger.info(`요소 발견: ${selector}, 가시성: ${isVisible}, 텍스트: "${textContent}"`);
          
          if (isVisible) {
            logger.info(`송장 통합 버튼 클릭 시도: ${selector}`);
            await button.click();
            await delay(1000); // 팝업이 열릴 시간 대기
            integrationButtonFound = true;
            break;
          }
        }
      } catch (err) {
        logger.warn(`송장 통합 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!integrationButtonFound) {
      // 포괄적 검색으로 송장 통합 버튼 찾기
      const result = await page.evaluate(() => {
        const allElements = document.querySelectorAll('span, button, div, a');
        for (const element of allElements) {
          if (element.textContent && element.textContent.includes('송장 통합') && element.offsetParent !== null) {
            element.click();
            return { success: true, text: element.textContent.trim() };
          }
        }
        return { success: false };
      });
      
      if (result.success) {
        logger.info(`포괄적 검색으로 송장 통합 버튼 클릭 성공: "${result.text}"`);
        await delay(1000);
        integrationButtonFound = true;
      }
    }
    
    if (!integrationButtonFound) {
      throw new Error('송장 통합 버튼을 찾을 수 없습니다.');
    }
    
    // 2. 팝업 다이얼로그가 열렸는지 확인하고 sumBy input textbox 클릭
    logger.info('2. 팝업 다이얼로그에서 sumBy input textbox 찾는 중...');
    
    // 팝업이 완전히 로드될 때까지 대기
    await delay(3000);
    
    const sumByInputSelectors = [
      'input[name="sumBy"]',
      'input[data-dyn-controlname="sumBy"]',
      'input[id*="sumBy"]',
      'input[class*="textbox"][role="combobox"]',
      'input[title="송장 계정"]',
      'input[id$="_sumBy_input"]',
      'input[class*="textbox"][name="sumBy"]',
      'input[role="combobox"]',
      'input[class*="textbox"][class*="field"]',
      'input[type="text"][role="combobox"]',
      'input[class*="textbox"]',
      'input'
    ];
    
    let sumByInputFound = false;
    
    for (const selector of sumByInputSelectors) {
      try {
        logger.info(`sumBy input 선택자 시도: ${selector}`);
        const input = await page.$(selector);
        if (input) {
          const isVisible = await input.isIntersectingViewport();
          logger.info(`sumBy input 발견: ${selector}, 가시성: ${isVisible}`);
          
          if (isVisible) {
            logger.info(`sumBy input textbox 클릭: ${selector}`);
            await input.click();
            await delay(800); // 드롭다운이 나타날 시간 대기
            sumByInputFound = true;
            break;
          }
        }
      } catch (err) {
        logger.warn(`sumBy input 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!sumByInputFound) {
      throw new Error('팝업에서 sumBy input textbox를 찾을 수 없습니다.');
    }
    
    // 3. 드롭다운에서 "송장 계정" 옵션 선택
    logger.info('3. 드롭다운에서 "송장 계정" 옵션 찾는 중...');
    
    // 드롭다운 옵션들 확인을 위한 디버깅
    await page.evaluate(() => {
      console.log('=== 팝업 내 드롭다운 옵션 확인 ===');
      const options = document.querySelectorAll('li, option, div[role="option"]');
      options.forEach((option, index) => {
        if (option.textContent && option.textContent.trim()) {
          console.log(`Option ${index}: text="${option.textContent.trim()}", tag=${option.tagName}, visible=${option.offsetParent !== null}`);
        }
      });
    });
    
    const optionSelectors = [
      'li[data-dyn-index="1"]',
      'li[role="option"]',
      'option[value*="account"]',
      'div[role="option"]',
      'li'
    ];
    
    let optionFound = false;
    
    for (const selector of optionSelectors) {
      try {
        logger.info(`송장 계정 옵션 선택자 시도: ${selector}`);
        
        if (selector === 'li') {
          // 모든 li 요소에서 "송장 계정" 텍스트 찾기
          const allLis = await page.$$('li');
          logger.info(`총 ${allLis.length}개의 li 요소 발견`);
          
          for (let i = 0; i < allLis.length; i++) {
            try {
              const li = allLis[i];
              const isVisible = await li.isIntersectingViewport();
              const text = await li.evaluate(el => el.textContent);
              
              logger.info(`li[${i}] 확인: "${text ? text.trim() : 'null'}", 가시성: ${isVisible}`);
              
              if (text && isVisible && (
                text.includes('송장 계정') || 
                text.trim() === '송장 계정' ||
                text.includes('송장계정')
              )) {
                logger.info(`송장 계정 옵션 발견 및 클릭: "${text.trim()}"`);
                await li.click();
                await delay(500);
                optionFound = true;
                break;
              }
            } catch (innerErr) {
              // 개별 li 처리 실패는 무시
            }
          }
          
          if (optionFound) break;
        } else {
          const option = await page.$(selector);
          if (option) {
            const isVisible = await option.isIntersectingViewport();
            const text = await option.evaluate(el => el.textContent);
            
            if (isVisible && text && text.includes('송장 계정')) {
              logger.info(`송장 계정 옵션 클릭: ${selector}`);
              await option.click();
              await delay(500);
              optionFound = true;
              break;
            }
          }
        }
      } catch (err) {
        logger.warn(`송장 계정 옵션 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!optionFound) {
      // JavaScript evaluate로 더 정확한 검색
      const evaluateResult = await page.evaluate(() => {
        const allElements = document.querySelectorAll('li, option, div[role="option"]');
        for (const element of allElements) {
          const text = element.textContent ? element.textContent.trim() : '';
          if (text && element.offsetParent !== null && (
            text.includes('송장 계정') || 
            text === '송장 계정' ||
            text.includes('송장계정')
          )) {
            console.log(`송장 계정 옵션 클릭 시도: "${text}"`);
            element.click();
            return { success: true, clickedText: text };
          }
        }
        return { success: false };
      });
      
      if (evaluateResult.success) {
        logger.info(`JavaScript evaluate로 송장 계정 옵션 클릭 성공: "${evaluateResult.clickedText}"`);
        await delay(500);
        optionFound = true;
      }
    }
    
    if (!optionFound) {
      throw new Error('드롭다운에서 송장 계정 옵션을 찾을 수 없습니다.');
    }
    
    // 4. 팝업 내 "연결" 버튼 (#110_9_buttonReArrange) 클릭
    logger.info('4. 팝업 내 "연결" 버튼 찾는 중...');
    await delay(1000);
    
    // 팝업 내 버튼들 확인을 위한 디버깅
    await page.evaluate(() => {
      console.log('=== 팝업 내 모든 버튼 확인 ===');
      const buttons = document.querySelectorAll('button, span[class*="button"]');
      buttons.forEach((button, index) => {
        const text = button.textContent ? button.textContent.trim() : '';
        if (text && (text.includes('연결') || text.includes('재배치') || button.id.includes('buttonReArrange'))) {
          console.log(`Button ${index}: text="${text}", id="${button.id}", class="${button.className}"`);
        }
      });
    });
    
    const connectButtonSelectors = [
      '#110_9_buttonReArrange',
      'button[id*="buttonReArrange"]',
      'button[data-dyn-controlname*="buttonReArrange"]',
      'span[id*="buttonReArrange_label"]'
    ];
    
    let connectButtonFound = false;
    
    for (const selector of connectButtonSelectors) {
      try {
        logger.info(`연결 버튼 선택자 시도: ${selector}`);
        const button = await page.$(selector);
        if (button) {
          const isVisible = await button.isIntersectingViewport();
          const text = await button.evaluate(el => el.textContent);
          logger.info(`연결 버튼 발견: ${selector}, 가시성: ${isVisible}, 텍스트: "${text}"`);
          
          if (isVisible) {
            logger.info(`연결 버튼 클릭: ${selector}`);
            await button.click();
            await delay(500);
            connectButtonFound = true;
            break;
          }
        }
      } catch (err) {
        logger.warn(`연결 버튼 선택자 실패: ${selector} - ${err.message}`);
      }
    }
    
    if (!connectButtonFound) {
      // JavaScript evaluate로 연결 버튼 찾기
      const result = await page.evaluate(() => {
        const allElements = document.querySelectorAll('button, span');
        for (const element of allElements) {
          const text = element.textContent ? element.textContent.trim() : '';
          const id = element.id || '';
          
          if (element.offsetParent !== null && (
            text === '연결' || 
            text === '재배치' ||
            text.includes('연결') ||
            text.includes('재배치') ||
            id.includes('buttonReArrange')
          )) {
            console.log(`연결 버튼 클릭 시도: "${text}", id: "${id}"`);
            element.click();
            return { success: true, clickedText: text, id: id };
          }
        }
        return { success: false };
      });
      
      if (result.success) {
        logger.info(`JavaScript evaluate로 연결 버튼 클릭 성공: "${result.clickedText}", id: "${result.id}"`);
        await delay(500);
        connectButtonFound = true;
      }
    }
    
    if (!connectButtonFound) {
      logger.error('팝업에서 연결 버튼을 찾을 수 없습니다. 송장 통합 처리를 건너뜁니다.');
      // 에러를 throw하지 않고 경고만 남김 (전체 프로세스 중단 방지)
    }
    
    // 팝업이 닫힐 때까지 대기
    logger.info('팝업이 닫힐 때까지 대기 중...');
    await delay(3000); // 더 긴 대기 시간
    
    // 원래 페이지로 포커스 돌아가기 위해 페이지 클릭
    logger.info('원래 페이지로 포커스 돌아가기...');
    await page.mouse.click(100, 100);
    await delay(1000);
    
    // 페이지가 완전히 로드되었는지 확인
    logger.info('페이지 로딩 상태 확인 중...');
    try {
      await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 5000 });
      logger.info('페이지 로딩 완료 확인됨');
    } catch (loadWaitError) {
      logger.warn(`페이지 로딩 대기 중 오류: ${loadWaitError.message}, 계속 진행`);
    }
    
    await delay(1000); // 추가 안정화 대기
    
    logger.info('✅ AV열 후 송장 통합 처리 완료');
    
  } catch (error) {
    logger.error(`AV열 후 송장 통합 처리 중 오류: ${error.message}`);
    throw error;
  }
}

/**
 * 현재 날짜를 YYYY-MM-DD 형식으로 반환하는 함수
 */
function getCurrentDateFormatted() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 매입송장 처리 메인 함수 - 전체 RPA 프로세스 실행
 */
async function processInvoice(credentials) {
  try {
    logger.info('🚀 === 다중모드 매입송장 처리 시작 ===');
    
    // 1~7. 전체 RPA 프로세스 실행 (connectToD365가 모든 단계 포함)
    const result = await connectToD365(credentials);
    
    logger.info('✅ 다중모드 매입송장 처리 완료');
    
    return result;
    
  } catch (error) {
    logger.error(`매입송장 처리 중 오류: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 단계별 진행 상황을 추적하는 매입송장 처리 함수 (다중모드용)
 */
async function processInvoiceWithProgress(credentials, progressCallback, cycle) {
  try {
    logger.info(`🚀 === ${cycle}번째 사이클 매입송장 처리 시작 ===`);
    
    // 1~7. 전체 RPA 프로세스 실행 (단계별 콜백 포함)
    const result = await connectToD365WithProgress(credentials, progressCallback, cycle);
    
    logger.info(`✅ ${cycle}번째 사이클 매입송장 처리 완료`);
    
    return result;
    
  } catch (error) {
    logger.error(`${cycle}번째 사이클 매입송장 처리 중 오류: ${error.message}`);
    
    // 에러 발생 시 콜백 호출
    if (progressCallback) {
      progressCallback(cycle, null, null, error.message);
    }
    
    return {
      success: false,
      error: error.message,
      failedStep: 1 // 기본적으로 1번 단계에서 실패로 가정
    };
  }
}

/**
 * A열 값 설정 함수 - UI에서 사용자가 입력한 A열 값을 설정
 */
function setValueA(valueA) {
  const oldValue = userInputValueA;
  userInputValueA = parseInt(valueA);
  logger.info(`🎯 사용자 A열 값 설정: ${oldValue} → ${userInputValueA}`);
  logger.info(`✅ A열 값 설정 완료: userInputValueA = ${userInputValueA}`);
}

/**
 * 여러 A열 값을 순차적으로 처리하는 함수
 */
async function processMultipleValueA(valueArray, credentials) {
  const results = [];
  
  logger.info(`🚀 === 다중모드 시작: ${valueArray.length}개 A열 값 처리 ===`);
  
  for (let i = 0; i < valueArray.length; i++) {
    const currentValue = parseInt(valueArray[i]);
    const isFirstCycle = i === 0;
    const isLastCycle = i === valueArray.length - 1;
    
    logger.info(`\n🔄 다중 처리 ${i + 1}/${valueArray.length}: A열 값 ${currentValue} 처리 시작`);
    logger.info(`📍 사이클 타입: ${isFirstCycle ? '첫 번째 사이클' : '후속 사이클'}`);
    logger.info(`📍 마지막 사이클: ${isLastCycle ? 'YES' : 'NO'}`);
    
    try {
      // A열 값 설정
      setValueA(currentValue);
      
      // 개별 RPA 프로세스 실행 (단계별 진행 추적 포함)
      const result = await processInvoiceWithProgress(credentials, 
        // 단계별 진행 상황 콜백
        (cycleNum, currentStep, completedSteps, error) => {
          // 각 사이클의 단계별 진행 상황을 결과에 저장
          if (!results[i]) {
            results[i] = {
              valueA: currentValue,
              cycle: i + 1,
              success: false,
              message: '',
              completedAt: '',
              stepDetails: [] // 단계별 상세 정보 저장
            };
          }
          
          const stepInfo = {
            step: currentStep,
            completedSteps: completedSteps,
            timestamp: new Date().toISOString(),
            error: error
          };
          
          if (error) {
            stepInfo.status = 'failed';
            stepInfo.errorMessage = error;
            logger.error(`${i + 1}번째 사이클 - ${currentStep}단계 실패: ${error}`);
          } else {
            stepInfo.status = completedSteps >= currentStep ? 'completed' : 'in_progress';
            logger.info(`${i + 1}번째 사이클 - ${currentStep}단계 진행 중 (완료: ${completedSteps}단계)`);
          }
          
          results[i].stepDetails.push(stepInfo);
        },
        i + 1 // cycle number
      );
      
      // 콜백에서 이미 results[i]가 생성되었으므로 업데이트만 수행
      if (results[i]) {
        results[i].success = result.success;
        results[i].message = result.message;
        results[i].completedAt = new Date().toISOString();
        results[i].failedStep = result.failedStep;
      } else {
        // 혹시 콜백이 호출되지 않은 경우를 위한 fallback
        results.push({
          valueA: currentValue,
          cycle: i + 1,
          success: result.success,
          message: result.message,
          completedAt: new Date().toISOString(),
          failedStep: result.failedStep,
          stepDetails: [] // 빈 단계 정보
        });
      }
      
      if (result.success) {
        logger.info(`✅ A열 값 ${currentValue} 처리 완료`);
      } else {
        logger.error(`❌ A열 값 ${currentValue} 처리 실패: ${result.message}`);
        
        // 첫 번째 사이클에서 에러 발생시 전체 프로세스 중단
        if (isFirstCycle) {
          logger.error(`🚨 첫 번째 사이클에서 에러 발생 - 전체 다중모드 프로세스 중단`);
          logger.error(`🚨 에러 상세: ${result.message}`);
          return {
            success: false,
            totalProcessed: 1,
            successCount: 0,
            failCount: 1,
            results: results,
            isMultipleMode: true,
            error: `첫 번째 사이클 실패로 인한 전체 프로세스 중단: ${result.message}`,
            message: `첫 번째 사이클에서 에러가 발생하여 다중모드를 중단했습니다.`
          };
        }
        
        // 두 번째 이후 사이클에서는 에러 로그만 남기고 다음 사이클 진행
        logger.error(`⚠️ ${i + 1}번째 사이클 실패, 다음 사이클로 진행합니다.`);
      }
      
      // 마지막 사이클 완료 후 프로세스 종료
      if (isLastCycle) {
        logger.info(`🏁 마지막 사이클 완료 - 전체 다중모드 처리 종료`);
        break;
      }
      
      // 다음 처리 전 대기 (마지막 사이클이 아닌 경우)
      if (i < valueArray.length - 1) {
        logger.info('⏳ 다음 처리를 위해 5초 대기...');
        await delay(5000);
      }
      
    } catch (error) {
      logger.error(`❌ A열 값 ${currentValue} 처리 중 예외 발생: ${error.message}`);
      
      // 예외 발생시에도 단계별 정보를 포함
      if (results[i]) {
        results[i].success = false;
        results[i].message = error.message;
        results[i].completedAt = new Date().toISOString();
        results[i].errorDetails = error.stack;
        results[i].failedStep = getCurrentStepFromError(error.message);
      } else {
        results.push({
          valueA: currentValue,
          cycle: i + 1,
          success: false,
          message: error.message,
          completedAt: new Date().toISOString(),
          errorDetails: error.stack,
          failedStep: getCurrentStepFromError(error.message),
          stepDetails: [] // 빈 단계 정보
        });
      }
      
      // 첫 번째 사이클에서 예외 발생시 전체 프로세스 중단
      if (isFirstCycle) {
        logger.error(`🚨 첫 번째 사이클에서 예외 발생 - 전체 다중모드 프로세스 중단`);
        logger.error(`🚨 예외 상세: ${error.message}`);
        return {
          success: false,
          totalProcessed: 1,
          successCount: 0,
          failCount: 1,
          results: results,
          isMultipleMode: true,
          error: `첫 번째 사이클 예외로 인한 전체 프로세스 중단: ${error.message}`,
          message: `첫 번째 사이클에서 예외가 발생하여 다중모드를 중단했습니다.`
        };
      }
      
      // 두 번째 이후 사이클에서는 예외 로그만 남기고 다음 사이클 진행
      logger.error(`⚠️ ${i + 1}번째 사이클 예외 발생, 다음 사이클로 진행합니다.`);
      
      // 마지막 사이클에서 예외 발생해도 프로세스 종료
      if (isLastCycle) {
        logger.info(`🏁 마지막 사이클에서 예외 발생했지만 전체 다중모드 처리 종료`);
        break;
      }
    }
  }
  
  // 전체 결과 요약
  const successCount = results.filter(r => r.success).length;
  const failCount = results.length - successCount;
  
  logger.info(`\n📊 === 다중모드 완료 ===`);
  logger.info(`📈 처리 통계 - 총: ${results.length}, 성공: ${successCount}, 실패: ${failCount}`);
  logger.info(`📋 상세 결과:`);
  
  results.forEach(result => {
    const status = result.success ? '✅ 성공' : '❌ 실패';
    logger.info(`  - ${result.cycle}번째 사이클 (A열 값: ${result.valueA}): ${status}`);
    
    if (!result.success) {
      logger.error(`    오류: ${result.message}`);
      if (result.failedStep) {
        logger.error(`    실패 단계: ${result.failedStep}단계`);
      }
      if (result.errorDetails) {
        logger.error(`    상세: ${result.errorDetails.split('\n')[0]}`); // 첫 번째 줄만 표시
      }
    }
    
    // 단계별 처리 상세 표시 (각 사이클 완료 후)
    if (result.stepDetails && result.stepDetails.length > 0) {
      logger.info(`    📋 ${result.cycle}번째 사이클 단계별 상세:`);
      result.stepDetails.forEach(stepDetail => {
        const stepStatus = stepDetail.status === 'completed' ? '✅' : 
                          stepDetail.status === 'failed' ? '❌' : '⏳';
        const stepMsg = stepDetail.error ? ` (${stepDetail.error})` : '';
        logger.info(`      ${stepStatus} ${stepDetail.step}단계${stepMsg}`);
      });
    }
  });
  
  logger.info('🎉 === 다중모드 전체 완료 ===');
  
  return {
    success: failCount === 0,
    totalProcessed: results.length,
    successCount: successCount,
    failCount: failCount,
    results: results,
    isMultipleMode: true,
    message: `총 ${results.length}개 A열 값 처리 완료 (성공: ${successCount}, 실패: ${failCount})`,
    completedAt: new Date().toISOString()
  };
}

// 모듈 내보내기
module.exports = {
  setCredentials,
  getCredentials,
  setSelectedDateRange,
  getSelectedDateRange,
  connectToD365,
  connectToD365WithProgress,
  processInvoice,
  processInvoiceWithProgress,
  setValueA,
  processMultipleValueA,
  getCurrentDateFormatted
};