/**
 * ERP RPA 전역 로그인 모듈
 * 두 페이지(전표 처리, 매입송장 처리)에서 공통으로 사용하는 로그인 시스템
 */

// 전역 로그인 상태 관리
let globalLoginState = {
  isLoggedIn: false,
  username: '',
  password: '',
  loginTime: null
};

// 로그인 오버레이 HTML 템플릿
function getLoginOverlayHTML() {
  return `
    <!-- 전역 로그인 오버레이 -->
    <div class="login-overlay" id="loginOverlay">
      <div class="login-container">
        <img src="ERP_RPA아이콘.png" alt="ERP RPA" class="login-logo">
        <div class="login-title">ERP RPA 시스템</div>
        <div class="login-subtitle">로그인이 필요합니다</div>
        
        <form class="login-form" onsubmit="globalLogin(event)">
          <div class="login-form-group">
            <label for="globalUserId">사용자 ID</label>
            <input type="text" id="globalUserId" placeholder="ERP 사용자 ID를 입력하세요" required>
          </div>
          <div class="login-form-group">
            <label for="globalUserPw">비밀번호</label>
            <input type="password" id="globalUserPw" placeholder="비밀번호를 입력하세요" required>
          </div>
          <div class="login-options">
            <div class="remember-me">
              <input type="checkbox" id="rememberLogin">
              <label for="rememberLogin">로그인 상태 유지</label>
            </div>
          </div>
          <button type="submit" class="login-btn" id="globalLoginBtn">로그인</button>
        </form>
      </div>
    </div>

    <!-- 사용자 상태 표시 -->
    <div class="user-status hidden" id="userStatus">
      <span>👤</span>
      <span id="currentUser">사용자</span>
      <button class="logout-btn" onclick="globalLogout()">로그아웃</button>
    </div>
  `;
}

// 로그인 오버레이 CSS 스타일
function getLoginOverlayCSS() {
  return `
    /* 전역 로그인 오버레이 스타일 */
    .login-overlay {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.8);
      backdrop-filter: blur(10px);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all 0.3s ease;
    }    .login-overlay.hidden {
      display: none !important;
    }

    .login-container {
      background: white;
      border-radius: 20px;
      padding: 40px;
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
      max-width: 400px;
      width: 90%;
      text-align: center;
      transform: scale(1);
      transition: transform 0.3s ease;
    }

    .login-container:hover {
      transform: scale(1.02);
    }

    .login-logo {
      width: 80px;
      height: 80px;
      margin: 0 auto 20px;
      border-radius: 16px;
      object-fit: cover;
    }

    .login-title {
      font-size: 24px;
      font-weight: bold;
      color: #2c3e50;
      margin-bottom: 10px;
    }

    .login-subtitle {
      color: #666;
      margin-bottom: 30px;
      font-size: 14px;
    }

    .login-form {
      text-align: left;
    }

    .login-form-group {
      margin-bottom: 20px;
    }

    .login-form-group label {
      display: block;
      margin-bottom: 8px;
      font-weight: 500;
      color: #333;
    }

    .login-form-group input {
      width: 100%;
      padding: 12px 16px;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      font-size: 14px;
      transition: border-color 0.3s ease;
      box-sizing: border-box;
    }

    .login-form-group input:focus {
      border-color: #2980b9;
      outline: none;
    }

    .login-options {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 25px;
      font-size: 14px;
    }

    .remember-me {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .remember-me input[type="checkbox"] {
      width: auto;
    }

    .login-btn {
      width: 100%;
      padding: 14px;
      background: linear-gradient(135deg, #2980b9, #3498db);
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s ease;
      margin-bottom: 15px;
    }

    .login-btn:hover {
      background: linear-gradient(135deg, #1c6692, #2980b9);
      transform: translateY(-2px);
    }

    .login-btn:disabled {
      background: #bdc3c7;
      cursor: not-allowed;
      transform: none;
    }

    .user-status {
      position: fixed;
      top: 80px;
      right: 20px;
      background: rgba(255, 255, 255, 0.95);
      padding: 10px 15px;
      border-radius: 25px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      font-size: 14px;
      font-weight: 500;
      color: #2c3e50;
      z-index: 1000;
      display: flex;
      align-items: center;
      gap: 8px;
      transition: all 0.3s ease;
    }    .user-status.hidden {
      display: none !important;
    }

    .logout-btn {
      background: #e74c3c;
      color: white;
      border: none;
      border-radius: 15px;
      padding: 4px 12px;
      font-size: 12px;
      cursor: pointer;
      transition: background 0.3s ease;
    }

    .logout-btn:hover {
      background: #c0392b;
    }

    /* 다크모드 로그인 스타일 */
    body.dark-mode .login-container {
      background: #2d2d2d;
      color: #e0e0e0;
    }

    body.dark-mode .login-title {
      color: #e0e0e0;
    }

    body.dark-mode .login-subtitle {
      color: #b0b0b0;
    }

    body.dark-mode .login-form-group label {
      color: #e0e0e0;
    }

    body.dark-mode .login-form-group input {
      background: #404040;
      border-color: #555;
      color: #e0e0e0;
    }

    body.dark-mode .login-form-group input:focus {
      border-color: #4a9eff;
    }

    body.dark-mode .user-status {
      background: rgba(45, 45, 45, 0.95);
      color: #e0e0e0;
    }

    .main-content.blurred {
      filter: blur(3px);
      pointer-events: none;
    }
  `;
}

// 저장된 로그인 정보 확인
function checkSavedLogin() {
  const savedLogin = localStorage.getItem('erpLogin');
  const rememberLogin = localStorage.getItem('rememberLogin') === 'true';
  
  if (savedLogin && rememberLogin) {
    try {
      const loginData = JSON.parse(savedLogin);
      globalLoginState = {
        isLoggedIn: true,
        username: loginData.username,
        password: loginData.password,
        loginTime: loginData.loginTime
      };
      
      // 로그인 정보를 Electron API에도 설정
      if (window.electronAPI) {
        window.electronAPI.setCredentials({
          username: loginData.username,
          password: loginData.password
        });
      }
    } catch (error) {
      console.error('저장된 로그인 정보 복원 실패:', error);
      localStorage.removeItem('erpLogin');
    }
  }
}

// 전역 로그인 함수
function globalLogin(event) {
  event.preventDefault();
  
  const userId = document.getElementById('globalUserId').value.trim();
  const userPw = document.getElementById('globalUserPw').value;
  const rememberLogin = document.getElementById('rememberLogin').checked;
  
  if (!userId || !userPw) {
    alert('ID와 비밀번호를 모두 입력해주세요.');
    return;
  }
  
  // 사용자 ID에 '@nepes.co.kr' 추가 (이미 포함되어 있지 않은 경우)
  const username = userId.includes('@') ? userId : `${userId}@nepes.co.kr`;
  
  // 로딩 상태 표시
  const loginBtn = document.getElementById('globalLoginBtn');
  loginBtn.disabled = true;
  loginBtn.textContent = '로그인 중...';
  
  // contextBridge를 통해 노출된 API 사용
  if (window.electronAPI) {
    window.electronAPI.setCredentials({
      username: username,
      password: userPw
    }).then(result => {
      if (result.success) {
        completeLogin(username, userPw, rememberLogin);
        alert('로그인이 완료되었습니다. RPA 기능을 사용할 수 있습니다.');
      } else {
        throw new Error('로그인 정보 설정에 실패했습니다.');
      }
    }).catch(error => {
      console.error('로그인 오류:', error);
      alert('로그인에 실패했습니다: ' + error.message);
    }).finally(() => {
      loginBtn.disabled = false;
      loginBtn.textContent = '로그인';
    });
  } else {
    // Electron API가 없는 경우 (개발 환경)
    completeLogin(username, userPw, rememberLogin);
    alert('로그인이 완료되었습니다 (개발 모드).');
    
    loginBtn.disabled = false;
    loginBtn.textContent = '로그인';
  }
}

// 로그인 완료 처리
function completeLogin(username, password, rememberLogin) {
  // 전역 상태 업데이트
  globalLoginState = {
    isLoggedIn: true,
    username: username,
    password: password,
    loginTime: new Date().toISOString()
  };
  
  // 로그인 정보 저장 (사용자가 체크한 경우)
  if (rememberLogin) {
    localStorage.setItem('erpLogin', JSON.stringify(globalLoginState));
    localStorage.setItem('rememberLogin', 'true');
  } else {
    localStorage.removeItem('erpLogin');
    localStorage.setItem('rememberLogin', 'false');
  }
  
  // UI 업데이트
  updateLoginUI();
}

// 전역 로그아웃 함수
function globalLogout() {
  if (confirm('로그아웃 하시겠습니까?')) {
    globalLoginState = {
      isLoggedIn: false,
      username: '',
      password: '',
      loginTime: null
    };
    
    // 저장된 로그인 정보 제거
    localStorage.removeItem('erpLogin');
    localStorage.setItem('rememberLogin', 'false');
    
    // Electron API에서도 인증 정보 제거
    if (window.electronAPI) {
      window.electronAPI.clearCredentials();
    }
    
    // 로그인 폼 초기화
    const globalUserId = document.getElementById('globalUserId');
    const globalUserPw = document.getElementById('globalUserPw');
    const rememberLogin = document.getElementById('rememberLogin');
    
    if (globalUserId) globalUserId.value = '';
    if (globalUserPw) globalUserPw.value = '';
    if (rememberLogin) rememberLogin.checked = false;
      // UI 업데이트 - 강제로 즉시 실행
    updateLoginUI();
    
    // 추가 보장을 위해 약간의 지연 후 다시 UI 업데이트
    setTimeout(() => {
      updateLoginUI();
    }, 100);
    
    alert('로그아웃되었습니다.');
    
    // 로그아웃 후 페이지 새로고침
    setTimeout(() => {
      window.location.reload();
    }, 500);
  }
}

// 로그인 UI 상태 업데이트
function updateLoginUI() {
  const loginOverlay = document.getElementById('loginOverlay');
  const userStatus = document.getElementById('userStatus');
  const currentUser = document.getElementById('currentUser');
  const mainContent = document.querySelector('.main-content');
  
  if (globalLoginState.isLoggedIn) {
    // 로그인된 상태
    if (loginOverlay) {
      loginOverlay.classList.add('hidden');
      loginOverlay.style.display = 'none';
    }
    if (userStatus) {
      userStatus.classList.remove('hidden');
      userStatus.style.display = 'flex';
    }
    if (mainContent) {
      mainContent.classList.remove('blurred');
    }
    
    // 사용자 이름 표시 (도메인 제거)
    const displayName = globalLoginState.username.split('@')[0];
    if (currentUser) {
      currentUser.textContent = displayName;
    }
    
    // 저장된 로그인 정보가 있으면 폼에 미리 입력
    if (globalLoginState.username) {
      const globalUserId = document.getElementById('globalUserId');
      const rememberLogin = document.getElementById('rememberLogin');
      if (globalUserId) {
        globalUserId.value = globalLoginState.username.split('@')[0];
      }
      if (localStorage.getItem('rememberLogin') === 'true' && rememberLogin) {
        rememberLogin.checked = true;
      }
    }
  } else {
    // 로그아웃된 상태 - 강제로 로그인 오버레이 표시
    if (loginOverlay) {
      loginOverlay.classList.remove('hidden');
      loginOverlay.style.display = 'flex';
    }
    if (userStatus) {
      userStatus.classList.add('hidden');
      userStatus.style.display = 'none';
    }
    if (mainContent) {
      mainContent.classList.add('blurred');
    }
  }
  
  // 페이지별 버튼 상태 업데이트
  if (typeof updateButtonStates === 'function') {
    updateButtonStates();
  }
  
  // 일반 버튼 활성화/비활성화 처리
  const buttons = document.querySelectorAll('button');
  buttons.forEach(button => {
    const buttonText = button.textContent.toLowerCase();
    if (buttonText.includes('업로드') || 
        buttonText.includes('처리') || 
        buttonText.includes('다운로드') || 
        buttonText.includes('보기') || 
        buttonText.includes('내보내기')) {
      if (globalLoginState.isLoggedIn) {
        button.disabled = false;
        button.style.opacity = '1';
        button.title = '';
      } else {
        button.disabled = true;
        button.style.opacity = '0.5';
        button.title = '로그인이 필요합니다';
      }
    }
  });
}

// 로그인 상태 확인 함수 (다른 함수에서 사용)
function isLoggedIn() {
  return globalLoginState.isLoggedIn;
}

// 현재 로그인 정보 반환 함수
function getCurrentLoginInfo() {
  if (globalLoginState.isLoggedIn) {
    return {
      username: globalLoginState.username,
      password: globalLoginState.password
    };
  }
  return null;
}

// 로그인 시스템 초기화
function initializeLoginSystem() {
  // CSS 스타일 추가
  const style = document.createElement('style');
  style.textContent = getLoginOverlayCSS();
  document.head.appendChild(style);
  
  // HTML 추가
  const loginHTML = getLoginOverlayHTML();
  document.body.insertAdjacentHTML('afterbegin', loginHTML);
  
  // 저장된 로그인 정보 확인
  checkSavedLogin();
  
  // UI 업데이트
  updateLoginUI();
  
  // DOM이 완전히 로드된 후 한 번 더 UI 업데이트 (보장)
  setTimeout(() => {
    updateLoginUI();
  }, 100);
}

// 전역 스코프에 함수들 노출
window.globalLogin = globalLogin;
window.globalLogout = globalLogout;
window.isLoggedIn = isLoggedIn;
window.getCurrentLoginInfo = getCurrentLoginInfo;
window.initializeLoginSystem = initializeLoginSystem;
