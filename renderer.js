// UI 이벤트 핸들러 및 렌더러 프로세스 로직

// DOM이 로드되기 전에도 이벤트 리스너를 등록할 수 있도록 수정
function setupEventListeners() {
  console.log('이벤트 리스너 설정 시작...');
  
  // Electron API가 있는지 확인
  if (!window.electron || !window.electron.ipcRenderer) {
    console.error('Electron IPC API를 찾을 수 없습니다.');
    return;
  }
  // 파일 범위 처리 결과 수신
  window.electron.ipcRenderer.on('file-range-processing-result', (event, result) => {
    console.log('=== 파일 범위 처리 결과 수신 ===');
    console.log('결과 데이터:', result);
    
    const button = document.getElementById('executeSelectedBtn');
    
    // 버튼 초기화
    if (button) {
      button.textContent = '선택된 파일 처리';
      button.disabled = false;
      console.log('파일 범위 처리 버튼 초기화 완료');
    } else {
      console.error('executeSelectedBtn 버튼을 찾을 수 없습니다!');
    }    if (result.success) {
      // 성공 시 상세한 정보와 함께 팝업 표시
      const successMessage = `🎉 정상적으로 RPA 동작이 마무리되었습니다! 🎉

📊 작업 결과:
• 파일 범위: ${result.startFileNumber}번 ~ ${result.endFileNumber}번
• 총 처리 파일: ${(result.endFileNumber - result.startFileNumber + 1)}개
• ✅ 성공: ${result.successCount || (result.endFileNumber - result.startFileNumber + 1)}개
• ❌ 실패: ${result.failCount || 0}개
• ⏱️ 완료 시간: ${new Date().toLocaleString()}

모든 작업이 성공적으로 완료되었습니다.`;
      
      console.log('성공 팝업 표시:', successMessage);
      alert(successMessage);
    } else {
      // 실패 시 에러 내용 포함한 상세한 팝업 표시
      const errorMessage = `❌ RPA 동작 중 오류가 발생했습니다 ❌

📊 작업 정보:
• 파일 범위: ${result.startFileNumber}번 ~ ${result.endFileNumber}번
• 총 처리 예정: ${(result.endFileNumber - result.startFileNumber + 1)}개
• ✅ 성공: ${result.successCount || 0}개
• ❌ 실패: ${result.failCount || 0}개

🚫 오류 내용:
${result.error}

⚠️ 문제 해결 방법:
1. 네트워크 연결 상태를 확인해주세요
2. VPN 연결이 필요한지 확인해주세요
3. 로그인 정보가 올바른지 확인해주세요
4. 잠시 후 다시 시도해주세요`;
      
      console.log('오류 팝업 표시:', errorMessage);
      alert(errorMessage);
    }
    
    // 버튼 상태를 다시 업데이트
    if (typeof window.updateButtonStates === 'function') {
      window.updateButtonStates();
      console.log('버튼 상태 업데이트 완료');
    } else {
      console.log('updateButtonStates 함수를 찾을 수 없습니다');
    }
  });
  // 단일 파일 처리 결과 수신
  window.electron.ipcRenderer.on('single-file-processing-result', (event, result) => {
    console.log('=== 단일 파일 처리 결과 수신 ===');
    console.log('결과 데이터:', result);
    
    const button = document.getElementById('executeSingleBtn');
    
    // 버튼 초기화
    if (button) {
      button.textContent = '단일 파일 처리';
      button.disabled = false;
      console.log('단일 파일 처리 버튼 초기화 완료');
    } else {
      console.error('executeSingleBtn 버튼을 찾을 수 없습니다!');
    }    if (result.success) {
      // 성공 시 상세한 정보와 함께 팝업 표시
      const successMessage = `🎉 정상적으로 RPA 동작이 마무리되었습니다! 🎉

📊 작업 결과:
• 처리 파일: ${result.fileNumber}번
• ✅ 상태: 성공
• ⏱️ 완료 시간: ${new Date().toLocaleString()}

단일 파일 작업이 성공적으로 완료되었습니다.`;
      
      console.log('단일파일 성공 팝업 표시:', successMessage);
      alert(successMessage);
    } else {
      // 실패 시 에러 내용 포함한 상세한 팝업 표시
      const errorMessage = `❌ RPA 동작 중 오류가 발생했습니다 ❌

📊 작업 정보:
• 처리 파일: ${result.fileNumber}번
• ❌ 상태: 실패

🚫 오류 내용:
${result.error}

⚠️ 문제 해결 방법:
1. 네트워크 연결 상태를 확인해주세요
2. VPN 연결이 필요한지 확인해주세요
3. 로그인 정보가 올바른지 확인해주세요
4. 해당 파일이 존재하는지 확인해주세요
5. 잠시 후 다시 시도해주세요`;
      
      console.log('단일파일 오류 팝업 표시:', errorMessage);
      alert(errorMessage);
    }
    
    // 버튼 상태를 다시 업데이트
    if (typeof window.updateButtonStates === 'function') {
      window.updateButtonStates();
      console.log('버튼 상태 업데이트 완료');
    } else {
      console.log('updateButtonStates 함수를 찾을 수 없습니다');
    }
  });

  console.log('이벤트 리스너 설정 완료!');
}

// 즉시 이벤트 리스너 설정 (DOM 로드 전에도 동작하도록)
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', setupEventListeners);
} else {
  setupEventListeners();
}

document.addEventListener('DOMContentLoaded', () => {
  // renderer.js 수정
  document.getElementById('loginBtn').addEventListener('click', async () => {
    const userId = document.getElementById('userId').value;
    const userPw = document.getElementById('userPw').value;
    
    console.log('로그인 버튼 클릭됨, 아이디:', userId); // 디버깅용
    
    if (!userId || !userPw) {
      alert('ID와 비밀번호를 입력해주세요.');
      return;
    }
    
    try {
      // 로딩 표시
      document.getElementById('loginBtn').textContent = '로그인 중...';
      document.getElementById('loginBtn').disabled = true;
      
      console.log('메인 프로세스로 IPC 메시지 전송...'); // 디버깅용
      
      // 메인 프로세스로 데이터 전송
      window.electron.ipcRenderer.send('start-ezvoucher', { userId, userPw });
    } catch (error) {
      console.error('IPC 전송 오류:', error); // 디버깅용
      alert('오류 발생: ' + error.message);
      document.getElementById('loginBtn').textContent = '로그인';
      document.getElementById('loginBtn').disabled = false;
    }
  });

  // 결과 수신 - 중복된 핸들러 하나만 유지
  window.electron.ipcRenderer.on('ezvoucher-status', (response) => {
    console.log('메인 프로세스로부터 응답 수신:', response); // 디버깅용
    document.getElementById('loginBtn').textContent = '로그인';
    document.getElementById('loginBtn').disabled = false;
    
    if (response.success) {
      showNotification('작업 성공', response.message, 'success');
    } else {
      showNotification('작업 실패', response.message || response.error, 'error');
    }
  });

  // 작업 실행 버튼 이벤트 핸들러
  document.querySelectorAll('[data-action="run"]').forEach(button => {
    button.addEventListener('click', async (e) => {
      const taskRow = e.target.closest('tr');
      const taskName = taskRow.querySelector('[data-field="taskName"]').textContent;
      
      // 여기에 사용자 ID와 비밀번호를 가져와서 전달해야 함
      const userId = document.getElementById('userId').value;
      const userPw = document.getElementById('userPw').value;
      
      if (!userId || !userPw) {
        showNotification('입력 오류', 'ID와 비밀번호를 입력해주세요.', 'error');
        return;
      }
      
      try {
        // 작업 상태 업데이트
        updateTaskStatus(taskRow, 'running', '실행 중');
        
        // RPA 작업 실행 요청 - ID와 PW 전달
        window.electron.ipcRenderer.send('run-specific-task', { 
          taskName, 
          userId, 
          userPw 
        });
      } catch (error) {
        updateTaskStatus(taskRow, 'error', '오류');
        showNotification('오류 발생', error.message, 'error');
      }
    });
  });
  
  // 특정 작업 실행 결과 수신
  window.electron.ipcRenderer.on('task-result', (result) => {
    const taskRow = findTaskRow(result.taskName);
    
    if (result.success) {
      updateTaskStatus(taskRow, 'done', '완료');
      showNotification('작업 완료', `"${result.taskName}" 작업이 성공적으로 완료되었습니다.`, 'success');
    } else {
      updateTaskStatus(taskRow, 'error', '오류');
      showNotification('작업 실패', result.error, 'error');
    }
  });    // 작업 상태 업데이트 이벤트 수신
  window.rpaAPI.onTaskStatusUpdate && window.rpaAPI.onTaskStatusUpdate((data) => {
    const { taskName, status, message } = data;
    const taskRow = findTaskRow(taskName);
    
    if (taskRow) {
      updateTaskStatus(taskRow, status, message);
    }
  });
});

// 알림 표시 함수
function showNotification(title, message, type = 'info') {
  // 간단한 알림 UI 표시 로직
  console.log(`[${type}] ${title}: ${message}`);
  
  // 실제 UI에 알림 표시
  const notificationElement = document.createElement('div');
  notificationElement.className = `notification ${type}`;
  notificationElement.innerHTML = `
    <strong>${title}</strong>
    <p>${message}</p>
  `;
  
  const container = document.getElementById('notificationContainer') || document.body;
  container.appendChild(notificationElement);
  
  // 3초 후 자동 제거
  setTimeout(() => {
    notificationElement.classList.add('fade-out');
    setTimeout(() => container.removeChild(notificationElement), 500);
  }, 3000);
}

// 작업 상태 업데이트 함수
function updateTaskStatus(row, status, statusText) {
  if (!row) return;
  const statusCell = row.querySelector('[data-field="status"]');
  if (statusCell) {
    statusCell.innerHTML = `<span class="badge ${status}">${statusText}</span>`;
  }
}

// 작업명으로 테이블 행 찾기
function findTaskRow(taskName) {
  const rows = document.querySelectorAll('#taskTable tbody tr');
  for (const row of rows) {
    const nameCell = row.querySelector('[data-field="taskName"]');
    if (nameCell && nameCell.textContent === taskName) {
      return row;
    }
  }
  return null;
}

// 폴더 선택 함수 추가
async function selectFolder() {
    try {
        const selectFolderBtn = document.getElementById('selectFolderBtn');
        if (selectFolderBtn) {
            selectFolderBtn.disabled = true;
            selectFolderBtn.textContent = '폴더 선택 중...';
        }

        const result = await window.electronAPI.selectFolder();
        
        if (result.success) {
            const folderPathDisplay = document.getElementById('folderPathDisplay');
            if (folderPathDisplay) {
                folderPathDisplay.textContent = `선택된 폴더: ${result.path}`;
                folderPathDisplay.style.color = 'green';
            }
            
            // 모든 실행 버튼들 활성화
            const executeBtn = document.getElementById('executeBtn');
            const executeAllBtn = document.getElementById('executeAllBtn');
            const executeSelectedBtn = document.getElementById('executeSelectedBtn');
            const executeSingleBtn = document.getElementById('executeSingleBtn');
            
            if (executeBtn) executeBtn.disabled = false;
            if (executeAllBtn) executeAllBtn.disabled = false;
            if (executeSelectedBtn) executeSelectedBtn.disabled = false;
            if (executeSingleBtn) executeSingleBtn.disabled = false;
            
            alert(`폴더가 성공적으로 선택되었습니다:\n${result.path}`);
        } else {
            alert(`폴더 선택 실패: ${result.message}`);
        }
    } catch (error) {
        console.error('폴더 선택 중 오류:', error);
        alert(`폴더 선택 중 오류가 발생했습니다: ${error.message}`);
    } finally {
        const selectFolderBtn = document.getElementById('selectFolderBtn');
        if (selectFolderBtn) {
            selectFolderBtn.disabled = false;
            selectFolderBtn.textContent = '폴더 지정';
        }
    }
}

// 선택된 범위의 파일 처리 함수
async function executeSelectedFiles() {
    try {
        const startNumber = document.getElementById('startFileNumber').value;
        const endNumber = document.getElementById('endFileNumber').value;
        
        if (!startNumber || !endNumber) {
            alert('시작 파일 번호와 끝 파일 번호를 모두 입력해주세요.');
            return;
        }
        
        const start = parseInt(startNumber);
        const end = parseInt(endNumber);
        
        if (start < 1 || end < 1 || start > end) {
            alert('올바른 파일 번호를 입력해주세요. (1 이상, 시작 번호 ≤ 끝 번호)');
            return;
        }
        
        // 버튼 비활성화
        const executeBtn = document.getElementById('executeSelectedBtn');
        if (executeBtn) {
            executeBtn.disabled = true;
            executeBtn.textContent = '처리 중...';
        }
        
        showNotification('작업 시작', `파일 ${start}번부터 ${end}번까지 처리를 시작합니다.`, 'info');
        
        const result = await window.electronAPI.processSelectedFiles(start, end);
        
        if (result.success) {
            showNotification('작업 완료', result.message, 'success');
        } else {
            showNotification('작업 실패', result.error, 'error');
        }
    } catch (error) {
        console.error('선택된 파일 처리 중 오류:', error);
        showNotification('오류 발생', `처리 중 오류가 발생했습니다: ${error.message}`, 'error');
    } finally {
        // 버튼 다시 활성화
        const executeBtn = document.getElementById('executeSelectedBtn');
        if (executeBtn) {
            executeBtn.disabled = false;
            executeBtn.textContent = '선택된 파일 처리';
        }
    }
}

// 단일 파일 처리 함수
async function executeSingleFile() {
    try {
        const fileNumber = document.getElementById('singleFileNumber').value;
        
        if (!fileNumber) {
            alert('파일 번호를 입력해주세요.');
            return;
        }
        
        const number = parseInt(fileNumber);
        
        if (number < 1) {
            alert('올바른 파일 번호를 입력해주세요. (1 이상)');
            return;
        }
        
        // 버튼 비활성화
        const executeBtn = document.getElementById('executeSingleBtn');
        if (executeBtn) {
            executeBtn.disabled = true;
            executeBtn.textContent = '처리 중...';
        }
        
        showNotification('작업 시작', `파일 ${number}번 처리를 시작합니다.`, 'info');
        
        const result = await window.electronAPI.processSingleFile(number);
        
        if (result.success) {
            showNotification('작업 완료', result.message, 'success');
        } else {
            showNotification('작업 실패', result.error, 'error');
        }
    } catch (error) {
        console.error('단일 파일 처리 중 오류:', error);
        showNotification('오류 발생', `처리 중 오류가 발생했습니다: ${error.message}`, 'error');
    } finally {
        // 버튼 다시 활성화
        const executeBtn = document.getElementById('executeSingleBtn');
        if (executeBtn) {
            executeBtn.disabled = false;
            executeBtn.textContent = '단일 파일 처리';
        }
    }
}