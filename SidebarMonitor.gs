/**
 * 1. 자습실 시간 초과 명단 사이드바 UI 호출 (HTML 파일 분리 불필요 방식)
 * Google Sheets 우측에 2.5시간 이상 경과한 학생 명단을 보여주는 팝업창을 엽니다.
 */
function showSidebar() {
  const htmlString = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Malgun Gothic', sans-serif; font-size: 13px; padding: 12px; margin: 0; background-color: #f9f9f9; }
          h2 { font-size: 15px; color: #333; margin-top: 0; border-bottom: 2px solid #1976d2; padding-bottom: 8px; margin-bottom: 10px; }
          .status-box { background: #e3f2fd; padding: 10px; border-radius: 6px; margin-bottom: 15px; }
          .status { font-size: 12px; color: #444; line-height: 1.6; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; background-color: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
          th, td { border: 1px solid #eee; padding: 8px; text-align: center; font-size: 12px; }
          th { background-color: #f4f4f4; font-weight: bold; color: #333; }
          tr:nth-child(even) { background-color: #fafafa; }
          .passed-time { color: #d32f2f; font-weight: bold; }
          .msg { text-align: center; color: #555; padding: 30px 0; font-size: 13px; font-weight: bold; }
          .loader { text-align: center; color: #1976d2; font-weight: bold; margin-top: 40px; }
          .btn-refresh { display: block; width: 100%; padding: 10px; background-color: #1976d2; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2); transition: 0.2s; }
          .btn-refresh:hover { background-color: #115293; }
        </style>
      </head>
      <body>
        <h2>⏰ 시간 초과 명단</h2>
        
        <div class="status-box">
          <div class="status">
            마지막 확인: <b id="update-time">-</b><br>
            <span id="next-update" style="color: #1976d2; font-weight: bold;">(데이터 불러오는 중...)</span>
          </div>
        </div>

        <div id="content">
          <div class="loader">⏳ 데이터를 조회하고 있습니다...</div>
        </div>

        <button class="btn-refresh" onclick="fetchData()">🔄 지금 바로 확인하기</button>

        <script>
          const REFRESH_INTERVAL_MINUTES = 10; 
          let countdownTimer;

          // Apps Script 서버 함수 호출
          function fetchData() {
            document.getElementById('next-update').innerText = "업데이트 진행 중...";
            google.script.run
              .withSuccessHandler(updateUi)
              .withFailureHandler(showError)
              .getOvertimeData();
          }

          // UI 업데이트 처리
          function updateUi(data) {
            document.getElementById('content').innerHTML = data.html;
            document.getElementById('update-time').innerText = data.lastUpdated;
            startCountdown(REFRESH_INTERVAL_MINUTES * 60);
          }

          function showError(error) {
            document.getElementById('content').innerHTML = \`<div class="msg" style="color:red;">오류 발생: \${error.message}</div>\`;
          }

          // 다음 업데이트 자동 카운트다운
          function startCountdown(seconds) {
            if (countdownTimer) clearInterval(countdownTimer); 
            
            let timeLeft = seconds;
            countdownTimer = setInterval(() => {
              timeLeft--;
              if (timeLeft <= 0) {
                clearInterval(countdownTimer);
                fetchData(); 
              } else {
                const m = Math.floor(timeLeft / 60);
                const s = timeLeft % 60;
                document.getElementById('next-update').innerText = \`다음 자동 확인: \${m}분 \${s}초 후\`;
              }
            }, 1000);
          }

          // 초기 로드 시 바로 데이터 호출
          fetchData();
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlString)
      .setTitle('2.5시간 경과 모니터링')
      .setWidth(300);
      
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * 2. 사이드바가 10분마다 호출할 데이터 검사 함수
 * 오늘 요일 시트를 확인하여 입실 시간이 150분 이상 경과한 학생 데이터를 추출 및 시트 갱신
 */
function getOvertimeData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const days = ["일", "월", "화", "수", "목", "금", "토"];
  const todayDay = days[new Date().getDay()];
  const sheet = ss.getSheetByName(todayDay);
  
  if (!sheet || todayDay === "일") {
    return { count: 0, html: `<div class="msg">⚠️ 오늘(${todayDay}요일)에 해당하는 시트가 없습니다.</div>` };
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const numRows = values.length;
  const now = new Date();
  
  let targetCount = 0;
  let isUpdated = false;
  
  let tableHtml = `
    <table>
      <thead>
        <tr>
          <th>담당T</th>
          <th>이름</th>
          <th>경과시간</th>
        </tr>
      </thead>
      <tbody>
  `;

  for (let i = 0; i < numRows; i++) {
    // 2열: 담당T, 3열: 이름, 6열: 입실시간, 8열: 퇴실/현황, 10열: 완료상태 등 (인덱스 기준)
    const teacher = String(values[i][1] || "").trim();
    const name = String(values[i][2] || "").trim();    
    const fValue = String(values[i][5] || "").trim();  
    const hValue = String(values[i][7] || "").trim();  
    const jValue = String(values[i][9] || "").trim();  

    if (!name) continue; // 이름이 없으면 스킵
    if (jValue !== "") continue; // 이미 처리된/완료된 학생이면 스킵

    // 입실 시간 파싱 (HH:MM 형태)
    const timeMatch = fValue.match(/(\d{1,2}):(\d{2})/);
    if (!timeMatch) continue;

    const hh = parseInt(timeMatch[1], 10);
    const mm = parseInt(timeMatch[2], 10);

    const arrivalTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, mm, 0);
    const diffMs = now.getTime() - arrivalTime.getTime();
    const diffMins = Math.floor(diffMs / 60000);

    // 입실 후 150분(2.5시간) 이상 경과한 경우
    if (diffMins >= 150) {
      targetCount++;
      const timeStr = `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;

      // H열(퇴실/현황) 텍스트 업데이트
      if (!hValue.startsWith(`${timeStr}~`)) {
        values[i][7] = hValue ? `${timeStr}~ ${hValue}` : `${timeStr}~`;
        isUpdated = true;
      }
      
      // 배경색 하이라이트 추가
      if (backgrounds[i][7].toUpperCase() !== "#E0F7FA") {
        backgrounds[i][7] = "#E0F7FA";
        isUpdated = true;
      }

      const passedHours = Math.floor(diffMins / 60);
      const passedRemainingMins = diffMins % 60;
      
      tableHtml += `
        <tr>
          <td>${teacher}</td>
          <td><b>${name}</b></td>
          <td class="passed-time">${passedHours}시간 ${passedRemainingMins}분</td>
        </tr>
      `;
    }
  }

  if (targetCount === 0) {
    tableHtml = `<div class="msg">✅ 2시간 반 이상 경과한 학생이 없습니다.</div>`;
  } else {
    tableHtml += `</tbody></table>`;
  }

  // 시트에 변경사항이 있을 경우에만 일괄 반영하여 속도 최적화
  if (isUpdated) {
    dataRange.setValues(values);
    dataRange.setBackgrounds(backgrounds);
  }

  // 현재 시간 포맷팅
  const currentTimeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}`;

  return { html: tableHtml, lastUpdated: currentTimeStr };
}
