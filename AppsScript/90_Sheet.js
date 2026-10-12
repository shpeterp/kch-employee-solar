// 시트 읽기/쓰기: 헤더 보정, 로그 기록, 행 → 데이터 변환

/**
 * 기존 "법인목록" 시트에 새로 추가된 헤더(응답토큰/응답유형/응답일시/변경내용)가 없으면 채워 넣음.
 * mainDaily와 웹앱 응답 처리 시작 시 자동 호출되므로, 사용자가 따로 실행할 필요는 없음.
 */
function ensureHeaders_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  const width = Math.max(sheet.getLastColumn(), HEADERS.length);
  const current = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (let i = 0; i < HEADERS.length; i++) {
    if (current[i] !== HEADERS[i]) sheet.getRange(1, i + 1).setValue(HEADERS[i]);
  }
}

/** "로그" 시트에 발송/리마인드/회신 이력을 한 줄씩 기록 */
function appendLog_(corpName, action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['일시', '법인명', '액션', '상세']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), corpName, action, detail || '']);
}

function rowToData_(row) {
  const idx = name => HEADERS.indexOf(name);
  return {
    corpName: row[idx('법인명')],
    recipientName: row[idx('수신자성함')],
    recipientTitle: row[idx('수신자직책')],
    recipientEmail: row[idx('수신자이메일')],
    cc: row[idx('참조이메일')],
    exec1Name: row[idx('임원1성함')],
    exec1Title: row[idx('임원1직책')],
    exec2Name: row[idx('임원2성함')],
    exec2Title: row[idx('임원2직책')],
    solarTarget: row[idx('태양광대상자')],
    cycleMonths: Number(row[idx('발송주기(개월)')]) || CONFIG.DEFAULT_CYCLE_MONTHS,
    nextSendDate: row[idx('다음발송예정일')] ? new Date(row[idx('다음발송예정일')]) : null,
    lastSendDate: row[idx('마지막발송일')] ? new Date(row[idx('마지막발송일')]) : null,
    status: row[idx('상태')],
    threadId: row[idx('ThreadID')],
    reminderDate: row[idx('리마인드발송일')] ? new Date(row[idx('리마인드발송일')]) : null,
    replyDate: row[idx('회신일')] ? new Date(row[idx('회신일')]) : null,
    responseToken: row[idx('응답토큰')],
    responseType: row[idx('응답유형')],
    responseAt: row[idx('응답일시')] ? new Date(row[idx('응답일시')]) : null,
    changeDetail: row[idx('변경내용')],
    escalatedDate: row[idx('무응답보고일')] ? new Date(row[idx('무응답보고일')]) : null,
  };
}
