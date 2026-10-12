// 편집기 실행 목록에서 직접 돌리는 함수(점검·미리보기·트리거)

function 트리거_켜기() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'mainDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mainDaily').timeBased().everyDays(1).atHour(9).create();
  Logger.log('매일 오전 9시 트리거가 생성되었습니다.');
}

/** mainDaily 트리거가 몇 개인지 확인 (2개 이상이면 중복 → 트리거_켜기 실행해 1개로 정리). */
function 트리거_점검() {
  const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'mainDaily');
  Logger.log('현재 mainDaily 트리거 개수: ' + ts.length + (ts.length > 1 ? ' ⚠️ 중복! 트리거_켜기()를 실행하면 1개로 정리됩니다.' : ' (정상)'));
  return ts.length;
}

/**
 * 지난번 재알림이 중간에 끊겨 못 나간 법인들에 재알림을 지금 한 번 따라잡기로 발송.
 * mainDaily와 완전히 같은 로직(실행 잠금 + 행별 오류 격리 포함)이라 안전합니다.
 * 이미 리마인드/회신완료된 곳은 다시 안 보내고, 발송완료로 7일 지난 곳에만 나갑니다.
 */
function 재알림_따라잡기() {
  mainDaily();
}

/**
 * [발송 없음 · 임시보관함 시험] 지금 리마인드가 나갈 법인(발송완료 + 7일 경과 + 회신 없음)마다
 * 등기부등본을 첨부한 리마인드를 기존 메일 스레드의 '답장 초안'(받는 사람=수신자, 참조=참조이메일)으로만 만듦.
 * 시트·로그는 바꾸지 않음.
 * 확인이 끝나면 초안은 삭제하고, 실제 발송은 재알림_따라잡기()로 할 것.
 * (초안을 직접 [보내기]하면 시트 상태가 안 바뀌어 다음 실행 때 리마인드가 한 번 더 나감)
 */
function 리마인드_초안_만들기() {
  const missing = missingProfile_();
  if (missing.length) { Logger.log(profileWarning_(missing)); return profileWarning_(missing); }
  ensureHeaders_();
  const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME).getDataRange().getValues();
  const today = stripTime_(new Date());
  const lines = [];
  for (let r = 1; r < values.length; r++) {
    const data = rowToData_(values[r]);
    if (!data.corpName || data.status !== '발송완료' || !data.threadId) continue;
    try {
      const daysSince = Math.floor((today - stripTime_(data.lastSendDate)) / (1000 * 60 * 60 * 24));
      if (daysSince < CONFIG.REMIND_AFTER_DAYS) continue;
      const thread = GmailApp.getThreadById(data.threadId);
      if (!thread) { lines.push(`[스레드없음] ${data.corpName}`); continue; }
      if (classifyIncoming_(thread).hasReply) { lines.push(`[회신있음-제외] ${data.corpName}`); continue; }
      const regFile = findRegistryFile_(data.corpName);
      if (!regFile) { lines.push(`[첨부없음] ${data.corpName}`); continue; }
      draftReminder_(thread, data, regFile);
      lines.push(`[초안작성] ${data.corpName} → 받는 사람 ${data.recipientEmail} / 참조 ${data.cc || '(없음)'} ← ${regFile.getName()}`);
    } catch (e) {
      lines.push(`[오류] ${data.corpName} - ${e}`);
    }
  }
  const msg = `리마인드 초안 ${lines.filter(l => l.indexOf('[초안작성]') === 0).length}건 작성 (발송 안 함)\n\n` + lines.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * [발송 없음 · 임시보관함 시험] 처음 메일이 어떻게 나가는지 보려고, 첫 번째 법인 기준으로
 * 등기부등본을 첨부한 처음 메일을 '본인 앞 초안'으로만 만듦. 시트·로그는 바꾸지 않음.
 */
function 처음메일_초안_미리보기() {
  const missing = missingProfile_();
  if (missing.length) { Logger.log(profileWarning_(missing)); return; }
  ensureHeaders_();
  const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME).getDataRange().getValues();
  const myEmail = Session.getEffectiveUser().getEmail();
  for (let r = 1; r < values.length; r++) {
    const data = rowToData_(values[r]);
    if (!data.corpName) continue;
    const regFile = findRegistryFile_(data.corpName);
    if (!regFile) { Logger.log(`[첨부없음] ${data.corpName}`); continue; }
    // 실제 응답토큰을 쓰면 미리보기 버튼 클릭이 그 법인 응답으로 기록되므로, 시트에 없는 임시 토큰을 씀
    const html = buildRequestHtml_(data, Utilities.getUuid());
    GmailApp.createDraft(myEmail, `[미리보기] [직원태양광] ${data.corpName} 법인 등기부등본 내용 변경 사항 알림 요청`, stripHtml_(html), {
      htmlBody: html, name: CONFIG.발신표시명, attachments: [withRetry_(() => regFile.getBlob())],
    });
    Logger.log(`처음 메일 미리보기 초안을 ${myEmail} 앞으로 만들었습니다: ${data.corpName} ← ${regFile.getName()}`);
    return;
  }
}

/**
 * [안전장치③ · 발송/재알림 전에 직접 돌려보는 점검]
 * 법인목록을 훑어 "메일이 깨질 수 있는 행"을 미리 잡아냅니다.
 *  - 수신자/참조 이메일 형식 오류
 *  - 이미 발송된(발송완료·리마인드발송) 행인데 응답토큰이 없음  ← 이번 "접근 권한 필요" 사고의 원인
 *  - 이미 발송된 행인데 ThreadID 없음(리마인드 스레드 추적 불가)
 * 결과는 실행 로그 + 본인 메일로 보내 눈으로 확인하게 함. 아무 발송도 하지 않습니다.
 */
function 발송전_점검() {
  ensureHeaders_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const issues = [];
  const missing = missingProfile_();
  if (missing.length) issues.push(profileWarning_(missing));
  const users = dashboardUsers_();
  if (!users.length) issues.push('대시보드 사용자 없음: 스크립트 속성 DASHBOARD_USERS(또는 DASHBOARD_KEY)를 넣어야 대시보드가 열립니다.');
  const rawUsers = PropertiesService.getScriptProperties().getProperty('DASHBOARD_USERS');
  if (rawUsers) {
    try { JSON.parse(rawUsers); } catch (e) { issues.push('DASHBOARD_USERS 형식 오류(JSON): ' + e); }
  }
  let checked = 0;
  for (let r = 1; r < values.length; r++) {
    const data = rowToData_(values[r]);
    if (!data.corpName) continue;
    checked++;
    const probs = [];
    if (!isValidEmail_(data.recipientEmail)) probs.push('수신자이메일 형식오류/누락');
    if (data.cc && !isValidEmail_(data.cc)) probs.push('참조이메일 형식오류');
    if (data.status === '발송완료' || data.status === '리마인드발송') {
      if (!data.responseToken) probs.push('응답토큰 없음(버튼 링크가 빈 토큰 → 접근권한 오류)');
      if (!data.threadId) probs.push('ThreadID 없음');
    }
    if (data.status === STATUS_SENDING) probs.push('발송중으로 남아 있음(다음 자동 실행 때 보낸편지함과 대조해 복구)');
    if (data.status === STATUS_CHANGE_CHECK) probs.push('변경확인필요 미처리(대시보드에서 처리완료 전까지 다음 분기 발송 보류)');
    try {
      if (!findRegistryFile_(data.corpName)) probs.push('등기부등본 파일 없음(첨부 불가 → 발송 보류됨)');
    } catch (e) {
      probs.push(`등기부등본 폴더 접근 오류: ${e}`);
    }
    if (probs.length) issues.push(`${r + 1}행 ${data.corpName}: ${probs.join(' / ')}`);
  }
  const msg = issues.length
    ? `⚠️ 점검 결과 — ${checked}개 중 ${issues.length}건 문제 발견:\n\n` + issues.join('\n')
    : `✅ 점검 완료 — ${checked}개 법인 모두 발송/재알림에 문제 없음.`;
  Logger.log(msg);
  GmailApp.sendEmail(CONFIG.SUMMARY_TO, `[점검]${issues.length ? '⚠️ ' : ''}발송전 법인목록 점검 (${formatDate_(new Date())})`, msg);
  return msg;
}

/**
 * [발송 없음] 사전 알림 메일이 어떤 내용으로 나갈지 실행 로그로만 확인합니다.
 * 발송일이 아직 멀어도 가장 가까운 발송예정일 기준으로 본문을 만들어 보여줍니다.
 */
function 사전알림_미리보기() {
  ensureHeaders_();
  const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME).getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const d = rowToData_(values[r]);
    if (d.corpName) rows.push(d);
  }
  const today = stripTime_(new Date());
  const target = nextSendTarget_(rows, today);
  if (!target) {
    Logger.log('다음발송예정일이 설정된 법인이 없습니다.');
    return;
  }
  const daysUntil = Math.round((stripTime_(target) - today) / 86400000);
  const text = `(미리보기 · 실제로 보내지 않음)\n제목: ${preNoticeSubject_(target, daysUntil)}\n\n${buildPreNoticeBody_(rows, target, daysUntil)}`;
  Logger.log(text);
  return text;
}
