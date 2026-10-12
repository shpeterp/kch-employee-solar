/**
 * 법인 등기부등본 변경사항 확인 요청 - 분기별 자동 발송 + 리마인드 스크립트
 *
 * 편집기 실행 목록에 보이는 함수 (이름 끝이 _ 인 함수는 내부용이라 목록에 안 보임)
 * - 리마인드_초안_만들기   : [발송 없음] 리마인드 대상마다 등기부등본 첨부 전체답장 초안 작성
 * - 처음메일_초안_미리보기 : [발송 없음] 처음 메일을 본인 앞 초안으로 1통 작성
 * - 발송전_점검           : [발송 없음] 이메일 형식·응답토큰·등기부등본 파일 점검 결과를 본인 메일로
 * - 사전알림_미리보기      : [발송 없음] 분기 발송 사전알림 내용을 실행 로그로 확인
 * - 재알림_따라잡기        : 매일 자동 실행(mainDaily)을 지금 한 번 실행 → 리마인드 등 실제 발송
 * - 트리거_켜기           : 매일 오전 9시 자동 실행 트리거 생성(중복이면 1개로 정리)
 * - 트리거_점검           : 트리거 개수 확인
 * - mainDaily / doGet / submitResponse : 트리거·웹앱 버튼이 부르는 함수 (직접 실행할 일 없음)
 */

// 이름·직책·이메일·주소는 코드에 적지 않고 "프로젝트 설정 > 스크립트 속성"에 보관 (코드를 공개 저장소에 올리기 위함)
// 아래 8개 값이 하나라도 비어 있으면 매일 자동 실행이 메일을 보내지 않고 본인에게 알림만 보냄
const PROFILE_KEYS_ = ['STAFF_NAME', 'STAFF_TITLE', 'STAFF_EMAIL', 'MANAGER_NAME', 'MANAGER_TITLE', 'MANAGER_EMAIL', 'OFFICE_ADDRESS', 'SENDER_NAME'];
const PROPS_ = PropertiesService.getScriptProperties().getProperties();

const CONFIG = {
  SHEET_NAME: '법인목록',
  LOG_SHEET_NAME: '로그',
  REMIND_AFTER_DAYS: 7,           // 회신 없을 때 리마인드까지 대기 일수
  ESCALATE_AFTER_DAYS: 1,         // 리마인드 후에도 무응답이면 며칠 뒤 관리자에게 알릴지
  PRENOTICE_DAYS: 7,              // 분기 자동발송 며칠 전에 본인에게 사전 알림을 보낼지
  DEFAULT_CYCLE_MONTHS: 3,        // 기본 발송 주기(개월)
  SUMMARY_TO: Session.getEffectiveUser().getEmail(), // 결과 요약을 받을 본인 이메일 (트리거 실행 시 getActiveUser()는 빈 값이 될 수 있음)

  // 메일 서명·연락처에 쓰는 값 (스크립트 속성에서 읽음)
  담당자성함: PROPS_.STAFF_NAME || '',
  담당자직책: PROPS_.STAFF_TITLE || '',
  담당자이메일: PROPS_.STAFF_EMAIL || '',

  사수성함: PROPS_.MANAGER_NAME || '',
  사수직책: PROPS_.MANAGER_TITLE || '',
  사수이메일: PROPS_.MANAGER_EMAIL || '',

  본사주소: PROPS_.OFFICE_ADDRESS || '',
  발신표시명: PROPS_.SENDER_NAME || '',

  // 메일 속 [변경사항 없음] 버튼이 연결되는 웹앱(배포) 주소.
  // 코드를 "새 버전"으로 재배포해도 이 /exec 주소는 그대로 유지됩니다.
  // (혹시 "새 배포"로 새 주소가 생기면 그 주소로 여기만 바꿔주세요.)
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwGsrUErmXCc1H58qchvyX3fWcN86nuARWaqa6SAIi3ndSTHhAGjgB0nmTzbqZ7J7rooQ/exec',

  // 웹앱은 누구나(로그인 없이) 접근 가능하지만, 대시보드는 비밀키가 있는 주소로만 열립니다.
  // 비밀키는 코드에 적지 않고 "프로젝트 설정 > 스크립트 속성"의 DASHBOARD_KEY 값으로 보관합니다.
  // 대시보드 주소: <WEB_APP_URL>?key=<스크립트 속성 DASHBOARD_KEY 값>

  // 법인별 등기부등본 PDF가 들어 있는 드라이브 폴더 ID. 처음 메일·리마인드에 첨부됨.
  // 폴더를 옮기거나(공용 드라이브 포함) 폴더 이름을 바꿔도 ID는 그대로라 여기는 고칠 필요 없음.
  // 파일 이름 규칙: '등기부등본_법인명.pdf' (예: 등기부등본_티에스쏠라.pdf). '주식회사'·띄어쓰기는 빼고 비교함.
  // 분기마다 새로 뽑으면 같은 폴더에 같은 이름으로 넣으면 됨. 같은 법인 파일이 여러 개면 가장 최근에 올린 파일을 씀.
  등기부등본_폴더ID: '1tTDV0upL9afFkP0sM6zTOKjIQTWXgwYa',
};

const HEADERS = [
  '법인명', '수신자성함', '수신자직책', '수신자이메일', '참조이메일',
  '임원1성함', '임원1직책', '임원2성함', '임원2직책', '태양광대상자',
  '발송주기(개월)', '다음발송예정일', '마지막발송일', '상태',
  'ThreadID', '리마인드발송일', '회신일',
  '응답토큰', '응답유형', '응답일시', '변경내용', '무응답보고일',
];
// 상태 값: '' (대기) / '발송완료' / '리마인드발송' / '회신완료' / '발송실패'
// 다음발송예정일이 되면 이 상태들은 지난 분기 기록을 비우고 새 분기 메일을 보냄(무응답이었던 법인 포함)
const QUARTER_RESET_STATUSES = ['회신완료', '발송완료', '리마인드발송'];
// 응답유형 값: '' (미응답) / '변동없음'  ← 메일 [변경사항 없음] 버튼으로 응답 시 자동 기록
//   (변경 있음은 안내대로 이메일/서면으로 별도 통보 — 버튼 응답 대상 아님)

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

function mainDaily() {
  // 중복 트리거 등으로 mainDaily가 동시에 두 번 실행되면 이중 발송/이중 기록이 생기므로,
  // 실행 잠금으로 한 번에 하나만 돌게 함. 이미 실행 중이면 이번 건은 건너뜀.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log('다른 mainDaily 실행이 진행 중이라 이번 실행은 건너뜁니다.');
    return;
  }
  try {
    mainDailyBody_();
  } finally {
    lock.releaseLock();
  }
}

function mainDailyBody_() {
  // 서명·연락처 값이 비어 있으면 빈 서명으로 메일이 나가지 않도록 아무것도 보내지 않고 본인에게만 알림
  const missing = missingProfile_();
  if (missing.length) {
    GmailApp.sendEmail(CONFIG.SUMMARY_TO, `[직원태양광]⚠️ 스크립트 속성 확인 필요 - 오늘 발송 중단 (${formatDate_(new Date())})`, profileWarning_(missing));
    return;
  }
  ensureHeaders_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const today = stripTime_(new Date());
  const actions = [];
  const escalated = [];   // 리마인드 후에도 무응답이라 오늘 새로 관리자에게 보고할 곳
  const allData = [];     // 전체 현황 요약용

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const data = rowToData_(row);
    if (!data.corpName) continue;
    allData.push(data);

    const rowNum = r + 1;
    // 한 법인 처리 중 오류가 나도 나머지 법인은 계속 처리되도록 행별로 격리.
    try {
      processRow_(sheet, rowNum, data, today, actions, escalated);
    } catch (e) {
      actions.push(`[처리오류] ${data.corpName} - ${e}`);
      appendLog_(data.corpName, '처리오류', String(e));
    }
  }

  if (actions.length > 0) sendSummaryMail_(actions);
  // 리마인드에도 반응 없는 곳이 새로 생겼으면 관리자에게 요약 알림 (하루 뒤)
  if (escalated.length > 0) notifyEscalation_(escalated, allData);
  // 다음 분기 자동발송이 임박했으면 본인에게 사전 알림 (직원 발송 아님)
  notifyUpcomingSend_(allData, today);
}

/** 법인목록 한 행(한 법인)에 대한 발송/회신확인/리마인드/무응답보고 처리. mainDailyBody에서 행별 try로 감싸 호출. */
function processRow_(sheet, rowNum, data, today, actions, escalated) {
    // 새 분기가 도래했는데(다음발송예정일 도달) 지난 분기 기록이 남아 있으면 비우고 '대기'로 되돌린다.
    // 회신완료뿐 아니라 지난 분기 무응답(발송완료·리마인드발송)도 초기화해서 이번 분기 메일이 다시 나가게 함.
    // (발송실패는 이메일 주소를 고쳐야 하므로 초기화하지 않음 → 사전알림에 '발송되지 않는 곳'으로 표시됨)
    if (QUARTER_RESET_STATUSES.indexOf(data.status) >= 0 && data.nextSendDate && stripTime_(data.nextSendDate) <= today) {
      const prevStatus = data.status;
      ['상태', 'ThreadID', '리마인드발송일', '회신일', '응답토큰', '응답유형', '응답일시', '변경내용', '무응답보고일']
        .forEach(h => sheet.getRange(rowNum, HEADERS.indexOf(h) + 1).setValue(''));
      data.status = '';
      data.threadId = '';
      data.responseToken = '';
      const note = prevStatus === '회신완료' ? '' : ` (지난 분기 무응답: 상태=${prevStatus})`;
      appendLog_(data.corpName, '분기초기화', `${formatDate_(data.nextSendDate)} 분기 발송 대상으로 초기화${note}`);
    }

    const status = data.status;

    if (!status || status === '') {
      if (data.nextSendDate && stripTime_(data.nextSendDate) <= today) {
        // 등기부등본이 없으면 이 법인만 보류(상태를 바꾸지 않으므로 파일을 넣으면 다음 실행 때 발송됨)
        const regFile = findRegistryFile_(data.corpName);
        if (!regFile) {
          actions.push(`[첨부없음-보류] ${data.corpName} - 등기부등본 파일이 없어 처음 메일을 보내지 않음`);
          appendLog_(data.corpName, '첨부없음-보류', '처음 메일 보류 (폴더에 파일을 넣으면 다음 실행 때 발송)');
          return;
        }
        try {
          const sent = sendInitialRequest_(data, regFile);
          const targetMonth = addMonths_(today, data.cycleMonths || CONFIG.DEFAULT_CYCLE_MONTHS);
          const nextDate = firstMondayOfMonth_(targetMonth.getFullYear(), targetMonth.getMonth());
          sheet.getRange(rowNum, HEADERS.indexOf('마지막발송일') + 1).setValue(today);
          sheet.getRange(rowNum, HEADERS.indexOf('다음발송예정일') + 1).setValue(nextDate);
          sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('발송완료');
          sheet.getRange(rowNum, HEADERS.indexOf('ThreadID') + 1).setValue(sent.threadId);
          sheet.getRange(rowNum, HEADERS.indexOf('응답토큰') + 1).setValue(sent.token);
          sheet.getRange(rowNum, HEADERS.indexOf('응답유형') + 1).setValue('');
          sheet.getRange(rowNum, HEADERS.indexOf('응답일시') + 1).setValue('');
          sheet.getRange(rowNum, HEADERS.indexOf('변경내용') + 1).setValue('');
          sheet.getRange(rowNum, HEADERS.indexOf('리마인드발송일') + 1).setValue('');
          sheet.getRange(rowNum, HEADERS.indexOf('회신일') + 1).setValue('');
          sheet.getRange(rowNum, HEADERS.indexOf('무응답보고일') + 1).setValue('');
          actions.push(`[발송완료] ${data.corpName} (${data.recipientEmail})`);
          appendLog_(data.corpName, '발송완료', `수신: ${data.recipientEmail}`);
        } catch (e) {
          // 이메일 형식 오류 등 발송 자체가 실패한 경우: 완료로 남기지 않고 '발송실패'로 기록
          sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('발송실패');
          actions.push(`[발송실패] ${data.corpName} (${data.recipientEmail}) - ${e}`);
          appendLog_(data.corpName, '발송실패', `${data.recipientEmail} / ${e}`);
        }
      }
      return;
    }

    if (status === '발송완료' || status === '리마인드발송') {
      if (!data.threadId) return;
      const thread = GmailApp.getThreadById(data.threadId);
      if (!thread) return;

      // 스레드에 새 메시지가 있으면 '실제 회신'인지 '반송(주소오류 등)'인지 구분
      const incoming = classifyIncoming_(thread);

      if (incoming.hasBounce && !incoming.hasReply) {
        // 반송만 돌아온 경우: 회신이 아니라 발송실패로 처리
        sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('발송실패');
        actions.push(`[발송실패-반송] ${data.corpName} (${data.recipientEmail})`);
        appendLog_(data.corpName, '발송실패', '반송(주소 오류 등) - 실제 전달되지 않음');
        return;
      }

      if (incoming.hasReply) {
        sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('회신완료');
        sheet.getRange(rowNum, HEADERS.indexOf('회신일') + 1).setValue(today);
        actions.push(`[회신접수] ${data.corpName} (${data.recipientEmail})`);
        appendLog_(data.corpName, '회신접수', '');
        return;
      }

      if (status === '발송완료') {
        const daysSince = Math.floor((today - stripTime_(data.lastSendDate)) / (1000 * 60 * 60 * 24));
        if (daysSince >= CONFIG.REMIND_AFTER_DAYS) {
          // 등기부등본이 없으면 이 법인 리마인드만 보류(상태 유지 → 파일을 넣으면 다음 실행 때 발송됨)
          const regFile = findRegistryFile_(data.corpName);
          if (!regFile) {
            actions.push(`[첨부없음-보류] ${data.corpName} - 등기부등본 파일이 없어 리마인드를 보내지 않음`);
            appendLog_(data.corpName, '첨부없음-보류', '리마인드 보류 (폴더에 파일을 넣으면 다음 실행 때 발송)');
            return;
          }
          // 옛 코드로 최초발송된 행은 응답토큰이 없어 리마인드 버튼 링크가 빈 토큰이 됨 → 없으면 생성해 기록.
          if (!data.responseToken) {
            data.responseToken = Utilities.getUuid();
            sheet.getRange(rowNum, HEADERS.indexOf('응답토큰') + 1).setValue(data.responseToken);
          }
          sendReminder_(thread, data, regFile);
          sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('리마인드발송');
          sheet.getRange(rowNum, HEADERS.indexOf('리마인드발송일') + 1).setValue(today);
          actions.push(`[리마인드발송] ${data.corpName} (${data.recipientEmail})`);
          appendLog_(data.corpName, '리마인드발송', '');
        }
      } else if (status === '리마인드발송') {
        // 리마인드까지 보냈는데도 응답(버튼/회신)이 없는 경우: 리마인드 후 N일 지나면 관리자에게 1회 보고.
        // (직원에게는 아무것도 안 감. 무응답보고일이 이미 있으면 재보고하지 않음.)
        const daysSinceRemind = data.reminderDate
          ? Math.floor((today - stripTime_(data.reminderDate)) / (1000 * 60 * 60 * 24)) : 0;
        if (!data.escalatedDate && data.reminderDate && daysSinceRemind >= CONFIG.ESCALATE_AFTER_DAYS) {
          sheet.getRange(rowNum, HEADERS.indexOf('무응답보고일') + 1).setValue(today);
          appendLog_(data.corpName, '리마인드무응답', `리마인드 ${formatDate_(data.reminderDate)} 후 ${daysSinceRemind}일 무응답`);
          escalated.push(data);
        }
      }
    }
}

function sendInitialRequest_(data, regFile) {
  if (!isValidEmail_(data.recipientEmail)) {
    throw new Error(`수신자 이메일 형식 오류: "${data.recipientEmail}"`);
  }
  if (data.cc && !isValidEmail_(data.cc)) {
    throw new Error(`참조 이메일 형식 오류: "${data.cc}"`);
  }
  const subject = `[직원태양광] ${data.corpName} 법인 등기부등본 내용 변경 사항 알림 요청`;
  const token = data.responseToken || Utilities.getUuid();
  const html = buildRequestHtml_(data, token);
  const options = { htmlBody: html, name: CONFIG.발신표시명, attachments: [withRetry_(() => regFile.getBlob())] };
  if (data.cc) options.cc = data.cc;
  const draft = GmailApp.createDraft(data.recipientEmail, subject, stripHtml_(html), options);
  const message = draft.send();
  return { threadId: message.getThread().getId(), token: token };
}

/**
 * 리마인드 발송. 원래 메일 스레드에 이어 붙이되 받는 사람=수신자, 참조=참조이메일로 직접 지정함.
 * (GmailApp의 replyAll은 원래 메일을 보낸 '나'를 받는 사람으로 잡아 수신자가 참조로 밀려나므로 Gmail API로 보냄)
 */
function sendReminder_(thread, data, regFile) {
  Gmail.Users.Messages.send({ raw: buildReminderRaw_(thread, data, regFile), threadId: thread.getId() }, 'me');
}

/** 리마인드를 같은 방식(받는 사람 지정 + 스레드 유지)으로 임시보관함 초안으로만 만듦 */
function draftReminder_(thread, data, regFile) {
  Gmail.Users.Drafts.create({ message: { raw: buildReminderRaw_(thread, data, regFile), threadId: thread.getId() } }, 'me');
}

/** 리마인드 메일 원문(MIME) 생성: 본문(텍스트+HTML) + 등기부등본 PDF 첨부, 원래 메일에 대한 답장 헤더 포함 */
function buildReminderRaw_(thread, data, regFile) {
  const messages = thread.getMessages();
  const last = messages[messages.length - 1];
  const refs = messages.map(m => m.getHeader('Message-ID')).filter(Boolean);
  const subject = 'Re: ' + messages[0].getSubject().replace(/^(re:\s*)+/i, '');
  const html = buildReminderHtml_(data, data.responseToken);
  const blob = withRetry_(() => regFile.getBlob());
  const fileName = blob.getName();
  const mixed = 'mixed_' + Utilities.getUuid();
  const alt = 'alt_' + Utilities.getUuid();
  const b64 = (v) => Utilities.base64Encode(v, Utilities.Charset.UTF_8).replace(/.{76}/g, '$&\r\n');
  const word = (s) => '=?UTF-8?B?' + Utilities.base64Encode(s, Utilities.Charset.UTF_8) + '?=';

  const headers = [
    `From: ${word(CONFIG.발신표시명)} <${Session.getEffectiveUser().getEmail()}>`,
    `To: ${data.recipientEmail}`,
    data.cc ? `Cc: ${data.cc}` : '',
    `Subject: ${word(subject)}`,
    last.getHeader('Message-ID') ? `In-Reply-To: ${last.getHeader('Message-ID')}` : '',
    refs.length ? `References: ${refs.join(' ')}` : '',
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${mixed}"`,
  ].filter(Boolean);

  const body = [
    `--${mixed}`,
    `Content-Type: multipart/alternative; boundary="${alt}"`,
    '',
    `--${alt}`,
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    b64(stripHtml_(html)),
    `--${alt}`,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    b64(html),
    `--${alt}--`,
    `--${mixed}`,
    `Content-Type: application/pdf; name="${word(fileName)}"`,
    `Content-Disposition: attachment; filename="${word(fileName)}"; filename*=UTF-8''${encodeURIComponent(fileName)}`,
    'Content-Transfer-Encoding: base64',
    '',
    Utilities.base64Encode(blob.getBytes()).replace(/.{76}/g, '$&\r\n'),
    `--${mixed}--`,
    '',
  ];
  return Utilities.base64EncodeWebSafe(headers.join('\r\n') + '\r\n\r\n' + body.join('\r\n'));
}

/** 드라이브 호출이 일시적으로 실패(Access denied 등)하면 잠시 쉬었다가 최대 3번까지 다시 시도 */
function withRetry_(fn) {
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      Utilities.sleep(2000 * (i + 1));
    }
  }
  throw lastErr;
}

/** 법인명 비교용 정규화: '주식회사'·'(주)'·띄어쓰기 제거 */
function normalizeCorpName_(name) {
  return String(name || '').normalize('NFC').replace(/주식회사|\(주\)|㈜/g, '').replace(/\s+/g, '');
}

// 등기부등본 폴더 파일 목록. 한 번 실행하는 동안 폴더를 한 번만 읽도록 보관함
let registryFilesCache_ = null;

function registryFiles_() {
  if (!registryFilesCache_) {
    registryFilesCache_ = withRetry_(() => {
      const list = [];
      const it = DriveApp.getFolderById(CONFIG.등기부등본_폴더ID).getFiles();
      while (it.hasNext()) {
        const f = it.next();
        if (!f.isTrashed()) list.push(f);
      }
      return list;
    });
  }
  return registryFilesCache_;
}

/**
 * 등기부등본 폴더에서 법인의 PDF를 찾아 반환. 없으면 null.
 * 파일 이름에서 '등기부등본_'과 '.pdf'를 뺀 나머지를 법인명과 비교함.
 * 같은 법인 파일이 여러 개면 가장 최근에 올린 파일을 씀.
 */
function findRegistryFile_(corpName) {
  const target = normalizeCorpName_(corpName);
  let latest = null;
  registryFiles_().forEach(f => {
    const name = f.getName().normalize('NFC').replace(/\.pdf$/i, '').replace(/^등기부등본[_\s-]*/, '');
    if (normalizeCorpName_(name) !== target) return;
    if (!latest || f.getDateCreated() > latest.getDateCreated()) latest = f;
  });
  return latest;
}

function sendSummaryMail_(actions) {
  const isIssue = a => a.indexOf('[발송실패') >= 0 || a.indexOf('[첨부없음') >= 0;
  const fails = actions.filter(isIssue);
  const oks = actions.filter(a => !isIssue(a));

  let body = `오늘(${formatDate_(new Date())}) 법인 등기부등본 확인 요청 자동화 처리 내역입니다.\n\n`;
  if (fails.length > 0) {
    body += `⚠️ 발송 실패 / 반송 / 첨부 없음 ${fails.length}건 — 확인이 필요합니다.\n`;
    body += fails.map(a => '- ' + a).join('\n') + '\n\n';
  }
  body += `처리 완료 ${oks.length}건\n`;
  body += (oks.length ? oks.map(a => '- ' + a).join('\n') : '- (없음)');

  const flag = fails.length > 0 ? '⚠️[확인 필요] ' : '';
  GmailApp.sendEmail(
    CONFIG.SUMMARY_TO,
    `[자동화]${flag}등기부등본 확인 요청 처리 결과 (${formatDate_(new Date())})`,
    body
  );
}

/**
 * 리마인드까지 보냈는데도 무응답인 곳을 관리자(담당자)에게 요약 보고. 직원에게는 발송되지 않음.
 * escalated: 오늘 새로 무응답 확정된 법인(data) 배열 / rows: 전체 법인 data 배열(현황 요약용).
 */
function notifyEscalation_(escalated, rows) {
  const stuck = rows.filter(d => d.status === '리마인드발송'); // 리마인드 후 아직 무응답 전체
  const counts = {};
  rows.forEach(d => { const k = d.status || '대기'; counts[k] = (counts[k] || 0) + 1; });

  let body = `리마인드 메일까지 발송했으나 응답(버튼/회신)이 없는 법인 안내입니다.\n`;
  body += `조치가 필요한 곳을 확인해 주세요.\n\n`;
  body += `■ 오늘 새로 '무응답'으로 확인된 곳 (리마인드 후 ${CONFIG.ESCALATE_AFTER_DAYS}일 경과): ${escalated.length}곳\n`;
  body += escalated.map(d => `- ${d.corpName} / ${d.recipientName} ${d.recipientTitle} (${d.recipientEmail}) · 리마인드 ${formatDate_(d.reminderDate)}`).join('\n') + '\n\n';
  body += `■ 현재 리마인드 후 무응답 상태 전체: ${stuck.length}곳\n`;
  body += (stuck.length ? stuck.map(d => `- ${d.corpName} / ${d.recipientName} ${d.recipientTitle}`).join('\n') : '- (없음)') + '\n\n';
  body += `■ 전체 현황: ` + Object.keys(counts).map(k => `${k} ${counts[k]}`).join(' / ');

  GmailApp.sendEmail(
    CONFIG.SUMMARY_TO,
    `⚠️[조치필요] 리마인드 무응답 ${escalated.length}곳 (${formatDate_(new Date())})`,
    body
  );
}

/**
 * 분기 자동발송 사전 알림. 다음 발송예정일이 PRENOTICE_DAYS 이내로 다가오면
 * 본인(CONFIG.SUMMARY_TO)에게 1회만 알린다. 직원에게는 아무것도 보내지 않는다.
 * 같은 발송예정일로 중복 발송하지 않도록 스크립트 속성에 보낸 날짜를 기록한다.
 */
function notifyUpcomingSend_(rows, today) {
  const target = nextSendTarget_(rows, today);
  if (!target) return;
  const daysUntil = Math.round((stripTime_(target) - today) / 86400000);
  if (daysUntil <= 0 || daysUntil > CONFIG.PRENOTICE_DAYS) return;

  const props = PropertiesService.getScriptProperties();
  const key = 'PRENOTICE_' + formatDate_(target);
  if (props.getProperty(key)) return;

  GmailApp.sendEmail(CONFIG.SUMMARY_TO, preNoticeSubject_(target, daysUntil), buildPreNoticeBody_(rows, target, daysUntil));
  props.setProperty(key, formatDate_(today));
  appendLog_('(전체)', '사전알림', `${formatDate_(target)} 발송 D-${daysUntil} 안내`);
}

/** 오늘 이후로 가장 가까운 다음발송예정일 (지난 날짜가 박혀 있는 행은 무시) */
function nextSendTarget_(rows, today) {
  const times = rows
    .filter(d => d.nextSendDate)
    .map(d => stripTime_(d.nextSendDate).getTime())
    .filter(t => t >= today.getTime());
  return times.length ? new Date(Math.min.apply(null, times)) : null;
}

function preNoticeSubject_(target, daysUntil) {
  return `[사전알림] 분기 등기부등본 확인 요청 D-${daysUntil} 자동발송 예정 (${formatDate_(target)})`;
}

/** 사전 알림 본문만 생성(발송하지 않음). 사전알림_미리보기()에서도 같은 함수를 씀. */
function buildPreNoticeBody_(rows, target, daysUntil) {
  const t = stripTime_(target).getTime();
  const due = rows.filter(d => d.nextSendDate && stripTime_(d.nextSendDate).getTime() === t);
  const ready = due.filter(d => !d.status || QUARTER_RESET_STATUSES.indexOf(d.status) >= 0);
  const blocked = due.filter(d => d.status && QUARTER_RESET_STATUSES.indexOf(d.status) < 0);
  const unanswered = ready.filter(d => d.status === '발송완료' || d.status === '리마인드발송');
  const badEmail = ready.filter(d => !isValidEmail_(d.recipientEmail));

  let body = `${formatDate_(target)}(D-${daysUntil}) 오전 9시에 분기 등기부등본 확인 요청 메일이 자동 발송될 예정입니다.\n\n`;
  body += `■ 발송 예정: ${ready.length}곳\n`;
  body += (ready.length
    ? ready.map(d => `- ${d.corpName} / ${d.recipientName} ${d.recipientTitle} (${d.recipientEmail})`).join('\n')
    : '- (없음)') + '\n\n';

  if (unanswered.length) {
    body += `ℹ️ 이 중 지난 분기 무응답 ${unanswered.length}곳 — 이번 분기 메일이 다시 나갑니다\n`;
    body += unanswered.map(d => `- ${d.corpName} : 상태=${d.status}`).join('\n') + '\n\n';
  }

  if (blocked.length) {
    body += `⚠️ 이대로면 발송되지 않는 곳: ${blocked.length}곳 — 상태 칸을 비우면 발송 대상이 됩니다\n`;
    body += blocked.map(d => `- ${d.corpName} : 상태=${d.status}`).join('\n') + '\n\n';
  }
  if (badEmail.length) {
    body += `⚠️ 수신자 이메일 형식 오류: ${badEmail.length}곳 — 발송일 전에 고쳐주세요\n`;
    body += badEmail.map(d => `- ${d.corpName} : "${d.recipientEmail}"`).join('\n') + '\n\n';
  }

  body += `발송을 미루거나 건너뛰시려면 발송일 전에 해당 행의 '다음발송예정일'을 바꾸거나 비워두시면 됩니다.\n\n`;
  body += `대시보드: ${CONFIG.WEB_APP_URL}?key=${dashboardKey_() || '(스크립트 속성 DASHBOARD_KEY 미설정)'}`;
  return body;
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

/** 스레드의 새 메시지를 훑어서 '실제 회신'과 '반송(bounce)'을 구분 */
function classifyIncoming_(thread) {
  const myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const messages = thread.getMessages();
  let hasBounce = false;
  let hasReply = false;
  for (const m of messages) {
    const from = (m.getFrom() || '').toLowerCase();
    const subject = (m.getSubject() || '').toLowerCase();
    const isBounce =
      from.indexOf('mailer-daemon') >= 0 ||
      from.indexOf('postmaster') >= 0 ||
      subject.indexOf('delivery status notification') >= 0 ||
      subject.indexOf('undelivered') >= 0 ||
      subject.indexOf('undeliverable') >= 0 ||
      subject.indexOf('전송 실패') >= 0 ||
      subject.indexOf('반송') >= 0;
    if (isBounce) { hasBounce = true; continue; }
    // 내가 보낸 최초 메일/리마인드는 회신이 아님 → 건너뜀
    if (myEmail && from.indexOf(myEmail) >= 0) continue;
    hasReply = true;
  }
  return { hasBounce, hasReply };
}

/** 이메일 형식이 올바른지 간단히 검사 */
function isValidEmail_(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
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

/** 메일 하단에 넣는 [변경사항 없음] 버튼 블록 */
function buildResponseButtons_(token) {
  // [안전장치①] 토큰이 비면 버튼 링크가 '?token=' 로 깨져 직원이 "접근 권한 필요"를 보게 됨.
  // 그런 메일은 아예 만들지 않고 발송을 실패시켜(→ 발송실패 기록 + 요약메일 경고) 눈에 띄게 함.
  if (!token) throw new Error('응답토큰이 비어 있어 버튼 메일을 만들 수 없습니다. (발송전_점검 실행 필요)');
  const noneUrl = `${CONFIG.WEB_APP_URL}?token=${encodeURIComponent(token)}`;
  return `
<div style="margin:22px 0 4px;">
  <div style="font-size:13px; color:#333; margin-bottom:10px;">
    <b>변경사항이 없으신 경우</b> 아래 <b>[변경사항 없음]</b> 버튼을 눌러주세요.<br>
    변경사항이 있으신 경우에는 위 안내대로 <b>2주 이내</b>에 인사총무팀으로 <b>이메일 또는 서면</b>으로 통보해 주시기 바랍니다.
  </div>
  <table cellpadding="0" cellspacing="0" border="0"><tr>
    <td>
      <a href="${noneUrl}" style="display:inline-block; background:#15803d; color:#fff; text-decoration:none; font-size:14px; font-weight:bold; padding:12px 22px; border-radius:8px;">✅ 변경사항 없음</a>
    </td>
  </tr></table>
  <div style="font-size:12px; color:#888; margin-top:9px;">
    * 버튼을 누르면 확인 화면이 한 번 뜨고, 거기서 [제출]을 눌러야 최종 접수됩니다.
  </div>
</div>`;
}

function buildRequestHtml_(d, token) {
  return `
<div style="font-size:14px; line-height:1.6; color:#000;">
안녕하십니까 ${d.recipientName} ${d.recipientTitle}님<br>
${CONFIG.발신표시명} ${CONFIG.담당자성함} ${CONFIG.담당자직책}입니다.<br><br>

당 그룹사에서 운용중인 직원태양광 법인 정보가 변경될 경우,<br>
경영지원본부 인사총무팀에 변경사항을 알려주시기를 요청드리려 메일드렸습니다.<br><br>

요청사항은 아래와 같습니다.<br>
-------------------------------------------------------------------<br>

<table cellpadding="0" cellspacing="0" border="0" style="font-size:14px; line-height:1.6; color:#000;">
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
등기부등본 내 정보의 변동사항 발생 시(<u>대표이사 주소지 변경 등</u>),<br>
변경일로부터 <b><span style="color:#0000FF;">2주 이내</span></b> 인사총무팀에 이메일 또는 서면으로 변경된 정보 통보
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
통보 시, <b><span style="color:#0000FF;">변경 내용과 변경된 날짜</span></b>를 명확히 기재
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
<b>첨부드린 등기부등본</b>에서 아래 항목이 현재와 같은지 확인<br>
${buildCheckItemsHtml_()}
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
법인 정보<br>
- 법인명: ${d.corpName}<br>
- 임원: ${d.exec1Name} ${d.exec1Title}${d.exec2Name ? ` / ${d.exec2Name} ${d.exec2Title}` : ''}<br>
- 태양광 대상자: ${d.solarTarget}
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
인사총무팀 ${CONFIG.담당자성함} ${CONFIG.담당자직책} E-mail: ${CONFIG.담당자이메일}<br>
인사총무팀 ${CONFIG.사수성함} ${CONFIG.사수직책} E-mail: ${CONFIG.사수이메일}<br>
본사 사무실 주소: ${CONFIG.본사주소}
</td></tr>
</table>

-------------------------------------------------------------------<br>

<b>변경 사항을 2주 내로 보고하지 않을 경우 과태료가 부과되오니, 신속하게 변경 사항을 통보해주시기 바랍니다.<br>
미통보로 인한 과태료 발생시 당사에서는 어떠한 책임도 지지 않음을 알려드립니다.</b><br><br>

내용 확인하여 주시고, 관련하여 궁금하신 사항은 언제든 편하게 문의주시면 감사하겠습니다.<br>

감사합니다.<br>
${buildResponseButtons_(token)}
<br>
<span style="text-decoration:underline; font-size:12px;">*본 메일은 분기마다 정기적으로 발송되는 메일입니다.</span>
</div>`;
}

/** 처음 메일·리마인드 공통: 첨부한 등기부등본에서 확인할 항목 */
function buildCheckItemsHtml_() {
  return `- 대표이사 주소<br>
- 임원 구성(대표이사·사내이사)<br>
- 본점 주소`;
}

function buildReminderHtml_(d, token) {
  return `
<div style="font-size:14px; line-height:1.6; color:#000;">
안녕하십니까 ${d.recipientName} ${d.recipientTitle}님<br>
${CONFIG.발신표시명} ${CONFIG.담당자성함} ${CONFIG.담당자직책}입니다.<br><br>

지난 ${formatDate_(d.lastSendDate)} ${d.corpName} 법인 등기부등본 변경사항 관련하여 안내드린 메일에 대해<br>
아직 회신을 받지 못하여 다시 한 번 안내드립니다.<br><br>

확인이 쉽도록 <b>${d.corpName} 등기부등본을 첨부</b>하였습니다.<br>
첨부 파일에서 아래 항목이 현재와 같은지 확인 부탁드립니다.<br>
${buildCheckItemsHtml_()}<br><br>

바쁘시겠지만 확인 부탁드립니다.<br>
감사합니다.
${buildResponseButtons_(token)}
</div>`;
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

function stripTime_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function addMonths_(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

/** 해당 연/월(월은 0-indexed)의 첫째 주 월요일을 반환 */
function firstMondayOfMonth_(year, month) {
  const d = new Date(year, month, 1);
  const offset = (1 - d.getDay() + 7) % 7;
  d.setDate(1 + offset);
  return d;
}

function formatDate_(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Seoul', 'yyyy-MM-dd');
}

/** 스크립트 속성에서 비어 있는 서명·연락처 값의 이름 목록 (모두 있으면 빈 배열) */
function missingProfile_() {
  return PROFILE_KEYS_.filter(k => !String(PROPS_[k] || '').trim());
}

function profileWarning_(missing) {
  return `스크립트 속성 확인 필요 — 비어 있는 값: ${missing.join(', ')}\n` +
    '편집기 "프로젝트 설정 > 스크립트 속성"에서 값을 넣기 전까지 메일을 보내지 않습니다.';
}

/** 대시보드 비밀키. "프로젝트 설정 > 스크립트 속성"의 DASHBOARD_KEY 값 (없으면 빈 값). */
function dashboardKey_() {
  return PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY') || '';
}

function stripHtml_(html) {
  return html.replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ');
}

/**
 * 웹 앱(Web App)으로 배포했을 때 접속하면 실행되는 함수.
 * "배포 > 새 배포 > 웹 앱"으로 배포하면 이 함수가 만든 페이지가 보입니다.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  // 메일 버튼에서 온 요청(token 파라미터 존재)은 응답 화면으로. 토큰이 비어 있어도
  // 대시보드 경로로 새지 않도록 여기서 처리(빈 토큰이면 '잘못된 링크' 안내가 뜸).
  if ('token' in p) {
    return HtmlService.createHtmlOutput(renderResponsePage_(p.token))
      .setTitle('등기부등본 변경사항 응답')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 대시보드는 비밀키가 있는 주소로만 열림 (수신자가 토큰만 지워도 내부 데이터가 안 보이도록).
  // 스크립트 속성에 키가 없으면 아무 주소로도 열리지 않음.
  const dashboardKey = dashboardKey_();
  if (!dashboardKey || p.key !== dashboardKey) {
    return HtmlService.createHtmlOutput(responseShell_(
      '접근 권한 필요',
      '<p>이 페이지는 접근 권한이 필요합니다.</p><p class="muted">담당자에게 정식 주소를 문의해 주세요.</p>',
      ''
    )).setTitle('접근 권한 필요');
  }
  const data = getDashboardData_();
  const html = renderDashboardHtml_(data);
  return HtmlService.createHtmlOutput(html)
    .setTitle('직원태양광 법인 등기부등본 변경사항 알림')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 토큰으로 법인목록에서 해당 행을 찾음. 없으면 null. */
function findRowByToken_(token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const tokenCol = HEADERS.indexOf('응답토큰');
  for (let r = 1; r < values.length; r++) {
    if (values[r][tokenCol] && String(values[r][tokenCol]) === String(token)) {
      return { rowNum: r + 1, data: rowToData_(values[r]) };
    }
  }
  return null;
}

/**
 * 메일 버튼을 눌렀을 때 뜨는 확인/입력 화면.
 * 실제 시트 기록은 여기서 하지 않고, [제출] 버튼을 눌러 submitResponse()가 호출될 때 이뤄짐
 * (메일 보안 스캐너가 링크를 자동으로 여는 것만으로는 접수되지 않도록).
 */
function renderResponsePage_(token) {
  ensureHeaders_();
  const found = findRowByToken_(token);
  if (!found) {
    // [안전장치②] 빈/잘못된 토큰 클릭이 오면 로그에 남김 → 깨진 링크가 나갔는지 사후에 바로 확인 가능.
    appendLog_('(알수없음)', '이상접근', `유효하지 않은 응답 링크 클릭: token="${token || '(빈값)'}"`);
    return responseShell_('잘못된 링크입니다', '이미 처리되었거나 유효하지 않은 링크입니다.<br>궁금하신 점은 인사총무팀으로 문의해 주세요.', '');
  }

  const d = found.data;
  const corp = escapeHtml_(d.corpName);
  const already = d.responseType
    ? `<div class="note">이미 <b>${escapeHtml_(d.responseType)}</b>(${formatDateTime_(d.responseAt)})으로 접수된 건입니다. 다시 제출하면 최신 내용으로 갱신됩니다.</div>`
    : '';

  const body = `
  <div class="corp">${corp}</div>
  <div class="lead">법인 등기부등본 변경사항 확인</div>
  ${already}
  <div id="form">
    <p>등기부등본에 <b>변경사항이 없으시면</b> 아래 버튼을 눌러 주세요.</p>
    <p class="muted">변경사항이 있으신 경우에는 이 버튼 대신, 안내 메일에 따라 <b>2주 이내</b>에 인사총무팀으로 <b>이메일 또는 서면</b>으로 통보해 주세요.</p>
    <button class="btn btn-green" onclick="submitNone()">변경사항 없음으로 제출</button>
  </div>
  <div id="busy" class="busy" style="display:none;">제출 중입니다…</div>
  `;

  const script = `
  <script>
    var TOKEN = ${JSON.stringify(token)};
    function lock(){ document.getElementById('form').style.display='none'; document.getElementById('busy').style.display='block'; }
    function done(res){
      var box = document.getElementById('card');
      box.innerHTML = '<div class="corp">'+res.corp+'</div><div class="ok">✅ 제출 완료</div><p>'+res.msg+'</p><p class="muted">이 창은 닫으셔도 됩니다. 감사합니다.</p>';
    }
    function fail(err){ document.getElementById('busy').style.display='none'; document.getElementById('form').style.display=''; alert('제출 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.\\n\\n'+err.message); }
    function submitNone(){
      lock();
      google.script.run.withSuccessHandler(done).withFailureHandler(fail).submitResponse(TOKEN);
    }
  </script>`;

  return responseShell_('등기부등본 변경사항 응답', body, script);
}

/** 응답 화면 공통 껍데기(브랜드 스타일). */
function responseShell_(title, innerHtml, script) {
  return `<!DOCTYPE html>
<html><head><base target="_top">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body{margin:0;background:#f3f5fa;font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic","Noto Sans KR",sans-serif;color:#1b2033;}
  .wrap{max-width:520px;margin:0 auto;padding:40px 18px;}
  .top{background:#286FB7;color:#fff;border-radius:14px 14px 0 0;padding:18px 22px;font-size:14px;font-weight:700;}
  #card{background:#fff;border-radius:0 0 14px 14px;box-shadow:0 2px 10px rgba(16,24,40,.08);padding:26px 22px 30px;}
  .corp{font-size:20px;font-weight:800;margin-bottom:4px;}
  .lead{font-size:13px;color:#6b7280;margin-bottom:18px;}
  .ok{font-size:20px;font-weight:800;color:#15803d;margin:6px 0 10px;}
  p{font-size:15px;line-height:1.6;}
  .muted{color:#6b7280;font-size:13px;}
  .note{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;font-size:13px;border-radius:8px;padding:10px 12px;margin-bottom:16px;}
  .btn{display:block;width:100%;border:none;border-radius:10px;color:#fff;font-size:16px;font-weight:700;padding:15px;cursor:pointer;margin-top:6px;}
  .btn-green{background:#15803d;}
  .busy{text-align:center;color:#6b7280;padding:20px;}
</style></head>
<body><div class="wrap">
  <div class="top">케이씨에이치그룹 경영지원본부 인사총무팀</div>
  <div id="card">${innerHtml}</div>
</div>${script || ''}</body></html>`;
}

/**
 * [변경사항 없음으로 제출] 버튼을 눌렀을 때 실제로 시트에 '변동없음' 응답을 기록.
 * 여러 명이 동시에 눌러도 안전하도록 잠금(LockService) 사용.
 */
function submitResponse(token) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureHeaders_();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const found = findRowByToken_(token);
    if (!found) throw new Error('유효하지 않은 링크');
    const rowNum = found.rowNum;
    const d = found.data;
    const now = new Date();

    sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('회신완료');
    sheet.getRange(rowNum, HEADERS.indexOf('응답유형') + 1).setValue('변동없음');
    sheet.getRange(rowNum, HEADERS.indexOf('응답일시') + 1).setValue(now);
    sheet.getRange(rowNum, HEADERS.indexOf('회신일') + 1).setValue(stripTime_(now));

    appendLog_(d.corpName, '회신접수', '버튼응답: 변동없음');
    notifyResponse_(d);

    return { corp: escapeHtml_(d.corpName), msg: '변경사항 없음으로 정상 접수되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

/** '변경사항 없음' 응답이 들어오면 담당자에게 알림 메일 발송 */
function notifyResponse_(d) {
  GmailApp.sendEmail(
    CONFIG.담당자이메일 || CONFIG.SUMMARY_TO,
    `[응답:변동없음] ${d.corpName}`,
    `${d.corpName} 법인이 등기부등본 '변경사항 없음'으로 확인 응답했습니다.\n수신자: ${d.recipientName} ${d.recipientTitle} (${d.recipientEmail})\n접수시각: ${formatDateTime_(new Date())}`
  );
}

function getDashboardData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const data = rowToData_(values[r]);
    if (!data.corpName) continue;
    rows.push(data);
  }

  const counts = { '대기': 0, '발송완료': 0, '리마인드발송': 0, '회신완료': 0, '발송실패': 0 };
  rows.forEach(d => {
    const key = d.status || '대기';
    if (counts[key] === undefined) counts[key] = 0;
    counts[key]++;
  });

  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  let allLogs = [];
  if (logSheet && logSheet.getLastRow() > 1) {
    allLogs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues()
      .filter(row => row[0]);
  }
  const logs = allLogs.slice().reverse().slice(0, 50);
  const quarterly = buildQuarterlyTracking_(allLogs, rows);

  return { rows, counts, logs, quarterly, generatedAt: new Date() };
}

/**
 * 로그 시트 전체 이력을 법인별로 재생(replay)해서, 연도 x 분기별 발송 사이클을 재구성.
 * 발송완료로 사이클이 시작되고, 그 뒤 리마인드발송/회신접수가 오면 같은 분기 셀의 상태를 갱신.
 */
function buildQuarterlyTracking_(allLogs, corpRows) {
  const currentYear = new Date().getFullYear();
  const byCorp = {};
  allLogs.forEach(row => {
    const corp = row[1];
    if (!byCorp[corp]) byCorp[corp] = [];
    byCorp[corp].push(row);
  });

  const byYear = {};
  const yearsSet = new Set([currentYear]);

  Object.keys(byCorp).forEach(corp => {
    let cycle = null;
    byCorp[corp].forEach(row => {
      const date = new Date(row[0]);
      const action = row[2];
      if (action === '발송완료') {
        const year = date.getFullYear();
        const quarter = Math.floor(date.getMonth() / 3) + 1;
        cycle = { status: '발송완료', sentDate: date };
        yearsSet.add(year);
        if (!byYear[year]) byYear[year] = {};
        if (!byYear[year][corp]) byYear[year][corp] = {};
        byYear[year][corp][quarter] = cycle;
      } else if (action === '리마인드발송' && cycle) {
        cycle.status = '리마인드발송';
      } else if (action === '회신접수' && cycle) {
        cycle.status = '회신완료';
      } else if (action === '발송실패') {
        // 반송/발송실패: 진행 중이던 사이클이 있으면 실패로 표시하고,
        // 없으면(초기 발송부터 실패) 그 분기 셀을 실패로 새로 만든다.
        const year = date.getFullYear();
        const quarter = Math.floor(date.getMonth() / 3) + 1;
        if (cycle) {
          cycle.status = '발송실패';
        } else {
          cycle = { status: '발송실패', sentDate: date };
          yearsSet.add(year);
          if (!byYear[year]) byYear[year] = {};
          if (!byYear[year][corp]) byYear[year][corp] = {};
          byYear[year][corp][quarter] = cycle;
        }
      }
    });
  });

  const years = Array.from(yearsSet).sort((a, b) => b - a);
  const corpNames = corpRows.map(d => d.corpName);

  const byYearSerialized = {};
  years.forEach(year => {
    byYearSerialized[year] = {};
    corpNames.forEach(corp => {
      const quarters = {};
      for (let q = 1; q <= 4; q++) {
        const cell = ((byYear[year] || {})[corp] || {})[q];
        quarters[q] = cell ? { status: cell.status, sentDate: formatDate_(cell.sentDate) } : null;
      }
      byYearSerialized[year][corp] = quarters;
    });
  });

  return { years, corpNames, byYear: byYearSerialized };
}

function renderDashboardHtml_(data) {
  const fmt = v => v ? formatDate_(v) : '-';
  const statusPill = status => {
    const s = status || '대기';
    const cls = { '대기': 'wait', '발송완료': 'sent', '리마인드발송': 'remind', '회신완료': 'done', '발송실패': 'fail' }[s] || 'wait';
    return `<span class="pill ${cls}">${escapeHtml_(s)}</span>`;
  };
  const actionPill = action => {
    const cls = { '발송완료': 'sent', '리마인드발송': 'remind', '회신접수': 'done', '발송실패': 'fail' }[action] || 'wait';
    return `<span class="pill ${cls}">${escapeHtml_(action)}</span>`;
  };

  const responsePill = t => {
    if (!t) return '<span class="pill wait">미응답</span>';
    const cls = t === '변경있음' ? 'remind' : 'done';
    return `<span class="pill ${cls}">${escapeHtml_(t)}</span>`;
  };

  const detailRows = data.rows.map(d => `
    <tr>
      <td>${escapeHtml_(d.corpName)}</td>
      <td>${escapeHtml_(d.recipientName || '-')}</td>
      <td>${statusPill(d.status)}</td>
      <td>${responsePill(d.responseType)}</td>
      <td class="mono">${fmt(d.lastSendDate)}</td>
      <td class="mono">${fmt(d.nextSendDate)}</td>
      <td class="mono">${fmt(d.reminderDate)}</td>
      <td class="mono">${fmt(d.replyDate)}</td>
    </tr>`).join('');

  const logRows = data.logs.length ? data.logs.map(row => `
    <li>
      <div class="log-top"><span class="mono log-date">${escapeHtml_(formatDateTime_(row[0]))}</span>${actionPill(row[2])}</div>
      <div class="log-corp">${escapeHtml_(row[1])}</div>
      <div class="log-detail">${escapeHtml_(row[3] || '-')}</div>
    </li>`).join('') : `<li class="empty">아직 기록된 로그가 없습니다.</li>`;

  const tiles = [
    { key: '대기', label: '대기', cls: 'wait' },
    { key: '발송완료', label: '발송완료', cls: 'sent' },
    { key: '리마인드발송', label: '리마인드발송', cls: 'remind' },
    { key: '회신완료', label: '회신완료', cls: 'done' },
    { key: '발송실패', label: '발송실패', cls: 'fail' },
  ].map(t => `
    <div class="tile ${t.cls}">
      <span class="tile-v">${data.counts[t.key] || 0}</span>
      <span class="tile-l">${t.label}</span>
    </div>`).join('');

  const q = data.quarterly;
  const quarterCell = cell => {
    if (!cell) return `<td><span class="q-cell q-empty">미발송</span></td>`;
    const cls = { '발송완료': 'sent', '리마인드발송': 'remind', '회신완료': 'done', '발송실패': 'fail' }[cell.status] || 'wait';
    return `<td><span class="q-cell q-${cls}"><span class="q-status">${escapeHtml_(cell.status)}</span><span class="q-date mono">${escapeHtml_(cell.sentDate)}</span></span></td>`;
  };
  const yearOptions = q.years.map((y, i) => `<option value="${y}"${i === 0 ? ' selected' : ''}>${y}년</option>`).join('');
  const yearTables = q.years.map((y, i) => `
    <div class="q-year-table" data-year="${y}" style="${i === 0 ? '' : 'display:none;'}">
      <div class="table-scroll">
        <table>
          <thead><tr><th>법인명</th><th>1분기</th><th>2분기</th><th>3분기</th><th>4분기</th></tr></thead>
          <tbody>${
            q.corpNames.length
              ? q.corpNames.map(corp => `
                <tr><td>${escapeHtml_(corp)}</td>${[1, 2, 3, 4].map(qn => quarterCell(q.byYear[y][corp][qn])).join('')}</tr>`).join('')
              : `<tr><td colspan="5" class="empty">등록된 법인이 없습니다.</td></tr>`
          }</tbody>
        </table>
      </div>
    </div>`).join('');

  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Ctext y='.9em' font-size='90'%3E%F0%9F%93%91%3C/text%3E%3C/svg%3E">
<link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css">
<style>
  :root {
    --brand:#286fb7; --brand-light:#69b3e3;
    --page-bg:#f3f5fa; --card-bg:#ffffff; --ink:#1b2033; --muted:#6b7280; --line:#e6e9f2;
    --wait-bg:#eef0f5; --wait-fg:#6b7280;
    --sent-bg:#dcfce7; --sent-fg:#15803d;
    --remind-bg:#ffedd5; --remind-fg:#ea580c;
    --done-bg:#dbeafe; --done-fg:#1d4ed8;
    --fail-bg:#fee2e2; --fail-fg:#dc2626;
    --shadow:0 1px 2px rgba(16,24,40,.06), 0 1px 3px rgba(16,24,40,.08);
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --brand:#286fb7; --brand-light:#69b3e3;
      --page-bg:#10121b; --card-bg:#1a1d29; --ink:#e7e9f2; --muted:#8b90a3; --line:#2a2e3d;
      --wait-bg:#242736; --wait-fg:#9096ab;
      --sent-bg:#123a24; --sent-fg:#34d399;
      --remind-bg:#3a2410; --remind-fg:#f97316;
      --done-bg:#1e2a4a; --done-fg:#60a5fa;
      --fail-bg:#3e1414; --fail-fg:#ef4444;
      --shadow:0 1px 2px rgba(0,0,0,.4), 0 1px 3px rgba(0,0,0,.5);
    }
  }
  * { box-sizing:border-box; }
  body {
    margin:0; background:var(--page-bg); color:var(--ink);
    font-family:"Pretendard",-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic","Noto Sans KR",sans-serif;
    font-size:14px; line-height:1.6;
  }
  .mono { font-family:ui-monospace,"IBM Plex Mono","SFMono-Regular",Menlo,Consolas,monospace; font-variant-numeric:tabular-nums; }
  .banner { background:var(--brand); padding:20px 24px 30px; }
  .banner-title { max-width:1000px; margin:0 auto; text-align:center; }
  .banner-title h1 { font-size:24px; margin:0 0 6px; color:#fff; letter-spacing:-0.01em; }
  .banner-title .sub { font-size:12.5px; color:rgba(255,255,255,.72); }
  .wrap { max-width:1000px; margin:0 auto; padding:28px 20px 64px; }
  .tiles { display:grid; grid-template-columns:repeat(5,1fr); gap:10px; margin-bottom:32px; }
  @media (max-width:640px) { .tiles { grid-template-columns:repeat(2,1fr); } }
  .tile { background:var(--card-bg); box-shadow:var(--shadow); border-left:4px solid var(--muted); border-radius:10px; padding:14px 14px; display:flex; flex-direction:column; gap:4px; }
  .tile.wait { border-left-color:var(--wait-fg); }
  .tile.sent { border-left-color:var(--sent-fg); }
  .tile.remind { border-left-color:var(--remind-fg); }
  .tile.done { border-left-color:var(--done-fg); }
  .tile.fail { border-left-color:var(--fail-fg); }
  .tile-v { font-family:ui-monospace,"IBM Plex Mono","SFMono-Regular",Menlo,Consolas,monospace; font-variant-numeric:tabular-nums; font-size:26px; font-weight:700; }
  .tile-l { font-size:12px; color:var(--muted); white-space:nowrap; }
  section { margin-bottom:32px; }
  .section-head { display:flex; align-items:center; justify-content:space-between; gap:12px; margin-bottom:14px; flex-wrap:wrap; }
  h2 { font-size:18px; margin:0; font-weight:800; color:var(--ink); }
  select#yearSelect {
    font-family:ui-monospace,"IBM Plex Mono","SFMono-Regular",Menlo,Consolas,monospace;
    font-size:13px; font-weight:600; color:var(--brand); background:var(--card-bg);
    border:1px solid var(--brand); border-radius:8px; padding:7px 14px; cursor:pointer;
  }
  select#yearSelect:focus-visible { outline:2px solid var(--brand); outline-offset:2px; }
  .legend { display:flex; gap:14px; flex-wrap:wrap; font-size:12px; color:var(--muted); margin-bottom:12px; }
  .legend span { display:inline-flex; align-items:center; gap:5px; }
  .legend i { width:8px; height:8px; border-radius:50%; display:inline-block; }
  .legend i.wait { background:var(--wait-fg); }
  .legend i.sent { background:var(--sent-fg); }
  .legend i.remind { background:var(--remind-fg); }
  .legend i.done { background:var(--done-fg); }
  .legend i.fail { background:var(--fail-fg); }
  .table-scroll { overflow-x:auto; background:var(--card-bg); box-shadow:var(--shadow); border-radius:10px; }
  table { border-collapse:collapse; width:100%; min-width:640px; font-size:13px; }
  th, td { padding:11px 14px; text-align:center; border-bottom:1px solid var(--line); white-space:nowrap; }
  th { font-size:12px; text-transform:uppercase; letter-spacing:0.04em; color:var(--ink); font-weight:800; }
  tbody tr:hover td { background:var(--wait-bg); }
  tr:last-child td { border-bottom:none; }
  .pill { display:inline-flex; align-items:center; gap:5px; padding:4px 11px; border-radius:999px; font-size:12.5px; font-weight:700; }
  .pill::before { content:""; width:6px; height:6px; border-radius:50%; background:currentColor; }
  .pill.wait { background:var(--wait-bg); color:var(--wait-fg); }
  .pill.sent { background:var(--sent-bg); color:var(--sent-fg); }
  .pill.remind { background:var(--remind-bg); color:var(--remind-fg); }
  .pill.done { background:var(--done-bg); color:var(--done-fg); }
  .pill.fail { background:var(--fail-bg); color:var(--fail-fg); }
  .q-cell { display:inline-flex; flex-direction:column; gap:1px; padding:5px 10px; border-radius:8px; font-size:12px; font-weight:600; min-width:64px; }
  .q-cell .q-date { font-size:11px; font-weight:500; opacity:0.85; }
  .q-cell.q-wait { background:var(--wait-bg); color:var(--wait-fg); }
  .q-cell.q-sent { background:var(--sent-bg); color:var(--sent-fg); }
  .q-cell.q-remind { background:var(--remind-bg); color:var(--remind-fg); }
  .q-cell.q-done { background:var(--done-bg); color:var(--done-fg); }
  .q-cell.q-fail { background:var(--fail-bg); color:var(--fail-fg); }
  .q-cell.q-empty { background:transparent; color:var(--muted); font-weight:500; padding-left:0; }
  .empty { text-align:center; color:var(--muted); padding:20px; white-space:normal; }
  .split { display:grid; grid-template-columns:1fr 300px; gap:24px; align-items:start; }
  @media (max-width:820px) { .split { grid-template-columns:1fr; } }
  .side-col { display:flex; flex-direction:column; }
  .log-feed { list-style:none; margin:0; padding:0; display:flex; flex-direction:column; gap:8px; overflow-y:auto; }
  .log-feed li { background:var(--card-bg); box-shadow:var(--shadow); border-radius:10px; padding:10px 12px; }
  .log-feed li.empty { box-shadow:none; }
  .log-top { display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:4px; }
  .log-date { font-size:11px; color:var(--muted); }
  .log-corp { font-size:13px; font-weight:700; margin-bottom:2px; }
  .log-detail { font-size:12px; color:var(--muted); }
  footer { font-size:12px; color:var(--muted); text-align:center; margin-top:24px; }
</style>
</head>
<body>
  <div class="banner">
    <div class="banner-title">
      <h1>직원태양광 법인 등기부등본 변경사항 알림</h1>
      <div class="sub">케이씨에이치그룹 경영지원본부 인사총무팀 · 마지막 갱신 ${formatDateTime_(data.generatedAt)}</div>
    </div>
  </div>

<div class="wrap">
  <div class="tiles">${tiles}</div>

  <section>
    <div class="section-head">
      <h2>연도별 · 분기별 발송 현황</h2>
      <select id="yearSelect" onchange="showYear(this.value)">${yearOptions}</select>
    </div>
    <div class="legend">
      <span><i class="wait"></i>대기</span>
      <span><i class="sent"></i>발송완료</span>
      <span><i class="remind"></i>리마인드발송</span>
      <span><i class="done"></i>회신완료</span>
      <span><i class="fail"></i>발송실패</span>
    </div>
    ${yearTables}
  </section>

  <div class="split">
    <section class="main-col">
      <h2 style="margin-bottom:12px;">법인별 현재 상태 (${data.rows.length}개)</h2>
      <div class="table-scroll">
        <table>
          <thead><tr>
            <th>법인명</th><th>수신자</th><th>상태</th><th>응답</th><th>마지막발송일</th><th>다음발송예정일</th><th>리마인드발송일</th><th>회신일</th>
          </tr></thead>
          <tbody>${detailRows || `<tr><td colspan="8" class="empty">등록된 법인이 없습니다.</td></tr>`}</tbody>
        </table>
      </div>
    </section>

    <aside class="side-col">
      <h2 style="margin-bottom:12px;">최근 로그</h2>
      <ul class="log-feed">${logRows}</ul>
    </aside>
  </div>

  <footer>이 페이지는 법인목록 · 로그 시트를 그대로 읽어 보여줍니다. 데이터를 바꾸려면 시트에서 수정하세요.</footer>
</div>
<script>
  function showYear(year) {
    document.querySelectorAll('.q-year-table').forEach(function (el) {
      el.style.display = el.dataset.year === year ? '' : 'none';
    });
  }
  function syncLogHeight() {
    var main = document.querySelector('.main-col');
    var feed = document.querySelector('.log-feed');
    if (!main || !feed) return;
    feed.style.maxHeight = main.offsetHeight + 'px';
  }
  window.addEventListener('load', syncLogHeight);
  window.addEventListener('resize', syncLogHeight);
</script>
</body>
</html>`;
}

function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function formatDateTime_(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}
