// 웹앱: [변경사항 없음] 버튼 응답 화면과 제출 처리

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
  // [접근 제어] 대시보드는 사람별 키가 있는 주소로만 열림 (수신자가 토큰만 지워도 내부 데이터가 안 보이도록).
  // 키 확인·역할 판정·데이터 거르기는 모두 서버(이 스크립트)에서 함. 화면에서 버튼만 숨기는 방식이 아님.
  const user = authDashboard_(p.key);
  if (!user) {
    if (p.key) appendLog_('(대시보드)', '접근거부', '등록되지 않은 키로 접속 시도');
    return HtmlService.createHtmlOutput(responseShell_(
      '접근 권한 필요',
      '<p>이 페이지는 접근 권한이 필요합니다.</p><p class="muted">담당자에게 정식 주소를 문의해 주세요.</p>',
      ''
    )).setTitle('접근 권한 필요');
  }
  appendLog_('(대시보드)', '대시보드열람', `${user.name} (${user.role})`);
  const data = getDashboardData_(user);
  const html = renderDashboardHtml_(data, user, p.key);
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

    // 이미 메일 회신으로 '변경확인필요'인 법인은 버튼만으로 닫지 않음(담당자 확인이 우선). 응답 기록만 남김.
    if (d.status === STATUS_CHANGE_CHECK) {
      appendLog_(d.corpName, '버튼응답', '변동없음 버튼 (변경확인필요 상태라 담당자 확인 대기 유지)');
      return { corp: escapeHtml_(d.corpName), msg: '응답이 접수되었습니다. 담당자가 확인 후 연락드리겠습니다.' };
    }
    // 발송 직후 기록이 끝나기 전(발송중)에 눌린 경우: 상태는 다음 실행의 복구 단계가 회신완료로 맞춤
    if (d.status !== STATUS_SENDING) sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('회신완료');
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
