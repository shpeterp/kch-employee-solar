// 담당자 알림: 처리 요약, 무응답 보고, 분기 사전알림, 버튼 응답 알림

function sendSummaryMail_(actions) {
  const issueTags = ['[발송실패', '[첨부없음', '[처리오류', '[변경확인', '[확인필요', '[시간초과'];
  const isIssue = a => issueTags.some(t => a.indexOf(t) === 0);
  const fails = actions.filter(isIssue);
  const oks = actions.filter(a => !isIssue(a));

  let body = `오늘(${formatDate_(new Date())}) 법인 등기부등본 확인 요청 자동화 처리 내역입니다.\n\n`;
  if (fails.length > 0) {
    body += `⚠️ 확인이 필요한 건 ${fails.length}건 (발송 실패·반송·첨부 없음·변경확인·처리 오류 등)\n`;
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

/** '변경사항 없음' 응답이 들어오면 담당자에게 알림 메일 발송 */
function notifyResponse_(d) {
  GmailApp.sendEmail(
    CONFIG.담당자이메일 || CONFIG.SUMMARY_TO,
    `[응답:변동없음] ${d.corpName}`,
    `${d.corpName} 법인이 등기부등본 '변경사항 없음'으로 확인 응답했습니다.\n수신자: ${d.recipientName} ${d.recipientTitle} (${d.recipientEmail})\n접수시각: ${formatDateTime_(new Date())}`
  );
}

/**
 * 회신 본문이 '변경 없음'으로 확실하지 않은 법인을 담당자에게 바로 알림.
 * 담당자가 내용을 보고 등기 반영 여부를 확인한 뒤 대시보드 [처리완료]로 닫음.
 */
function notifyChangeReported_(d, verdict) {
  let body = `${d.corpName} 법인에서 등기부등본 확인 메일에 회신했는데, '변경 없음'으로 확실하지 않아 확인이 필요합니다.\n`;
  body += `판단 이유: ${verdict.reason}\n\n`;
  body += `■ 회신 내용 (인용 부분 제외, 앞부분만)\n${verdict.excerpt}\n\n`;
  body += `■ 처리 순서\n`;
  body += `1. 회신 메일을 열어 변경 내용과 변경 날짜를 확인합니다.\n`;
  body += `2. 변경이 있으면 등기 반영 여부를 확인하고, 반영되면 새 등기부등본을 발급해 드라이브 폴더의 '등기부등본_${d.corpName}.pdf'를 교체합니다.\n`;
  body += `3. 대시보드에서 이 법인의 [처리완료]를 누릅니다. 처리 전까지 이 법인의 다음 분기 메일은 보류됩니다.\n\n`;
  body += `수신자: ${d.recipientName} ${d.recipientTitle} (${d.recipientEmail})`;
  GmailApp.sendEmail(
    CONFIG.담당자이메일 || CONFIG.SUMMARY_TO,
    `⚠️[변경확인필요] ${d.corpName} - 회신 내용 확인 (${formatDate_(new Date())})`,
    body
  );
}
