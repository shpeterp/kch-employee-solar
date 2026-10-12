// 회신 판별: 반송/실제 회신 구분 + 회신 본문으로 '변경 없음' / '확인 필요' 판정

/** 스레드의 메시지를 훑어서 '실제 회신'과 '반송(bounce)'을 구분. 실제 회신 메시지 목록도 함께 돌려줌. */
function classifyIncoming_(thread) {
  const myEmail = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const messages = thread.getMessages();
  let hasBounce = false;
  const replies = [];
  for (const m of messages) {
    if (m.isDraft()) continue; // 리마인드_초안_만들기로 만든 초안은 회신이 아님
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
    replies.push(m);
  }
  return { hasBounce, hasReply: replies.length > 0, replies };
}

/** 스레드에서 내가 보낸(초안 제외) 메시지 수. 리마인드가 이미 나갔는지 확인할 때 씀. */
function countMySentMessages_(thread) {
  const myEmail = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return thread.getMessages().filter(m => !m.isDraft() && (m.getFrom() || '').toLowerCase().indexOf(myEmail) >= 0).length;
}

/**
 * 회신 본문에서 인용된 원문(이전 메일)을 잘라내고 새로 쓴 부분만 남김.
 * Gmail/Outlook/네이버 등에서 흔한 인용 시작 표시를 기준으로 자름.
 */
function extractReplyText_(plainBody) {
  const lines = String(plainBody || '').replace(/\r\n/g, '\n').split('\n');
  const quoteStart = [
    /^\s*>/,                                   // > 인용
    /^-{2,}\s*(original message|원본 메시지|원래 메시지)/i,
    /^\s*(from|보낸 사람|보낸사람)\s*:/i,          // Outlook·네이버 인용 머리
    /^\s*on .+wrote:\s*$/i,                    // Gmail 영문
    /님이 작성:\s*$/,                           // Gmail 한글
    /^\s*20\d\d[.\-\/년]\s*\d{1,2}.*(작성|wrote)/,   // 날짜 + 작성
  ];
  const kept = [];
  for (const line of lines) {
    if (quoteStart.some(re => re.test(line))) break;
    kept.push(line);
  }
  return kept.join('\n').trim();
}

/**
 * 회신 본문만 보고 판정.
 *  - '변동없음' : "변경사항 없습니다"처럼 변동 없음이 분명하고, 그 밖에 변경을 뜻하는 말이 없을 때만
 *  - '확인필요' : 그 외 전부 (변경 언급, 첨부파일, 애매한 답장, 빈 본문 등)
 * 애매하면 사람이 보도록 '확인필요' 쪽으로 보냄 (변경을 놓치면 과태료 위험이 있기 때문).
 */
function classifyReplyText_(text, hasAttachment) {
  const compact = String(text || '').replace(/\s+/g, '');
  if (!compact) return { type: '확인필요', reason: '본문 없음' };
  if (hasAttachment) return { type: '확인필요', reason: '첨부파일 있음' };
  const noChange = /(변경|변동)(사항|내용|된것|된사항|된부분)?(이|은|는|도)?(전혀)?(없|무)|이상(이)?없|동일합니다|동일함|그대로입니다/g;
  const changeWord = /변경|변동|이전|바뀌|바꼈|바뀐|사임|취임|선임|퇴임|중임|해임|이사(를|했|함|갔|가)/;
  if (!new RegExp(noChange.source).test(compact)) return { type: '확인필요', reason: '변경 없음이라는 표현이 없음' };
  const rest = compact.replace(noChange, '');
  if (changeWord.test(rest)) return { type: '확인필요', reason: '변경 관련 표현이 함께 있음' };
  return { type: '변동없음', reason: '변경 없음으로 회신' };
}

/** 여러 회신을 한꺼번에 판정. 하나라도 '확인필요'면 확인필요. 담당자 알림용 발췌(excerpt)도 만듦. */
function classifyReplies_(replies) {
  let verdict = { type: '변동없음', reason: '' };
  const excerpts = [];
  replies.forEach(m => {
    const text = extractReplyText_(m.getPlainBody());
    const hasAttachment = m.getAttachments({ includeInlineImages: false }).length > 0;
    const v = classifyReplyText_(text, hasAttachment);
    excerpts.push(`[${formatDateTime_(m.getDate())} ${m.getFrom()}] ${text.slice(0, 300)}${hasAttachment ? ' (첨부파일 있음)' : ''}`);
    if (v.type === '확인필요') verdict = v;
    else if (!verdict.reason) verdict = v;
  });
  return { type: verdict.type, reason: verdict.reason, excerpt: excerpts.join('\n') };
}
