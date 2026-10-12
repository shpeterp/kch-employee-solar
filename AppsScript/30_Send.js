// 메일 발송: 처음 메일(GmailApp), 리마인드(Gmail API로 같은 스레드)

/**
 * 처음 메일 발송. token은 호출하는 쪽(processRow_)이 보내기 전에 시트에 먼저 저장한 값.
 * 같은 토큰이 메일 본문(버튼 링크)에 들어가므로, 기록이 끊겨도 보낸편지함에서 이 메일을 찾아 복구할 수 있음.
 */
function sendInitialRequest_(data, regFile, token) {
  if (!isValidEmail_(data.recipientEmail)) {
    throw new Error(`수신자 이메일 형식 오류: "${data.recipientEmail}"`);
  }
  if (data.cc && !isValidEmail_(data.cc)) {
    throw new Error(`참조 이메일 형식 오류: "${data.cc}"`);
  }
  const subject = `[직원태양광] ${data.corpName} 법인 등기부등본 내용 변경 사항 알림 요청`;
  token = token || data.responseToken || Utilities.getUuid();
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
