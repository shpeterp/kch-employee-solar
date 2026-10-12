// 매일 자동 실행: 잠금 → 법인별 처리(발송/회신확인/리마인드/무응답보고) → 결과 검증 → 요약

function mainDaily() {
  // [중복 실행 방지] 중복 트리거 등으로 mainDaily가 동시에 두 번 실행되면 이중 발송/이중 기록이 생기므로,
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
  const startedAt = Date.now();
  // [중단] 서명·연락처 값이 비어 있으면 빈 서명으로 메일이 나가지 않도록 아무것도 보내지 않고 본인에게만 알림
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

    // [이어서 처리] Apps Script 6분 제한 전에 스스로 멈춤. 처리 상태는 행마다 시트에 남아 있으므로
    // 남은 법인은 다음 실행 때 상태 그대로 이어서 처리됨(이미 처리한 법인은 상태가 바뀌어 다시 보내지 않음).
    if (Date.now() - startedAt > CONFIG.MAX_RUN_MS) {
      const left = values.slice(r).filter(v => v[HEADERS.indexOf('법인명')]).length;
      actions.push(`[시간초과-이어서] ${left}곳은 다음 실행 때 이어서 처리 (${data.corpName}부터)`);
      appendLog_('(전체)', '시간초과-이어서', `${data.corpName}부터 ${left}곳 다음 실행으로`);
      break;
    }

    const rowNum = r + 1;
    // [법인별 격리] 한 법인 처리 중 오류가 나도 나머지 법인은 계속 처리되도록 행별로 격리.
    try {
      processRow_(sheet, rowNum, data, today, actions, escalated);
    } catch (e) {
      actions.push(`[처리오류] ${data.corpName} - ${e}`);
      appendLog_(data.corpName, '처리오류', String(e));
    }
  }

  // [결과 확인] 처리 후 시트를 다시 읽어, 발송중으로 남은 행·처리 안 된 변경확인 건을 요약에 올림
  verifyRunResult_(sheet, actions, today);

  if (actions.length > 0) sendSummaryMail_(actions);
  // [사람에게 넘김] 리마인드에도 반응 없는 곳이 새로 생겼으면 관리자에게 요약 알림 (하루 뒤)
  if (escalated.length > 0) notifyEscalation_(escalated, allData);
  // 다음 분기 자동발송이 임박했으면 본인에게 사전 알림 (직원 발송 아님)
  notifyUpcomingSend_(allData, today);
}

/** 시트의 한 칸에 값 쓰기 (칸 이름으로) */
function setCell_(sheet, rowNum, header, value) {
  sheet.getRange(rowNum, HEADERS.indexOf(header) + 1).setValue(value);
}

/** 법인목록 한 행(한 법인)에 대한 발송/회신확인/리마인드/무응답보고 처리. mainDailyBody에서 행별 try로 감싸 호출. */
function processRow_(sheet, rowNum, data, today, actions, escalated) {
    // 새 분기가 도래했는데(다음발송예정일 도달) 지난 분기 기록이 남아 있으면 비우고 '대기'로 되돌린다.
    // 회신완료뿐 아니라 지난 분기 무응답(발송완료·리마인드발송)도 초기화해서 이번 분기 메일이 다시 나가게 함.
    // (발송실패는 이메일 주소를 고쳐야 하므로 초기화하지 않음 → 사전알림에 '발송되지 않는 곳'으로 표시됨)
    if (QUARTER_RESET_STATUSES.indexOf(data.status) >= 0 && data.nextSendDate && stripTime_(data.nextSendDate) <= today) {
      const prevStatus = data.status;
      ['상태', 'ThreadID', '리마인드발송일', '회신일', '응답토큰', '응답유형', '응답일시', '변경내용', '무응답보고일']
        .forEach(h => setCell_(sheet, rowNum, h, ''));
      data.status = '';
      data.threadId = '';
      data.responseToken = '';
      const note = prevStatus === '회신완료' ? '' : ` (지난 분기 무응답: 상태=${prevStatus})`;
      appendLog_(data.corpName, '분기초기화', `${formatDate_(data.nextSendDate)} 분기 발송 대상으로 초기화${note}`);
    }

    // [보류] 변경확인필요가 처리되지 않은 채 다음 분기가 되면 새 메일을 보내지 않고 담당자에게 계속 알림
    if (data.status === STATUS_CHANGE_CHECK) {
      if (data.nextSendDate && stripTime_(data.nextSendDate) <= today) {
        actions.push(`[변경확인 미처리-발송보류] ${data.corpName} - 변경확인을 처리해야 이번 분기 메일이 나갑니다`);
      }
      return;
    }

    // [복구] 지난 실행이 메일을 보낸 직후 끊겨 '발송중'으로 남은 행: 보낸편지함과 대조해서 기록을 맞춤
    if (data.status === STATUS_SENDING) {
      const recovered = reconcileSending_(sheet, rowNum, data, today, actions);
      if (recovered) return;
      data.status = '';   // 실제로 나간 메일이 없으면 대기로 되돌려 아래에서 정상 발송
    }

    const status = data.status;

    if (!status || status === '') {
      if (data.nextSendDate && stripTime_(data.nextSendDate) <= today) {
        // [첨부 확인 → 보류] 등기부등본이 없으면 이 법인만 보류(상태를 바꾸지 않으므로 파일을 넣으면 다음 실행 때 발송됨)
        const regFile = findRegistryFile_(data.corpName);
        if (!regFile) {
          actions.push(`[첨부없음-보류] ${data.corpName} - 등기부등본 파일이 없어 처음 메일을 보내지 않음`);
          appendLog_(data.corpName, '첨부없음-보류', '처음 메일 보류 (폴더에 파일을 넣으면 다음 실행 때 발송)');
          return;
        }

        // [중복 발송 방지] 보내기 전에 '발송중'과 응답토큰을 먼저 저장. 보낸 뒤 기록하다 끊겨도
        // 다음 실행이 이 표시를 보고 보낸편지함을 확인하므로 같은 메일이 두 번 나가지 않음.
        const token = Utilities.getUuid();
        setCell_(sheet, rowNum, '상태', STATUS_SENDING);
        setCell_(sheet, rowNum, '응답토큰', token);
        setCell_(sheet, rowNum, '마지막발송일', today);
        setCell_(sheet, rowNum, '응답유형', '');   // 지난 응답이 남아 있으면 새 발송이 바로 회신완료로 잡히지 않도록 비움
        setCell_(sheet, rowNum, '응답일시', '');
        SpreadsheetApp.flush();
        data.responseToken = token;

        let sent;
        try {
          sent = sendInitialRequest_(data, regFile, token);
        } catch (e) {
          // 이메일 형식 오류 등 발송 자체가 실패한 경우만 '발송실패'로 기록
          setCell_(sheet, rowNum, '상태', '발송실패');
          actions.push(`[발송실패] ${data.corpName} (${data.recipientEmail}) - ${e}`);
          appendLog_(data.corpName, '발송실패', `${data.recipientEmail} / ${e}`);
          return;
        }
        // 여기부터는 메일이 이미 나간 상태. 기록 중 오류가 나도 '발송실패'로 적지 않음(상태는 발송중으로 남아 다음 실행 때 복구)
        recordInitialSent_(sheet, rowNum, data, today, sent.threadId);
        actions.push(`[발송완료] ${data.corpName} (${data.recipientEmail})`);
        appendLog_(data.corpName, '발송완료', `수신: ${data.recipientEmail}`);
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
        setCell_(sheet, rowNum, '상태', '발송실패');
        actions.push(`[발송실패-반송] ${data.corpName} (${data.recipientEmail})`);
        appendLog_(data.corpName, '발송실패', '반송(주소 오류 등) - 실제 전달되지 않음');
        return;
      }

      if (incoming.hasReply) {
        // [회신 내용 검증] 답장이 왔다고 바로 닫지 않고, 본문이 '변경 없음'으로 확실할 때만 회신완료.
        // 변경 언급·첨부·애매한 답장은 '변경확인필요'로 돌려 담당자에게 넘김.
        const verdict = classifyReplies_(incoming.replies);
        setCell_(sheet, rowNum, '회신일', today);
        if (verdict.type === '변동없음') {
          setCell_(sheet, rowNum, '상태', '회신완료');
          setCell_(sheet, rowNum, '응답유형', '변동없음(메일)');
          setCell_(sheet, rowNum, '응답일시', new Date());
          actions.push(`[회신접수] ${data.corpName} - 메일 회신: 변경 없음`);
          appendLog_(data.corpName, '회신접수', '메일회신: 변동없음');
        } else {
          setCell_(sheet, rowNum, '상태', STATUS_CHANGE_CHECK);
          setCell_(sheet, rowNum, '응답유형', '확인필요');
          setCell_(sheet, rowNum, '응답일시', new Date());
          setCell_(sheet, rowNum, '변경내용', verdict.excerpt.slice(0, 2000));
          notifyChangeReported_(data, verdict);
          actions.push(`[변경확인필요] ${data.corpName} - ${verdict.reason}`);
          appendLog_(data.corpName, '변경확인필요', `메일회신 확인 필요: ${verdict.reason}`);
        }
        return;
      }

      if (status === '발송완료') {
        const daysSince = Math.floor((today - stripTime_(data.lastSendDate)) / (1000 * 60 * 60 * 24));
        if (daysSince >= CONFIG.REMIND_AFTER_DAYS) {
          // [중복 리마인드 방지] 스레드에 내가 보낸 메일이 이미 2통 이상이면 리마인드는 나간 것 → 기록만 맞춤
          if (countMySentMessages_(thread) >= 2) {
            setCell_(sheet, rowNum, '상태', '리마인드발송');
            setCell_(sheet, rowNum, '리마인드발송일', data.reminderDate || today);
            actions.push(`[복구] ${data.corpName} - 리마인드는 이미 발송됨, 기록만 맞춤(다시 보내지 않음)`);
            appendLog_(data.corpName, '리마인드기록복구', '스레드에 리마인드가 이미 있어 재발송하지 않음');
            return;
          }
          // [첨부 확인 → 보류] 등기부등본이 없으면 이 법인 리마인드만 보류(상태 유지 → 파일을 넣으면 다음 실행 때 발송됨)
          const regFile = findRegistryFile_(data.corpName);
          if (!regFile) {
            actions.push(`[첨부없음-보류] ${data.corpName} - 등기부등본 파일이 없어 리마인드를 보내지 않음`);
            appendLog_(data.corpName, '첨부없음-보류', '리마인드 보류 (폴더에 파일을 넣으면 다음 실행 때 발송)');
            return;
          }
          // 옛 코드로 최초발송된 행은 응답토큰이 없어 리마인드 버튼 링크가 빈 토큰이 됨 → 없으면 생성해 기록.
          if (!data.responseToken) {
            data.responseToken = Utilities.getUuid();
            setCell_(sheet, rowNum, '응답토큰', data.responseToken);
          }
          sendReminder_(thread, data, regFile);
          setCell_(sheet, rowNum, '상태', '리마인드발송');
          setCell_(sheet, rowNum, '리마인드발송일', today);
          actions.push(`[리마인드발송] ${data.corpName} (${data.recipientEmail})`);
          appendLog_(data.corpName, '리마인드발송', '');
        }
      } else if (status === '리마인드발송') {
        // [사람에게 넘김] 리마인드까지 보냈는데도 응답(버튼/회신)이 없는 경우: 리마인드 후 N일 지나면 관리자에게 1회 보고.
        // (직원에게는 아무것도 안 감. 무응답보고일이 이미 있으면 재보고하지 않음.)
        const daysSinceRemind = data.reminderDate
          ? Math.floor((today - stripTime_(data.reminderDate)) / (1000 * 60 * 60 * 24)) : 0;
        if (!data.escalatedDate && data.reminderDate && daysSinceRemind >= CONFIG.ESCALATE_AFTER_DAYS) {
          setCell_(sheet, rowNum, '무응답보고일', today);
          appendLog_(data.corpName, '리마인드무응답', `리마인드 ${formatDate_(data.reminderDate)} 후 ${daysSinceRemind}일 무응답`);
          escalated.push(data);
        }
      }
    }
}

/** 처음 메일 발송 후 시트 기록. 상태는 맨 마지막에 바꿔서, 중간에 끊기면 '발송중'이 남도록 함. */
function recordInitialSent_(sheet, rowNum, data, today, threadId) {
  const targetMonth = addMonths_(today, data.cycleMonths || CONFIG.DEFAULT_CYCLE_MONTHS);
  const nextDate = firstMondayOfMonth_(targetMonth.getFullYear(), targetMonth.getMonth());
  setCell_(sheet, rowNum, '마지막발송일', today);
  setCell_(sheet, rowNum, '다음발송예정일', nextDate);
  setCell_(sheet, rowNum, 'ThreadID', threadId);
  ['리마인드발송일', '회신일', '무응답보고일', '변경내용'].forEach(h => setCell_(sheet, rowNum, h, ''));
  // 발송중에 [변경사항 없음] 버튼이 먼저 눌렸으면(응답유형 있음) 바로 회신완료, 아니면 발송완료
  const responded = sheet.getRange(rowNum, HEADERS.indexOf('응답유형') + 1).getValue();
  setCell_(sheet, rowNum, '상태', responded ? '회신완료' : '발송완료');
}

/**
 * '발송중'으로 남은 행 복구. 보낸편지함에서 이 법인 앞으로 마지막발송일 이후 보낸 메일 중
 * 이 행의 응답토큰이 들어 있는 메일을 찾으면 → 이미 나간 것이므로 기록만 채우고 true.
 * 못 찾으면 → 실제로 안 나간 것이므로 false (호출한 쪽에서 정상 발송).
 */
function reconcileSending_(sheet, rowNum, data, today, actions) {
  const since = data.lastSendDate ? new Date(data.lastSendDate) : addMonths_(today, -1);
  since.setDate(since.getDate() - 1);
  const query = `in:sent to:${data.recipientEmail} after:${Utilities.formatDate(since, 'Asia/Seoul', 'yyyy/MM/dd')}`;
  const threads = GmailApp.search(query, 0, 20);
  for (const t of threads) {
    const hit = t.getMessages().find(m => !m.isDraft() && data.responseToken && m.getBody().indexOf(data.responseToken) >= 0);
    if (hit) {
      recordInitialSent_(sheet, rowNum, data, stripTime_(hit.getDate()), t.getId());
      actions.push(`[복구] ${data.corpName} - 지난 실행에서 이미 발송된 메일을 확인해 기록만 채움(다시 보내지 않음)`);
      appendLog_(data.corpName, '발송완료', `발송기록 복구(보낸편지함 대조): ${data.recipientEmail}`);
      return true;
    }
  }
  appendLog_(data.corpName, '발송중-미발송확인', '보낸편지함에 해당 메일 없음 → 대기로 되돌려 다시 발송');
  setCell_(sheet, rowNum, '상태', '');
  return false;
}

/** 실행이 끝난 뒤 시트를 다시 읽어, 사람이 봐야 할 상태가 남아 있는지 확인해 요약에 추가 */
function verifyRunResult_(sheet, actions, today) {
  const values = sheet.getDataRange().getValues();
  const stuck = [];
  const pendingChange = [];
  for (let r = 1; r < values.length; r++) {
    const d = rowToData_(values[r]);
    if (!d.corpName) continue;
    if (d.status === STATUS_SENDING) stuck.push(d.corpName);
    if (d.status === STATUS_CHANGE_CHECK) pendingChange.push(`${d.corpName}(회신 ${d.replyDate ? formatDate_(d.replyDate) : '-'})`);
  }
  if (stuck.length) actions.push(`[확인필요] 발송중으로 남은 곳 ${stuck.length}곳: ${stuck.join(', ')} - 다음 실행 때 보낸편지함과 대조해 복구`);
  // 변경확인 대기 건은 다른 처리 내역이 있는 날 요약에 함께 표시 (이것만으로 매일 메일을 보내지는 않음)
  if (pendingChange.length && actions.length) actions.push(`[변경확인 대기] ${pendingChange.length}곳: ${pendingChange.join(', ')}`);
}
