// 웹앱: 담당자용 현황 대시보드

/** 대시보드 비밀키. "프로젝트 설정 > 스크립트 속성"의 DASHBOARD_KEY 값 (없으면 빈 값). */
function dashboardKey_() {
  return PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY') || '';
}

/**
 * 대시보드 사용자 목록. 스크립트 속성 DASHBOARD_USERS(JSON 배열)에서 읽고,
 * 예전 방식 DASHBOARD_KEY가 있으면 '관리' 권한 1명으로 함께 취급.
 * 역할: '관리' = 전체 정보 + 변경확인 [처리완료] / '조회' = 상태만(수신자 이름 가림, 회신 내용·로그 상세 숨김)
 */
function dashboardUsers_() {
  const users = [];
  const raw = PropertiesService.getScriptProperties().getProperty('DASHBOARD_USERS');
  if (raw) {
    try {
      JSON.parse(raw).forEach(u => {
        if (u && u.key && String(u.key).length >= 16) {
          users.push({ key: String(u.key), name: String(u.name || '이름없음'), role: u.role === '관리' ? '관리' : '조회' });
        }
      });
    } catch (e) {
      Logger.log('DASHBOARD_USERS 형식 오류: ' + e);
    }
  }
  const legacy = dashboardKey_();
  if (legacy && !users.some(u => u.key === legacy)) users.push({ key: legacy, name: CONFIG.담당자성함 || '담당자', role: '관리' });
  return users;
}

/** 키로 사용자 찾기. 없으면 null. (키는 16자 이상만 인정) */
function authDashboard_(key) {
  if (!key) return null;
  return dashboardUsers_().find(u => u.key === String(key)) || null;
}

/** '조회' 권한용 이름 가림: 홍길동 → 홍OO */
function maskName_(name) {
  const s = String(name || '');
  return s ? s.charAt(0) + 'O'.repeat(Math.max(1, s.length - 1)) : '-';
}

/**
 * [관리 권한 전용] 변경확인필요 법인을 담당자가 확인한 뒤 닫음. 대시보드 [처리완료] 버튼이 호출.
 * 권한 확인은 서버에서 다시 함(화면 버튼을 숨기는 것만으로는 막을 수 없으므로).
 */
function resolveChange(key, corpName) {
  const user = authDashboard_(key);
  if (!user || user.role !== '관리') {
    appendLog_(String(corpName || ''), '접근거부', '처리완료 시도 - 관리 권한 없음');
    throw new Error('처리 권한이 없습니다.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    for (let r = 1; r < values.length; r++) {
      const d = rowToData_(values[r]);
      if (d.corpName !== corpName) continue;
      if (d.status !== STATUS_CHANGE_CHECK) throw new Error('이미 처리되었거나 변경확인 대상이 아닙니다.');
      const rowNum = r + 1;
      sheet.getRange(rowNum, HEADERS.indexOf('상태') + 1).setValue('회신완료');
      sheet.getRange(rowNum, HEADERS.indexOf('응답유형') + 1).setValue('변경처리완료');
      sheet.getRange(rowNum, HEADERS.indexOf('응답일시') + 1).setValue(new Date());
      appendLog_(d.corpName, '변경처리완료', `${user.name} 확인 후 처리완료`);
      return { corp: escapeHtml_(d.corpName), msg: '처리완료로 바꿨습니다. 등기 변경이 반영됐다면 드라이브의 등기부등본 파일도 새 파일로 교체해 주세요.' };
    }
    throw new Error('법인을 찾지 못했습니다.');
  } finally {
    lock.releaseLock();
  }
}

/** 대시보드 데이터. 역할에 따라 서버에서 미리 걸러서 보냄('조회'는 이름 가림, 회신 내용·로그 상세 제외). */
function getDashboardData_(user) {
  const isAdmin = user && user.role === '관리';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const data = rowToData_(values[r]);
    if (!data.corpName) continue;
    if (!isAdmin) {
      data.recipientName = maskName_(data.recipientName);
      data.recipientEmail = '';
      data.changeDetail = '';
    }
    rows.push(data);
  }

  const counts = { '대기': 0, '발송완료': 0, '리마인드발송': 0, '변경확인필요': 0, '회신완료': 0, '발송실패': 0 };
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
  // 대시보드 열람/거부 기록은 보안 기록이라 화면 피드에는 빼고, '조회' 권한에는 상세 내용을 숨김
  const logs = allLogs.slice().reverse()
    .filter(row => row[2] !== '대시보드열람' && row[2] !== '접근거부')
    .slice(0, 50)
    .map(row => isAdmin ? row : [row[0], row[1], row[2], '']);
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
      } else if ((action === '회신접수' || action === '변경처리완료') && cycle) {
        cycle.status = '회신완료';
      } else if (action === '변경확인필요' && cycle) {
        cycle.status = '변경확인필요';
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

function renderDashboardHtml_(data, user, key) {
  user = user || { name: '-', role: '조회' };
  const isAdmin = user.role === '관리';
  const fmt = v => v ? formatDate_(v) : '-';
  const STATUS_CLS = { '대기': 'wait', '발송중': 'wait', '발송완료': 'sent', '리마인드발송': 'remind', '변경확인필요': 'check', '회신완료': 'done', '발송실패': 'fail' };
  const statusPill = status => {
    const s = status || '대기';
    const cls = STATUS_CLS[s] || 'wait';
    return `<span class="pill ${cls}">${escapeHtml_(s)}</span>`;
  };
  const actionPill = action => {
    const cls = { '발송완료': 'sent', '리마인드발송': 'remind', '회신접수': 'done', '변경처리완료': 'done', '변경확인필요': 'check', '발송실패': 'fail' }[action] || 'wait';
    return `<span class="pill ${cls}">${escapeHtml_(action)}</span>`;
  };

  const responsePill = t => {
    if (!t) return '<span class="pill wait">미응답</span>';
    const cls = t === '확인필요' ? 'check' : 'done';
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
    { key: '변경확인필요', label: '변경확인필요', cls: 'check' },
    { key: '회신완료', label: '회신완료', cls: 'done' },
    { key: '발송실패', label: '발송실패', cls: 'fail' },
  ].map(t => `
    <div class="tile ${t.cls}">
      <span class="tile-v">${data.counts[t.key] || 0}</span>
      <span class="tile-l">${t.label}</span>
    </div>`).join('');

  // 변경확인필요 목록: '관리'는 회신 내용과 [처리완료] 버튼, '조회'는 법인명만
  const pendingChange = data.rows.filter(d => d.status === '변경확인필요');
  const changeSection = pendingChange.length ? `
  <section>
    <h2 style="margin-bottom:12px;">변경확인 필요 (${pendingChange.length}곳)</h2>
    <div class="change-list">${pendingChange.map(d => `
      <div class="change-card" id="chg-${escapeHtml_(d.corpName)}">
        <div class="change-top"><b>${escapeHtml_(d.corpName)}</b><span class="mono log-date">회신 ${fmt(d.replyDate)}</span></div>
        ${isAdmin
          ? `<pre class="change-text">${escapeHtml_(d.changeDetail || '(회신 내용 없음)')}</pre>
             <button class="resolve-btn" onclick="resolve(this, ${escapeHtml_(JSON.stringify(d.corpName))})">확인했음 · 처리완료</button>`
          : '<div class="log-detail">담당자가 회신 내용을 확인하고 있습니다.</div>'}
      </div>`).join('')}
    </div>
  </section>` : '';

  const q = data.quarterly;
  const quarterCell = cell => {
    if (!cell) return `<td><span class="q-cell q-empty">미발송</span></td>`;
    const cls = STATUS_CLS[cell.status] || 'wait';
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
    --check-bg:#ede9fe; --check-fg:#7c3aed;
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
      --check-bg:#2e1f4d; --check-fg:#a78bfa;
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
  .tiles { display:grid; grid-template-columns:repeat(6,1fr); gap:10px; margin-bottom:32px; }
  @media (max-width:640px) { .tiles { grid-template-columns:repeat(2,1fr); } }
  .tile { background:var(--card-bg); box-shadow:var(--shadow); border-left:4px solid var(--muted); border-radius:10px; padding:14px 14px; display:flex; flex-direction:column; gap:4px; }
  .tile.wait { border-left-color:var(--wait-fg); }
  .tile.sent { border-left-color:var(--sent-fg); }
  .tile.remind { border-left-color:var(--remind-fg); }
  .tile.done { border-left-color:var(--done-fg); }
  .tile.fail { border-left-color:var(--fail-fg); }
  .tile.check { border-left-color:var(--check-fg); }
  .pill.check { background:var(--check-bg); color:var(--check-fg); }
  .q-cell.q-check { background:var(--check-bg); color:var(--check-fg); }
  .legend i.check { background:var(--check-fg); }
  .change-list { display:flex; flex-direction:column; gap:10px; }
  .change-card { background:var(--card-bg); box-shadow:var(--shadow); border-left:4px solid var(--check-fg); border-radius:10px; padding:12px 14px; }
  .change-top { display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; }
  .change-text { white-space:pre-wrap; font-family:inherit; font-size:12.5px; color:var(--ink); background:var(--wait-bg); border-radius:8px; padding:10px; margin:0 0 10px; max-height:180px; overflow:auto; }
  .resolve-btn { border:none; border-radius:8px; background:var(--check-fg); color:#fff; font-weight:700; font-size:13px; padding:8px 14px; cursor:pointer; }
  .resolve-btn:disabled { opacity:.5; cursor:default; }
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
      <div class="sub">케이씨에이치그룹 경영지원본부 인사총무팀 · 마지막 갱신 ${formatDateTime_(data.generatedAt)} · 접속: ${escapeHtml_(user.name)} (${escapeHtml_(user.role)} 권한)</div>
    </div>
  </div>

<div class="wrap">
  <div class="tiles">${tiles}</div>
  ${changeSection}

  <section>
    <div class="section-head">
      <h2>연도별 · 분기별 발송 현황</h2>
      <select id="yearSelect" onchange="showYear(this.value)">${yearOptions}</select>
    </div>
    <div class="legend">
      <span><i class="wait"></i>대기</span>
      <span><i class="sent"></i>발송완료</span>
      <span><i class="remind"></i>리마인드발송</span>
      <span><i class="check"></i>변경확인필요</span>
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
  var DASH_KEY = ${JSON.stringify(isAdmin ? String(key || '') : '')};
  function resolve(btn, corp) {
    if (!confirm(corp + ' 변경확인을 처리완료로 바꿀까요?')) return;
    btn.disabled = true; btn.textContent = '처리 중…';
    google.script.run
      .withSuccessHandler(function (res) { btn.textContent = '처리완료'; alert(res.msg); location.reload(); })
      .withFailureHandler(function (err) { btn.disabled = false; btn.textContent = '확인했음 · 처리완료'; alert('처리하지 못했습니다: ' + err.message); })
      .resolveChange(DASH_KEY, corp);
  }
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
