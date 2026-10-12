// 실행: node tests/local-test.js  (Apps Script 서비스를 흉내 낸 모의 환경에서 시나리오 검증. 실제 메일·시트는 건드리지 않음)
// 로컬 모의 Apps Script 환경에서 시나리오 테스트
const fs = require('fs'), path = require('path'), vm = require('vm');
const dir = path.join(__dirname, '..', 'AppsScript');
const files = fs.readdirSync(dir).filter(f => f.endsWith('.js')).sort();

function makeEnv(opts = {}) {
  const ME = 'staff@kchglobal.co.kr';
  const HEAD = ['법인명', '수신자성함', '수신자직책', '수신자이메일', '참조이메일', '임원1성함', '임원1직책', '임원2성함', '임원2직책', '태양광대상자', '발송주기(개월)', '다음발송예정일', '마지막발송일', '상태', 'ThreadID', '리마인드발송일', '회신일', '응답토큰', '응답유형', '응답일시', '변경내용', '무응답보고일'];
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const sheets = {
    '법인목록': [HEAD.slice(), ...(opts.rows || [])],
    '로그': [['일시', '법인명', '액션', '상세']],
  };
  const mkSheet = (name) => ({
    getDataRange: () => ({ getValues: () => sheets[name].map(r => r.slice()) }),
    getLastColumn: () => sheets[name][0].length,
    getLastRow: () => sheets[name].length,
    getRange: (r, c, nr, nc) => ({
      setValue: v => {
        if (opts.failWriteAfterSend && env.__sentCount > 0 && HEAD[c - 1] === 'ThreadID') throw new Error('시트 쓰기 오류(모의)');
        while (sheets[name].length < r) sheets[name].push([]);
        sheets[name][r - 1][c - 1] = v;
      },
      getValue: () => (sheets[name][r - 1] || [])[c - 1],
      getValues: () => sheets[name].slice(r - 1, r - 1 + (nr || 1)).map(row => row.slice(c - 1, c - 1 + (nc || 1))),
      setValues: v => { v.forEach((row, i) => { sheets[name][r - 1 + i] = row.slice(); }); return { setFontWeight: () => {} }; },
    }),
    appendRow: row => sheets[name].push(row),
    setFrozenRows: () => {},
  });
  const mails = [];   // 보낸 메일 기록
  const threads = {}; // id -> {messages:[]}
  let tid = 0;
  const msg = (from, subject, body, extra = {}) => ({
    getFrom: () => from, getSubject: () => subject, getPlainBody: () => body, getBody: () => extra.html || body,
    getDate: () => extra.date || new Date(), isDraft: () => !!extra.draft, getAttachments: () => extra.att || [],
    getHeader: () => '<id@x>',
  });
  const mkThread = (id) => ({ getId: () => id, getMessages: () => threads[id].messages });
  const env = {
    __sentCount: 0, mails, threads, sheets, today,
    Logger: { log: (...a) => env.__logs.push(a.join(' ')) }, __logs: [],
    Session: { getEffectiveUser: () => ({ getEmail: () => ME }), getActiveUser: () => ({ getEmail: () => '' }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperties: () => opts.props,
      getProperty: k => opts.props[k] || null,
      setProperty: (k, v) => { opts.props[k] = v; },
    }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] ? mkSheet(n) : null, insertSheet: n => { sheets[n] = [[]]; return mkSheet(n); } }), flush: () => {} },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
    Utilities: {
      getUuid: () => 'tok-' + Math.random().toString(36).slice(2, 12),
      formatDate: (d, tz, f) => { const x = new Date(d), p = n => String(n).padStart(2, '0'); return f.replace('yyyy', x.getFullYear()).replace('MM', p(x.getMonth() + 1)).replace('dd', p(x.getDate())).replace('HH', p(x.getHours())).replace('mm', p(x.getMinutes())); },
      base64Encode: s => Buffer.from(typeof s === 'string' ? s : Buffer.from(s)).toString('base64'),
      base64EncodeWebSafe: s => Buffer.from(s).toString('base64url'),
      sleep: () => {}, Charset: { UTF_8: 'utf8' },
    },
    DriveApp: { getFolderById: () => { const list = (opts.files || []).map(n => ({ getName: () => n, isTrashed: () => false, getDateCreated: () => new Date(2026, 9, 1), getBlob: () => ({ getName: () => n, getBytes: () => [1, 2, 3] }) })); let i = 0; return { getFiles: () => ({ hasNext: () => i < list.length, next: () => list[i++] }) }; } },
    GmailApp: {
      createDraft: (to, subject, body, o) => ({ send: () => {
        env.__sentCount++;
        const id = 't' + (++tid);
        threads[id] = { messages: [msg(ME, subject, body, { html: o.htmlBody })] };
        mails.push({ kind: 'initial', to, subject, html: o.htmlBody, threadId: id });
        return { getThread: () => mkThread(id) };
      } }),
      sendEmail: (to, subject, body) => mails.push({ kind: 'notice', to, subject, body }),
      getThreadById: id => threads[id] ? mkThread(id) : null,
      search: (q) => Object.keys(threads).filter(id => threads[id].messages.some(m => m.getFrom() === ME)).map(mkThread),
    },
    Gmail: { Users: { Messages: { send: (o) => { threads[o.threadId].messages.push(msg(ME, 'Re: x', 'reminder')); mails.push({ kind: 'reminder', threadId: o.threadId }); } }, Drafts: { create: () => {} } } },
    HtmlService: { createHtmlOutput: h => ({ h, setTitle() { return this; }, addMetaTag() { return this; } }) },
    msg, mkThread, ME,
  };
  env.console = console;
  vm.createContext(env);
  const src = files.map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  vm.runInContext(src + '\n;this.__api = { mainDaily, classifyReplyText_, extractReplyText_, doGet, resolveChange, submitResponse, HEADERS, renderDashboardHtml_, getDashboardData_, 발송전_점검 };', env);
  return env;
}

const props = { STAFF_NAME: '홍길동', STAFF_TITLE: '매니저', STAFF_EMAIL: 'staff@kchglobal.co.kr', MANAGER_NAME: '김철수', MANAGER_TITLE: '팀장', MANAGER_EMAIL: 'm@kchglobal.co.kr', OFFICE_ADDRESS: '서울', SENDER_NAME: 'KCH',
  DASHBOARD_USERS: JSON.stringify([{ key: 'admin-key-1234567890', name: '박성현', role: '관리' }, { key: 'view-key-1234567890', name: '이윤재', role: '조회' }]) };
const today = new Date(); today.setHours(0, 0, 0, 0);
const row = (corp, status = '', extra = {}) => {
  const r = [corp, '이영희', '대표', 'ceo@example.com', '', '이영희', '대표이사', '', '', '이영희', 3, today, '', status, '', '', '', '', '', '', '', ''];
  Object.entries(extra).forEach(([k, v]) => { r[['법인명', '수신자성함', '수신자직책', '수신자이메일', '참조이메일', '임원1성함', '임원1직책', '임원2성함', '임원2직책', '태양광대상자', '발송주기(개월)', '다음발송예정일', '마지막발송일', '상태', 'ThreadID', '리마인드발송일', '회신일', '응답토큰', '응답유형', '응답일시', '변경내용', '무응답보고일'].indexOf(k)] = v; });
  return r;
};
const col = (env, r, h) => env.sheets['법인목록'][r][env.__api.HEADERS.indexOf(h)];
let pass = 0, fail = 0;
const check = (name, cond, info = '') => { if (cond) { pass++; console.log('PASS', name); } else { fail++; console.log('FAIL', name, info); } };

// 1. 회신 본문 판정
{
  const e = makeEnv({ props: { ...props } });
  const c = e.__api.classifyReplyText_;
  const cases = [
    ['변경사항 없습니다. 감사합니다.', '변동없음'],
    ['확인했습니다. 변동 없음', '변동없음'],
    ['대표이사 주소 변경 없습니다', '변동없음'],
    ['이상 없습니다', '변동없음'],
    ['대표이사 주소가 변경되었습니다. 변경일은 9월 3일입니다.', '확인필요'],
    ['변경사항 없습니다. 다만 사내이사가 사임 예정입니다.', '확인필요'],
    ['확인했습니다.', '확인필요'],
    ['', '확인필요'],
  ];
  cases.forEach(([t, want]) => check(`판정: "${t}" → ${want}`, c(t, false).type === want, c(t, false).type));
  check('판정: 첨부 있으면 확인필요', c('변경사항 없습니다', true).type === '확인필요');
  const q = e.__api.extractReplyText_('변경사항 없습니다\n\n2026년 10월 5일 (월) 오전 9:00, 홍길동 <a@b>님이 작성:\n> 대표이사 주소지 변경 등');
  check('인용 잘라내기', q === '변경사항 없습니다', JSON.stringify(q));
}

// 2. 정상 첫 발송
{
  const e = makeEnv({ props: { ...props }, rows: [row('샘플솔라')], files: ['등기부등본_샘플솔라.pdf'] });
  e.__api.mainDaily();
  check('첫 발송: 메일 1통', e.mails.filter(m => m.kind === 'initial').length === 1);
  check('첫 발송: 상태 발송완료', col(e, 1, '상태') === '발송완료', col(e, 1, '상태'));
  check('첫 발송: 메일 본문에 시트 토큰 포함', e.mails[0].html.includes(col(e, 1, '응답토큰')));
}

// 3. 발송 후 기록 실패 → 발송실패로 잘못 적지 않고 발송중 유지 → 다음 실행에서 복구, 재발송 없음
{
  const opts = { props: { ...props }, rows: [row('샘플솔라')], files: ['등기부등본_샘플솔라.pdf'], failWriteAfterSend: true };
  const e = makeEnv(opts);
  e.__api.mainDaily();
  check('기록 실패: 상태가 발송실패가 아님', col(e, 1, '상태') === '발송중', col(e, 1, '상태'));
  opts.failWriteAfterSend = false;
  e.__api.mainDaily();
  check('복구: 메일은 여전히 1통(재발송 없음)', e.mails.filter(m => m.kind === 'initial').length === 1, e.mails.filter(m => m.kind === 'initial').length);
  check('복구: 상태 발송완료 + ThreadID 채움', col(e, 1, '상태') === '발송완료' && !!col(e, 1, 'ThreadID'), col(e, 1, '상태'));
  check('복구: 로그에 복구 기록', e.sheets['로그'].some(r => String(r[3]).includes('복구')));
}

// 4. 발송중인데 실제로 안 나간 경우 → 다시 보냄
{
  const e = makeEnv({ props: { ...props }, rows: [row('샘플솔라', '발송중', { '응답토큰': 'tok-old', '마지막발송일': today })], files: ['등기부등본_샘플솔라.pdf'] });
  e.__api.mainDaily();
  check('미발송 발송중 → 정상 발송 1통', e.mails.filter(m => m.kind === 'initial').length === 1);
  check('미발송 발송중 → 발송완료', col(e, 1, '상태') === '발송완료', col(e, 1, '상태'));
}

// 5. 회신: 변경 없음 / 변경 있음
{
  const e = makeEnv({ props: { ...props }, rows: [row('가나'), row('다라')], files: ['등기부등본_가나.pdf', '등기부등본_다라.pdf'] });
  e.__api.mainDaily();
  const t1 = col(e, 1, 'ThreadID'), t2 = col(e, 2, 'ThreadID');
  e.threads[t1].messages.push(e.msg('ceo@example.com', 'Re', '변경사항 없습니다.\n\n> 원문'));
  e.threads[t2].messages.push(e.msg('ceo@example.com', 'Re', '대표이사 주소가 9/3에 변경되었습니다.\n> 원문'));
  e.__api.mainDaily();
  check('회신 변경없음 → 회신완료', col(e, 1, '상태') === '회신완료' && col(e, 1, '응답유형') === '변동없음(메일)', col(e, 1, '상태'));
  check('회신 변경있음 → 변경확인필요', col(e, 2, '상태') === '변경확인필요', col(e, 2, '상태'));
  check('변경내용 칸에 회신 발췌', String(col(e, 2, '변경내용')).includes('주소가 9/3에 변경'));
  check('담당자에게 변경확인필요 알림', e.mails.some(m => m.kind === 'notice' && m.subject.includes('[변경확인필요] 다라')));

  // 6. 권한: 조회 키로 처리완료 시도 → 거부, 관리 키 → 처리완료
  let denied = false;
  try { e.__api.resolveChange('view-key-1234567890', '다라'); } catch (err) { denied = true; }
  check('조회 권한은 처리완료 불가(서버 거부)', denied && col(e, 2, '상태') === '변경확인필요');
  const res = e.__api.resolveChange('admin-key-1234567890', '다라');
  check('관리 권한 처리완료 → 회신완료', col(e, 2, '상태') === '회신완료' && col(e, 2, '응답유형') === '변경처리완료', res && res.msg);

  // 7. 대시보드 접근
  const noKey = e.__api.doGet({ parameter: {} }).h;
  check('키 없으면 접근 거부', noKey.includes('접근 권한이 필요'));
  const badKey = e.__api.doGet({ parameter: { key: 'wrong-key-xxxxxxxxxxxx' } }).h;
  check('틀린 키 거부 + 로그', badKey.includes('접근 권한이 필요') && e.sheets['로그'].some(r => r[2] === '접근거부'));
  const adminHtml = e.__api.doGet({ parameter: { key: 'admin-key-1234567890' } }).h;
  const viewHtml = e.__api.doGet({ parameter: { key: 'view-key-1234567890' } }).h;
  check('관리 화면: 수신자 실명 표시', adminHtml.includes('이영희'));
  check('조회 화면: 수신자 이름 가림', !viewHtml.includes('<td>이영희</td>') && viewHtml.includes('이OO'), '');
  check('조회 화면: 키를 페이지에 싣지 않음', !viewHtml.includes('view-key-1234567890'));
  check('열람 기록 로그', e.sheets['로그'].filter(r => r[2] === '대시보드열람').length === 2);
}

// 8. 리마인드 중복 방지: 스레드에 내 메일이 이미 2통이면 다시 안 보냄
{
  const e = makeEnv({ props: { ...props }, rows: [row('가나')], files: ['등기부등본_가나.pdf'] });
  e.__api.mainDaily();
  const r = e.sheets['법인목록'][1];
  const old = new Date(today); old.setDate(old.getDate() - 8);
  r[e.__api.HEADERS.indexOf('마지막발송일')] = old;
  const tId = col(e, 1, 'ThreadID');
  e.threads[tId].messages.push(e.msg(e.ME, 'Re', 'reminder already'));
  e.__api.mainDaily();
  check('리마인드 이미 있음 → 재발송 안 함', e.mails.filter(m => m.kind === 'reminder').length === 0);
  check('리마인드 이미 있음 → 상태 리마인드발송으로 맞춤', col(e, 1, '상태') === '리마인드발송');
  // 일반 리마인드
  const e2 = makeEnv({ props: { ...props }, rows: [row('가나')], files: ['등기부등본_가나.pdf'] });
  e2.__api.mainDaily();
  e2.sheets['법인목록'][1][e2.__api.HEADERS.indexOf('마지막발송일')] = old;
  e2.__api.mainDaily();
  check('일반 리마인드 1통 발송', e2.mails.filter(m => m.kind === 'reminder').length === 1);
}

// 9. 변경확인필요 미처리 상태로 다음 분기 → 발송 보류
{
  const e = makeEnv({ props: { ...props }, rows: [row('가나', '변경확인필요')], files: ['등기부등본_가나.pdf'] });
  e.__api.mainDaily();
  check('변경확인 미처리 → 새 분기 메일 보류', e.mails.filter(m => m.kind === 'initial').length === 0);
  check('요약 메일에 보류 안내', e.mails.some(m => m.kind === 'notice' && m.body.includes('변경확인 미처리')));
}

// 10. 버튼 응답이 변경확인필요를 덮어쓰지 않음
{
  const e = makeEnv({ props: { ...props }, rows: [row('가나', '변경확인필요', { '응답토큰': 'tok-abc' })], files: [] });
  e.__api.submitResponse('tok-abc');
  check('버튼이 변경확인필요를 닫지 않음', col(e, 1, '상태') === '변경확인필요');
}

// 11. 점검 함수
{
  const e = makeEnv({ props: { ...props, DASHBOARD_USERS: '{bad json' }, rows: [row('가나', '발송중')], files: ['등기부등본_가나.pdf'] });
  const out = e.__api.발송전_점검();
  check('점검: JSON 오류·발송중 행 잡아냄', out.includes('DASHBOARD_USERS 형식 오류') && out.includes('발송중으로 남아'), out);
}

console.log(`\n결과: ${pass} 통과 / ${fail} 실패`);
process.exit(fail ? 1 : 0);
