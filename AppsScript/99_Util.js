// 공통 유틸(날짜·문자열·검사)

/** 이메일 형식이 올바른지 간단히 검사 */
function isValidEmail_(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
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

function formatDateTime_(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

/** 스크립트 속성에서 비어 있는 서명·연락처 값의 이름 목록 (모두 있으면 빈 배열) */
function missingProfile_() {
  return PROFILE_KEYS_.filter(k => !String(PROPS_[k] || '').trim());
}

function profileWarning_(missing) {
  return `스크립트 속성 확인 필요 — 비어 있는 값: ${missing.join(', ')}\n` +
    '편집기 "프로젝트 설정 > 스크립트 속성"에서 값을 넣기 전까지 메일을 보내지 않습니다.';
}

function stripHtml_(html) {
  return html.replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ');
}

function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
