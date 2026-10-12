// 드라이브 등기부등본 파일 찾기(재시도 포함)

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
