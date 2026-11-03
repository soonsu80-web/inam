/**
 * 학생작품 갤러리 - Code.gs
 * 구글 Apps Script 웹 앱 백엔드
 *
 * [필수 설정] 프로젝트 설정(⚙) > 스크립트 속성에 아래 값을 추가하세요.
 *   GEMINI_API_KEY  : Gemini API 키 (코드에 직접 쓰지 마세요)
 *   ADMIN_PASSWORD  : 관리자 모드 비밀번호
 *   GEMINI_MODEL    : (선택) 이미지 생성 모델명. 비워두면 아래 MODEL_CANDIDATES를 순서대로 시도
 */

const SHEET_ID = '1iP-J755UMKZdr7Rc7LEGRjINgyaarmjrG9n55HZWpFk';
const SHEET_NAME = '기록';
const HEADERS = ['번호', '등록일시', '학번', '이름', '작품URL', '작품종류', '설명', '이미지', '승인상태', '좋아요수'];
const TYPES = ['게임', '웹앱', '게시글', '자료'];
const IMAGE_FOLDER_NAME = '학생작품갤러리_이미지';

// 이미지 생성이 가능한 모델 후보 (Gemini 1.5 Flash는 텍스트 전용이라 이미지 생성 불가)
const MODEL_CANDIDATES = ['gemini-3.1-flash-image-preview', 'gemini-2.5-flash-image'];

const TYPE_PROMPTS = {
  '게임': 'A vibrant modern flat illustration representing a video game: a stylized game controller, pixel stars and floating platforms, purple and cyan gradient background. No text, no letters.',
  '웹앱': 'A clean modern flat illustration representing a web application: browser window with colorful UI cards, charts and toggles, blue and teal gradient background. No text, no letters.',
  '게시글': 'A warm modern flat illustration representing a written post or blog article: paper sheets, a pen, speech bubbles and a bookmark, orange and pink gradient background. No text, no letters.',
  '자료': 'A tidy modern flat illustration representing study materials and resources: stacked folders, documents, a lightbulb and a magnifier, green and yellow gradient background. No text, no letters.'
};

/* ============ 웹 앱 진입점 ============ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('작품 갤러리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ============ 시트 유틸 ============ */

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values.map(function (r, i) {
    let date = r[1];
    if (date instanceof Date) {
      date = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    }
    return {
      row: i + 2,
      id: String(r[0]),
      date: String(date),
      studentId: String(r[2]),
      name: String(r[3]),
      url: String(r[4]),
      type: String(r[5]),
      desc: String(r[6]),
      image: String(r[7]),
      status: String(r[8]),
      likes: Number(r[9]) || 0
    };
  });
}

function publicView_(p) {
  return {
    id: p.id, date: p.date, studentId: p.studentId, name: p.name,
    url: p.url, type: p.type, desc: p.desc, image: p.image, likes: p.likes
  };
}

// 수식 주입 방지: =, +, -, @ 로 시작하면 앞에 작은따옴표를 붙임
function safeText_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function findRowById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/* ============ 갤러리(공개) ============ */

function getApprovedPosts() {
  return readAll_()
    .filter(function (p) { return p.status === '승인'; })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; })
    .map(publicView_);
}

function addPost(data) {
  data = data || {};
  const studentId = safeText_(data.studentId, 20);
  const name = safeText_(data.name, 30);
  const url = String(data.url || '').trim().slice(0, 500);
  const type = String(data.type || '').trim();
  const desc = safeText_(data.desc, 300);

  if (!studentId || !name || !url || !desc) throw new Error('모든 항목을 입력해 주세요.');
  if (!/^https?:\/\//i.test(url)) throw new Error('작품 URL은 http:// 또는 https:// 로 시작해야 합니다.');
  if (TYPES.indexOf(type) < 0) throw new Error('작품 종류를 선택해 주세요.');

  const image = getTypeImage_(type);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const id = 'P' + Date.now();
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const rowNo = sh.getLastRow() + 1;
    // 학번(앞자리 0)과 일시가 자동 변환되지 않도록 텍스트 서식 지정
    sh.getRange(rowNo, 1, 1, 3).setNumberFormat('@');
    sh.getRange(rowNo, 1, 1, HEADERS.length).setValues([[id, now, studentId, name, url, type, desc, image, '대기', 0]]);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function likePost(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const row = findRowById_(sh, id);
    if (row < 0) throw new Error('게시글을 찾을 수 없습니다.');
    const cell = sh.getRange(row, 10);
    const next = (Number(cell.getValue()) || 0) + 1;
    cell.setValue(next);
    return next;
  } finally {
    lock.releaseLock();
  }
}

/* ============ 관리자 ============ */

function checkAdmin_(pw) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!saved) throw new Error('관리자 비밀번호가 설정되지 않았습니다. 스크립트 속성에 ADMIN_PASSWORD를 추가하세요.');
  if (String(pw) !== saved) throw new Error('비밀번호가 올바르지 않습니다.');
}

function adminLogin(pw) {
  checkAdmin_(pw);
  return true;
}

function getAllPosts(pw) {
  checkAdmin_(pw);
  return readAll_().sort(function (a, b) {
    const pa = a.status === '대기' ? 0 : 1;
    const pb = b.status === '대기' ? 0 : 1;
    if (pa !== pb) return pa - pb;
    return a.date < b.date ? 1 : -1;
  });
}

function approvePost(pw, id) {
  checkAdmin_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const row = findRowById_(sh, id);
    if (row < 0) throw new Error('게시글을 찾을 수 없습니다.');
    sh.getRange(row, 9).setValue('승인');
    return true;
  } finally {
    lock.releaseLock();
  }
}

function deletePost(pw, id) {
  checkAdmin_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const row = findRowById_(sh, id);
    if (row < 0) throw new Error('게시글을 찾을 수 없습니다.');
    sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/* ============ Gemini 이미지 (종류별 1회 생성 후 재사용) ============ */

function getTypeImage_(type) {
  const props = PropertiesService.getScriptProperties();
  const key = 'IMG_' + type;
  const cached = props.getProperty(key);
  if (cached) return cached;

  try {
    const url = generateTypeImage_(type);
    if (url) {
      props.setProperty(key, url);
      return url;
    }
  } catch (e) {
    console.error('이미지 생성 실패(' + type + '): ' + e);
  }
  return ''; // 비워두면 화면에서 종류별 기본 일러스트로 대체
}

function generateTypeImage_(type) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('스크립트 속성에 GEMINI_API_KEY가 없습니다.');

  const custom = props.getProperty('GEMINI_MODEL');
  const models = custom ? [custom] : MODEL_CANDIDATES;
  const prompt = TYPE_PROMPTS[type];
  let lastError = '';

  for (let i = 0; i < models.length; i++) {
    const endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent';
    const res = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseModalities: ['TEXT', 'IMAGE'] }
      }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      lastError = models[i] + ' → ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200);
      continue;
    }
    const json = JSON.parse(res.getContentText());
    const parts = (((json.candidates || [])[0] || {}).content || {}).parts || [];
    for (let j = 0; j < parts.length; j++) {
      const inline = parts[j].inlineData || parts[j].inline_data;
      if (inline && inline.data) {
        const mime = inline.mimeType || inline.mime_type || 'image/png';
        const blob = Utilities.newBlob(Utilities.base64Decode(inline.data), mime, 'gallery_' + type + '.png');
        return saveToDrive_(blob);
      }
    }
    lastError = models[i] + ' → 응답에 이미지가 없음';
  }
  throw new Error(lastError || '이미지 생성 실패');
}

function saveToDrive_(blob) {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(IMAGE_FOLDER_NAME);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    console.error('공유 설정 실패(조직 정책일 수 있음): ' + e);
  }
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

/* ============ 에디터에서 직접 실행하는 도구 ============ */

// 처음 한 번 실행하면 네 가지 종류 이미지를 미리 만들어 둡니다. (권한 승인도 이때 진행)
function setupImages() {
  TYPES.forEach(function (t) {
    const url = getTypeImage_(t);
    console.log(t + ' → ' + (url || '실패(로그 확인)'));
  });
}

// 이미지를 다시 만들고 싶을 때 캐시를 지웁니다.
function resetImages() {
  const props = PropertiesService.getScriptProperties();
  TYPES.forEach(function (t) { props.deleteProperty('IMG_' + t); });
  console.log('이미지 캐시를 삭제했습니다. setupImages()를 다시 실행하세요.');
}
