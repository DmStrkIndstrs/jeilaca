/**
 * 진도기록 시트 자동 입력용 Apps Script
 *
 * 배포 방법:
 * 1) 중등 또는 고등 진도과제 기록표 시트 아무거나 열기
 * 2) 상단 메뉴 확장 프로그램 > Apps Script
 * 3) 기본으로 열려있는 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기
 * 4) 우측 상단 저장(디스크 아이콘)
 * 5) 배포 > 새 배포 > 유형 선택(톱니바퀴) > 웹 앱
 *    - 설명: 아무거나 (예: 진도기록 입력 API)
 *    - 다음 사용자 권한으로 실행: 나
 *    - 액세스 권한이 있는 사용자: 전체
 * 6) 배포 클릭 → 권한 승인(본인 계정) → 생성된 웹 앱 URL 복사
 *    (예: https://script.google.com/macros/s/AKfycb.../exec)
 * 7) 그 URL을 그대로 전달해주면 홈페이지 쪽 입력 폼에 연결합니다.
 */

const SHEETS = {
  mid:  '1lMg29yBF5H1npC0afxubHsfnjxdjA6jl_IBK39eLp6s', // 진도과제 기록표(중등)
  high: '1Mqldv3QQZ1c5CaR6n9onFKTE8iTbyuyHlgKEBAzsBqo'  // 진도과제 기록표(고등)
};
const SHEET_TAB_NAME = '입력';
const DAY_COL_START = 7; // G열 = 월 (1-indexed)
const DAY_NAMES = ['월', '화', '수', '목', '금', '토', '일'];

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    result = writeCell(body);
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeCell(body) {
  const { target, week, grade, className, teacher, subject, rowType, day, text } = body;

  if (!SHEETS[target]) return { ok: false, error: 'invalid target: ' + target };
  const dayIdx = DAY_NAMES.indexOf(day);
  if (dayIdx === -1) return { ok: false, error: 'invalid day: ' + day };

  const ss = SpreadsheetApp.openById(SHEETS[target]);
  const sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return { ok: false, error: '"입력" 시트를 찾을 수 없습니다' };

  const data = sheet.getDataRange().getValues();

  let curWeek = '', curGrade = '', curClass = '', curTeacher = '', curSubject = '';
  let targetRow = -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawWeek = String(row[0] || '').trim();
    const rawGrade = String(row[1] || '').trim();
    const rawClass = String(row[2] || '').trim();
    const rawTeacher = String(row[3] || '').trim();
    const rawSubject = String(row[4] || '').trim();
    const rType = String(row[5] || '').trim();

    if (rawWeek) curWeek = rawWeek;
    if (rawGrade) curGrade = rawGrade;
    if (rawClass) curClass = rawClass;
    if (rawTeacher) curTeacher = rawTeacher;
    if (rawSubject) curSubject = rawSubject;

    if (curWeek === week && curGrade === grade && curClass === className &&
        curTeacher === teacher && curSubject === subject && rType === rowType) {
      targetRow = i + 1; // 1-indexed sheet row
      break;
    }
  }

  if (targetRow === -1) {
    return { ok: false, error: '해당 행을 찾지 못했습니다 (주차/학년/반/선생님/과목 확인 필요)' };
  }

  const col = DAY_COL_START + dayIdx;
  sheet.getRange(targetRow, col).setValue(text);

  return { ok: true, row: targetRow, col: col };
}

/** 배포 후 브라우저에서 웹앱 URL을 직접 열어보면 이 응답이 보입니다 (정상 배포 확인용) */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: '진도기록 입력 API 정상 작동 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}
