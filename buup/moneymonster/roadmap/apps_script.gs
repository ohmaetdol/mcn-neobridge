/**
 * 부업클래스 x 머니몬스터 — 런칭 로드맵 API
 *
 * [설치 방법]
 * 1. 구글 시트 → 확장 프로그램 → Apps Script
 * 2. 이 코드 전체를 복사해서 Code.gs에 붙여넣기
 * 3. 배포 → 새 배포 → 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 나오는 URL 복사
 *
 * [시트 구조]
 * 탭 이름: 로드맵
 * 열: 구간 | 날짜 | 카테고리 | 할일 | 담당 | 비고 | 상태
 * 상태 값: (빈칸)=대기, 진행중, 완료
 */

const SHEET_NAME = '로드맵';

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return respond({ error: '시트 탭 "' + SHEET_NAME + '"을 찾을 수 없습니다.' });
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respond({ headers: [], rows: [] });
    }

    const headers = data[0].map(h => String(h).trim());
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((h, j) => {
        row[h] = data[i][j] !== undefined ? String(data[i][j]).trim() : '';
      });
      row._row = i + 1;
      rows.push(row);
    }

    return respond({ headers, rows });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return respond({ error: '시트를 찾을 수 없습니다.' });
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => String(h).trim());

    if (payload.action === 'updateStatus') {
      const col = headers.indexOf('상태') + 1;
      if (col > 0 && payload.row > 1) {
        sheet.getRange(payload.row, col).setValue(payload.value);
        return respond({ success: true });
      }
      return respond({ error: '상태 열을 찾을 수 없습니다.' });
    }

    if (payload.action === 'updateNote') {
      const col = headers.indexOf('비고') + 1;
      if (col > 0 && payload.row > 1) {
        sheet.getRange(payload.row, col).setValue(payload.value);
        return respond({ success: true });
      }
      return respond({ error: '비고 열을 찾을 수 없습니다.' });
    }

    return respond({ error: '알 수 없는 액션: ' + payload.action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
