/**
 * 부업클래스 — 강사 액션시트 API
 *
 * [설치 방법]
 * 1. 구글 시트 새로 만들기
 * 2. 확장 프로그램 → Apps Script
 * 3. 이 코드 전체를 Code.gs에 붙여넣기
 * 4. 배포 → 새 배포 → 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 5. 배포 URL을 웹앱 index.html에 입력
 *
 * [시트 구조] (직접 만들어주세요)
 * 탭 이름: 액션시트
 * 열: 강사명 | 단계 | 할일 | 마감일 | 상태 | 비고
 * 상태 값: (빈칸)=대기, 진행중, 완료
 * 마감일 형식: 2026-05-15 또는 5/15
 */

const SHEET_NAME = '액션시트';

// === GET: 강사별 액션 조회 ===
function doGet(e) {
  try {
    const instructor = (e.parameter.name || '').trim();
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
        let val = data[i][j];
        // 날짜 객체면 문자열로 변환
        if (val instanceof Date) {
          val = Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
        }
        row[h] = val !== undefined ? String(val).trim() : '';
      });
      row._row = i + 1;

      // 강사명 필터
      if (instructor && row['강사명'] !== instructor) continue;
      rows.push(row);
    }

    return respond({ headers, rows });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// === POST: 상태 업데이트 (강사가 완료 체크) ===
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'updateStatus') {
      const rowNum = payload.row;
      const status = payload.status || '';
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return respond({ error: '시트를 찾을 수 없습니다.' });

      // 상태 열 찾기
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusCol = headers.findIndex(h => String(h).trim() === '상태') + 1;
      if (statusCol === 0) return respond({ error: '상태 열을 찾을 수 없습니다.' });

      sheet.getRange(rowNum, statusCol).setValue(status);
      return respond({ success: true });
    }

    return respond({ error: '알 수 없는 액션: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
