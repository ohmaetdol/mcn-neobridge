/**
 * 부업클래스 — 강사 온보딩 자료 수집 API
 *
 * [설치 방법]
 * 1. 구글 시트 새로 만들기 (또는 기존 시트 사용)
 * 2. 확장 프로그램 → Apps Script
 * 3. 이 코드 전체를 Code.gs에 붙여넣기
 * 4. DRIVE_FOLDER_ID에 Google Drive 폴더 ID 입력
 *    (폴더 URL에서 /folders/ 뒤의 문자열)
 * 5. 배포 → 새 배포 → 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 6. 배포 URL을 웹앱 index.html에 입력
 *
 * [시트 구조] (자동 생성됨)
 * 탭: 제출현황
 * 열: 강사명 | 항목코드 | 항목명 | 텍스트응답 | 파일URL | 파일명 | 제출일시
 */

// === 설정 ===
const SHEET_NAME = '제출현황';
// Google Drive 폴더 ID (업로드 파일 저장 위치)
// 빈 문자열이면 루트에 저장
const DRIVE_FOLDER_ID = '';

// === GET: 강사별 제출 데이터 조회 ===
function doGet(e) {
  try {
    const instructor = (e.parameter.name || '').trim();
    if (!instructor) {
      return respond({ error: '강사명이 필요합니다. ?name=강사명' });
    }

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respond({ instructor: instructor, items: {} });
    }

    const headers = data[0];
    const items = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = String(row[0]).trim();
      if (name !== instructor) continue;

      const key = String(row[1]).trim();
      items[key] = {
        text: String(row[3] || ''),
        fileUrl: String(row[4] || ''),
        fileName: String(row[5] || ''),
        date: String(row[6] || '')
      };
    }

    return respond({ instructor: instructor, items: items });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// === POST: 텍스트 저장 / 파일 업로드 ===
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const instructor = (payload.instructor || '').trim();
    const itemKey = (payload.key || '').trim();
    const itemTitle = (payload.title || '').trim();

    if (!instructor || !itemKey) {
      return respond({ error: '강사명과 항목코드가 필요합니다.' });
    }

    const sheet = getOrCreateSheet();

    if (action === 'saveText') {
      const text = payload.text || '';
      saveItem(sheet, instructor, itemKey, itemTitle, text, '', '');
      return respond({ success: true, key: itemKey });
    }

    if (action === 'uploadFile') {
      const fileData = payload.fileData; // base64
      const fileName = payload.fileName || 'file';
      const mimeType = payload.mimeType || 'application/octet-stream';

      if (!fileData) {
        return respond({ error: '파일 데이터가 없습니다.' });
      }

      // Drive에 저장
      const folder = getOrCreateFolder(instructor);
      const decoded = Utilities.base64Decode(fileData);
      const blob = Utilities.newBlob(decoded, mimeType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileUrl = file.getUrl();

      saveItem(sheet, instructor, itemKey, itemTitle, '', fileUrl, fileName);
      return respond({ success: true, key: itemKey, fileUrl: fileUrl, fileName: fileName });
    }

    if (action === 'deleteItem') {
      deleteItem(sheet, instructor, itemKey);
      return respond({ success: true, key: itemKey });
    }

    return respond({ error: '알 수 없는 액션: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// === 헬퍼 ===
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['강사명', '항목코드', '항목명', '텍스트응답', '파일URL', '파일명', '제출일시']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

function getOrCreateFolder(instructor) {
  let parentFolder;
  if (DRIVE_FOLDER_ID) {
    parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } else {
    // 루트에 '부업클래스_온보딩' 폴더 생성/사용
    const rootFolders = DriveApp.getFoldersByName('부업클래스_온보딩');
    if (rootFolders.hasNext()) {
      parentFolder = rootFolders.next();
    } else {
      parentFolder = DriveApp.createFolder('부업클래스_온보딩');
    }
  }

  // 강사별 하위 폴더
  const subFolders = parentFolder.getFoldersByName(instructor);
  if (subFolders.hasNext()) {
    return subFolders.next();
  }
  return parentFolder.createFolder(instructor);
}

function saveItem(sheet, instructor, key, title, text, fileUrl, fileName) {
  const data = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

  // 기존 행 찾기 (같은 강사 + 같은 항목코드)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === instructor && String(data[i][1]).trim() === key) {
      const row = i + 1;
      if (text !== '') sheet.getRange(row, 4).setValue(text);
      if (fileUrl !== '') {
        sheet.getRange(row, 5).setValue(fileUrl);
        sheet.getRange(row, 6).setValue(fileName);
      }
      sheet.getRange(row, 7).setValue(now);
      return;
    }
  }

  // 새 행 추가
  sheet.appendRow([instructor, key, title, text, fileUrl, fileName, now]);
}

function deleteItem(sheet, instructor, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === instructor && String(data[i][1]).trim() === key) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
