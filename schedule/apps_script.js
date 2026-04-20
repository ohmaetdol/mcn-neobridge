// 스케줄 웹앱 연동 — Apps Script
// 시트: https://docs.google.com/spreadsheets/d/1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s
//
// 배포 방법:
// 1. 위 시트 > 확장 프로그램 > Apps Script
// 2. 이 코드 전체 붙여넣기 > 저장
// 3. 배포 > 새 배포 > 웹 앱 > 실행 주체: 본인 / 액세스: 모든 사용자
// 4. 배포 URL을 schedule/index.html의 API_URL에 입력

var SHEET_ID = '1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s';

function doGet(e) {
  var data = getScheduleData();
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScheduleData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var channels = [
    {name: '사장찍어주는남자', color: '#ea580c'},
    {name: '부업소개하는남자', color: '#4f46e5'}
  ];

  var result = [];

  for (var c = 0; c < channels.length; c++) {
    var sheet = ss.getSheetByName(channels[c].name);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();
    var months = [];
    var currentMonth = null;
    var currentItems = [];

    for (var i = 2; i < data.length; i++) {
      var row = data[i];

      // 월 구분선 감지 (── 2026년 X월 ──)
      var cell3 = String(row[3] || '');
      if (cell3.indexOf('──') === 0) {
        if (currentMonth) {
          months.push({label: currentMonth, items: currentItems});
        }
        currentMonth = cell3.replace(/──/g, '').trim();
        currentItems = [];
        continue;
      }

      // 빈 행 스킵 (No도 없고 강사명도 없는 행)
      if (!row[0] && !row[3]) continue;

      // 첫 월이 없으면 현재 월 기준
      if (!currentMonth) {
        var today = new Date();
        currentMonth = today.getFullYear() + '년 ' + (today.getMonth() + 1) + '월';
      }

      currentItems.push({
        no: String(row[0] || ''),
        lectureDate: String(row[1] || ''),
        day: String(row[2] || ''),
        instructor: String(row[3] || ''),
        status: String(row[4] || ''),
        editDeadline: String(row[5] || ''),
        uploadDeadline: String(row[6] || ''),
        dday: String(row[7] || ''),
        contact: String(row[8] || ''),
        refLink: String(row[9] || ''),
        memo: String(row[10] || '')
      });
    }

    // 마지막 월 추가
    if (currentMonth) {
      months.push({label: currentMonth, items: currentItems});
    }

    result.push({
      name: channels[c].name,
      color: channels[c].color,
      months: months
    });
  }

  var now = new Date();
  var todayStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
  var updatedStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy.MM.dd HH:mm');

  return {
    channels: result,
    today: todayStr,
    updated: updatedStr
  };
}
