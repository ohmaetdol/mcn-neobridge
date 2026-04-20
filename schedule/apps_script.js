// 스케줄 웹앱 연동 + 메뉴 + 자동정렬 — Apps Script
// 시트: https://docs.google.com/spreadsheets/d/1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s
//
// 배포 방법:
// 1. 위 시트 > 확장 프로그램 > Apps Script
// 2. 이 코드 전체 붙여넣기 > 저장
// 3. 배포 > 배포 관리 > 연필 아이콘 > 새 버전 > 배포

var SHEET_ID = '1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s';
var CHANNELS = ['사장찍어주는남자', '부업소개하는남자'];

// ── 메뉴 ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('스케줄 관리')
    .addItem('새 강의 추가', 'showAddDialog')
    .addItem('전체 현황', 'showSummary')
    .addItem('날짜순 정렬', 'sortAllChannels')
    .addSeparator()
    .addItem('웹앱 열기', 'openWebApp')
    .addToUi();
}

// ── 자동 정렬 (날짜 입력/변경 시) ─────────────────
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var name = sheet.getName();
  if (CHANNELS.indexOf(name) === -1) return;

  var col = e.range.getColumn();
  // B열(날짜) 또는 D열(강사명) 변경 시 자동 정렬
  if (col === 2 || col === 4) {
    sortChannel(sheet);
  }
}

function sortAllChannels() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  for (var i = 0; i < CHANNELS.length; i++) {
    var sheet = ss.getSheetByName(CHANNELS[i]);
    if (sheet) sortChannel(sheet);
  }
  SpreadsheetApp.getUi().alert('날짜순 정렬 완료!');
}

function sortChannel(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  var range = sheet.getRange(3, 1, lastRow - 2, 11);
  var values = range.getValues();

  // 데이터 있는 행 / 빈 행 분리
  var filled = [];
  var empty = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][3]) { // 강사명 있으면 데이터 행
      filled.push(values[i]);
    } else {
      empty.push(values[i]);
    }
  }

  // 데이터 행을 날짜순 정렬
  filled.sort(function(a, b) {
    var da = a[1] instanceof Date ? a[1].getTime() : 0;
    var db = b[1] instanceof Date ? b[1].getTime() : 0;
    if (da === 0 && db === 0) return 0;
    if (da === 0) return 1;
    if (db === 0) return -1;
    return da - db;
  });

  // 번호 재부여 + 합치기
  var result = [];
  for (var i = 0; i < filled.length; i++) {
    filled[i][0] = i + 1;
    result.push(filled[i]);
  }
  var nextNo = filled.length + 1;
  for (var i = 0; i < empty.length; i++) {
    empty[i][0] = nextNo++;
    result.push(empty[i]);
  }

  // 시트에 쓰기
  range.setValues(result);

  // 수식 재설정 (C, F, G, H열)
  for (var i = 0; i < result.length; i++) {
    var row = i + 3;
    sheet.getRange(row, 3).setFormula('=IF(B' + row + '="","",CHOOSE(WEEKDAY(B' + row + ',1),"일","월","화","수","목","금","토"))');
    sheet.getRange(row, 6).setFormula('=IF(B' + row + '="","",B' + row + '+9)');
    sheet.getRange(row, 7).setFormula('=IF(B' + row + '="","",B' + row + '+14)');
    sheet.getRange(row, 8).setFormula('=IF(G' + row + '="","",IF(G' + row + '-TODAY()>=0,"D-"&(G' + row + '-TODAY()),"D+"&(TODAY()-G' + row + ')))');
  }
}

// ── 새 강의 추가 대화상자 ─────────────────────────
function showAddDialog() {
  var html = HtmlService.createHtmlOutput(getAddFormHtml())
    .setWidth(420)
    .setHeight(480)
    .setTitle('새 강의 추가');
  SpreadsheetApp.getUi().showModalDialog(html, '새 강의 추가');
}

function getAddFormHtml() {
  return '<style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:-apple-system,sans-serif;padding:24px;background:#fff}' +
    'h2{font-size:16px;margin-bottom:20px;color:#1a1d23}' +
    'label{display:block;font-size:13px;font-weight:600;color:#5a6170;margin-bottom:4px;margin-top:14px}' +
    'select,input{width:100%;padding:10px 12px;border:1px solid #d1d5db;border-radius:8px;font-size:14px;outline:none}' +
    'select:focus,input:focus{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(79,70,229,0.1)}' +
    '.btn{width:100%;padding:12px;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-top:20px}' +
    '.btn-primary{background:#1c2b3a;color:#fff}' +
    '.btn-primary:hover{background:#2d3f52}' +
    '.msg{padding:12px;border-radius:8px;margin-top:12px;font-size:13px;display:none}' +
    '.msg-ok{background:#f0fdf4;color:#16a34a;display:block}' +
    '.msg-err{background:#fef2f2;color:#dc2626;display:block}' +
    '</style>' +
    '<h2>새 강의 추가</h2>' +
    '<label>채널</label>' +
    '<select id="channel">' +
    '<option value="사장찍어주는남자">사장찍어주는남자</option>' +
    '<option value="부업소개하는남자">부업소개하는남자</option>' +
    '</select>' +
    '<label>무료강의일</label>' +
    '<input type="date" id="lectureDate">' +
    '<label>강사명</label>' +
    '<input type="text" id="instructor" placeholder="강사 이름">' +
    '<label>상태</label>' +
    '<select id="status">' +
    '<option value="섭외중">섭외중</option>' +
    '<option value="섭외완료">섭외완료</option>' +
    '<option value="기획중">기획중</option>' +
    '<option value="촬영완료">촬영완료</option>' +
    '<option value="편집중">편집중</option>' +
    '<option value="검수중">검수중</option>' +
    '<option value="업로드완료">업로드완료</option>' +
    '</select>' +
    '<label>연락처</label><input type="tel" id="contact" placeholder="010-0000-0000">' +
    '<label>참고링크</label><input type="url" id="refLink" placeholder="https://">' +
    '<label>메모</label><input type="text" id="memo" placeholder="메모 (선택)">' +
    '<button class="btn btn-primary" onclick="submit()">추가</button>' +
    '<div id="msg" class="msg"></div>' +
    '<script>' +
    'function submit(){' +
    '  var d={channel:document.getElementById("channel").value,lectureDate:document.getElementById("lectureDate").value,instructor:document.getElementById("instructor").value,status:document.getElementById("status").value,contact:document.getElementById("contact").value,refLink:document.getElementById("refLink").value,memo:document.getElementById("memo").value};' +
    '  if(!d.instructor){showMsg("강사명을 입력하세요","err");return}' +
    '  google.script.run.withSuccessHandler(function(r){if(r.success){showMsg(r.message,"ok");setTimeout(function(){google.script.host.close()},1500)}else{showMsg(r.message,"err")}}).withFailureHandler(function(e){showMsg("오류: "+e.message,"err")}).addLecture(d);' +
    '}' +
    'function showMsg(t,type){var m=document.getElementById("msg");m.textContent=t;m.className="msg msg-"+type}' +
    '</script>';
}

function addLecture(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(data.channel);
  if (!sheet) return {success: false, message: '채널 탭을 찾을 수 없습니다.'};

  // 빈 데이터 행 찾기 (강사명 비어있는 첫 번째 행)
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  for (var i = 3; i <= lastRow; i++) {
    var inst = sheet.getRange(i, 4).getValue();
    if (!inst) { targetRow = i; break; }
  }
  if (targetRow === -1) {
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1).setValue(lastRow - 1);
  }

  // 데이터 입력
  if (data.lectureDate) sheet.getRange(targetRow, 2).setValue(new Date(data.lectureDate));
  sheet.getRange(targetRow, 4).setValue(data.instructor);
  sheet.getRange(targetRow, 5).setValue(data.status);
  if (data.contact) sheet.getRange(targetRow, 9).setValue(data.contact);
  if (data.refLink) sheet.getRange(targetRow, 10).setValue(data.refLink);
  if (data.memo) sheet.getRange(targetRow, 11).setValue(data.memo);

  // 자동 정렬
  sortChannel(sheet);

  return {success: true, message: data.instructor + ' 강의가 추가되었습니다!'};
}

// ── 전체 현황 ─────────────────────────────────────
function showSummary() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var msg = '';
  for (var c = 0; c < CHANNELS.length; c++) {
    var sheet = ss.getSheetByName(CHANNELS[c]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var total = 0, statusCounts = {};
    for (var i = 2; i < data.length; i++) {
      if (!data[i][3]) continue;
      total++;
      var st = data[i][4] || '미정';
      statusCounts[st] = (statusCounts[st] || 0) + 1;
    }
    msg += '[ ' + CHANNELS[c] + ' ] — ' + total + '건\n';
    for (var st in statusCounts) msg += '  ' + st + ': ' + statusCounts[st] + '건\n';
    msg += '\n';
  }
  SpreadsheetApp.getUi().alert('스케줄 현황\n\n' + msg);
}

// ── 웹앱 열기 ─────────────────────────────────────
function openWebApp() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("https://ohmaetdol.github.io/mcn-neobridge/schedule/");google.script.host.close();</script>'
  ).setWidth(200).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, '웹앱 열기');
}

// ══════════════════════════════════════════════════
// ── 웹앱 API ─────────────────────────────────────
// ══════════════════════════════════════════════════

function doGet(e) {
  var data = getScheduleData();
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScheduleData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var colors = {'사장찍어주는남자': '#ea580c', '부업소개하는남자': '#4f46e5'};
  var result = [];

  for (var c = 0; c < CHANNELS.length; c++) {
    var sheet = ss.getSheetByName(CHANNELS[c]);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();
    var monthMap = {};

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // No 없으면 스킵

      var monthLabel = '';
      if (row[1] instanceof Date) {
        monthLabel = row[1].getFullYear() + '년 ' + (row[1].getMonth() + 1) + '월';
      } else {
        monthLabel = '미정';
      }

      if (!monthMap[monthLabel]) monthMap[monthLabel] = [];
      monthMap[monthLabel].push({
        no: String(row[0] || ''),
        lectureDate: fmtDate(row[1]),
        day: String(row[2] || ''),
        instructor: String(row[3] || ''),
        status: String(row[4] || ''),
        editDeadline: fmtDate(row[5]),
        uploadDeadline: fmtDate(row[6]),
        dday: String(row[7] || ''),
        contact: String(row[8] || ''),
        refLink: String(row[9] || ''),
        memo: String(row[10] || '')
      });
    }

    // 월 순서대로 정렬
    var months = [];
    var keys = Object.keys(monthMap).sort();
    for (var k = 0; k < keys.length; k++) {
      if (keys[k] === '미정') continue;
      months.push({label: keys[k], items: monthMap[keys[k]]});
    }
    if (monthMap['미정']) months.push({label: '미정', items: monthMap['미정']});

    result.push({
      name: CHANNELS[c],
      color: colors[CHANNELS[c]] || '#666',
      months: months
    });
  }

  var now = new Date();
  return {
    channels: result,
    today: Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd'),
    updated: Utilities.formatDate(now, 'Asia/Seoul', 'yyyy.MM.dd HH:mm')
  };
}

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Seoul', 'M/d');
  return String(val);
}
