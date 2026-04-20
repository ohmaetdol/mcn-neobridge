// 스케줄 웹앱 연동 + 메뉴 — Apps Script
// 시트: https://docs.google.com/spreadsheets/d/1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s
//
// 배포 방법:
// 1. 위 시트 > 확장 프로그램 > Apps Script
// 2. 이 코드 전체 붙여넣기 > 저장
// 3. 배포 > 배포 관리 > 연필 아이콘 > 새 버전 > 배포
// (최초 배포 시: 배포 > 새 배포 > 웹 앱 > 실행 주체: 본인 / 액세스: 모든 사용자)

var SHEET_ID = '1BGmP1bYCYKxTm2QtKnoNtH_avRbPgO1M9-wpPqNVb-s';

// ── 메뉴 ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('스케줄 관리')
    .addItem('새 강의 추가', 'showAddDialog')
    .addItem('전체 현황', 'showSummary')
    .addSeparator()
    .addItem('웹앱 열기', 'openWebApp')
    .addToUi();
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
    '.row{display:flex;gap:12px}' +
    '.row>div{flex:1}' +
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
    '<div class="row">' +
    '<div><label>연락처</label><input type="tel" id="contact" placeholder="010-0000-0000"></div>' +
    '</div>' +
    '<label>참고링크</label>' +
    '<input type="url" id="refLink" placeholder="https://">' +
    '<label>메모</label>' +
    '<input type="text" id="memo" placeholder="메모 (선택)">' +
    '<button class="btn btn-primary" onclick="submit()">추가</button>' +
    '<div id="msg" class="msg"></div>' +
    '<script>' +
    'function submit(){' +
    '  var d={' +
    '    channel:document.getElementById("channel").value,' +
    '    lectureDate:document.getElementById("lectureDate").value,' +
    '    instructor:document.getElementById("instructor").value,' +
    '    status:document.getElementById("status").value,' +
    '    contact:document.getElementById("contact").value,' +
    '    refLink:document.getElementById("refLink").value,' +
    '    memo:document.getElementById("memo").value' +
    '  };' +
    '  if(!d.instructor){showMsg("강사명을 입력하세요","err");return}' +
    '  google.script.run.withSuccessHandler(function(r){' +
    '    if(r.success){showMsg(r.message,"ok");setTimeout(function(){google.script.host.close()},1500)}' +
    '    else{showMsg(r.message,"err")}' +
    '  }).withFailureHandler(function(e){showMsg("오류: "+e.message,"err")}).addLecture(d);' +
    '}' +
    'function showMsg(t,type){var m=document.getElementById("msg");m.textContent=t;m.className="msg msg-"+type}' +
    '</script>';
}

function addLecture(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(data.channel);
  if (!sheet) return {success: false, message: '채널 탭을 찾을 수 없습니다.'};

  var all = sheet.getDataRange().getValues();

  // 빈 데이터 행 찾기 (No는 있지만 강사명이 비어있는 행)
  var targetRow = -1;
  var targetMonth = '';
  if (data.lectureDate) {
    var dt = new Date(data.lectureDate);
    targetMonth = dt.getFullYear() + '년 ' + (dt.getMonth() + 1) + '월';
  }

  // 적절한 월 구간에서 빈 행 찾기
  var inTargetMonth = false;
  for (var i = 2; i < all.length; i++) {
    var cell3 = String(all[i][3] || '');
    if (cell3.indexOf('──') === 0) {
      var monthLabel = cell3.replace(/──/g, '').trim();
      inTargetMonth = (monthLabel === targetMonth);
      continue;
    }
    // 빈 행이면 사용 (No만 있고 강사명 없는 행)
    if (all[i][0] && !all[i][3] && (inTargetMonth || !targetMonth)) {
      targetRow = i + 1; // 1-indexed
      break;
    }
  }

  // 빈 행이 없으면 마지막에 추가
  if (targetRow === -1) {
    targetRow = all.length + 1;
    var lastNo = 0;
    for (var i = all.length - 1; i >= 2; i--) {
      if (all[i][0] && typeof all[i][0] === 'number') { lastNo = all[i][0]; break; }
    }
    sheet.getRange(targetRow, 1).setValue(lastNo + 1);
  }

  // 데이터 입력 (B: 날짜, D: 강사명, E: 상태, I: 연락처, J: 참고, K: 메모)
  if (data.lectureDate) sheet.getRange(targetRow, 2).setValue(new Date(data.lectureDate));
  sheet.getRange(targetRow, 4).setValue(data.instructor);
  sheet.getRange(targetRow, 5).setValue(data.status);
  if (data.contact) sheet.getRange(targetRow, 9).setValue(data.contact);
  if (data.refLink) sheet.getRange(targetRow, 10).setValue(data.refLink);
  if (data.memo) sheet.getRange(targetRow, 11).setValue(data.memo);

  // 수식 설정 (C: 요일, F: 편집마감, G: 업로드마감, H: D-Day)
  sheet.getRange(targetRow, 3).setFormula('=IF(B' + targetRow + '="","",CHOOSE(WEEKDAY(B' + targetRow + ',1),"일","월","화","수","목","금","토"))');
  sheet.getRange(targetRow, 6).setFormula('=IF(B' + targetRow + '="","",B' + targetRow + '+9)');
  sheet.getRange(targetRow, 7).setFormula('=IF(B' + targetRow + '="","",B' + targetRow + '+14)');
  sheet.getRange(targetRow, 8).setFormula('=IF(G' + targetRow + '="","",IF(G' + targetRow + '-TODAY()>=0,"D-"&(G' + targetRow + '-TODAY()),"D+"&(TODAY()-G' + targetRow + ')))');

  return {success: true, message: data.instructor + ' 강의가 추가되었습니다!'};
}

// ── 전체 현황 ─────────────────────────────────────
function showSummary() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var channels = ['사장찍어주는남자', '부업소개하는남자'];
  var msg = '';

  for (var c = 0; c < channels.length; c++) {
    var sheet = ss.getSheetByName(channels[c]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();

    var total = 0, statusCounts = {};
    for (var i = 2; i < data.length; i++) {
      if (!data[i][3] || String(data[i][3]).indexOf('──') === 0) continue;
      total++;
      var st = data[i][4] || '미정';
      statusCounts[st] = (statusCounts[st] || 0) + 1;
    }

    msg += '[ ' + channels[c] + ' ] — ' + total + '건\n';
    for (var st in statusCounts) {
      msg += '  ' + st + ': ' + statusCounts[st] + '건\n';
    }
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
// ── 웹앱 API (아래는 기존 코드) ──────────────────
// ══════════════════════════════════════════════════

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

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'M/d');
  }
  return String(val);
}
