// 편집자 보드 — Apps Script 백엔드
// 시트: https://docs.google.com/spreadsheets/d/1RjiV35RUuWWiqbQkYugkgXfnoUpEzRifFPn6lTxZX4g
//
// 배포:
// 1. 위 시트 > 확장 프로그램 > Apps Script
// 2. 이 코드 전체 붙여넣기 > 저장
// 3. 배포 > 새 배포 > 웹 앱 > 실행 주체: 본인 / 액세스: 모든 사용자
// 4. 배포 URL을 editor/index.html의 API_URL에 입력

var SHEET_ID = '1RjiV35RUuWWiqbQkYugkgXfnoUpEzRifFPn6lTxZX4g';
var ADMIN_EMAIL = 'ceo@flowmus.com';

// ── 메뉴 ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('편집자 보드')
    .addItem('전체 현황 보기', 'showSummary')
    .addSeparator()
    .addItem('마감 임박 알림 발송', 'sendDeadlineAlerts')
    .addItem('자동 알림 트리거 설정 (매일 9시)', 'setupTriggers')
    .addToUi();
}

// ── 웹앱 GET ─────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'videos') return jsonResp(getVideos());
  if (action === 'my') return jsonResp(getMyWork(e.parameter.editor || ''));
  if (action === 'admin') return jsonResp(getAdminData());
  if (action === 'editors') return jsonResp(getEditors());

  return jsonResp({error: 'unknown action'});
}

// ── 웹앱 POST ────────────────────────────────────
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action || '';

  if (action === 'apply') return jsonResp(applyForVideo(data));
  if (action === 'submit_draft') return jsonResp(submitDraft(data));
  if (action === 'approve') return jsonResp(approveApplication(data));
  if (action === 'reject') return jsonResp(rejectApplication(data));
  if (action === 'feedback') return jsonResp(sendFeedback(data));
  if (action === 'complete') return jsonResp(completeVideo(data));

  return jsonResp({error: 'unknown action'});
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════
// ── 데이터 조회 ──────────────────────────────────
// ══════════════════════════════════════════════════

function getVideos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('영상목록').getDataRange().getValues();
  var videos = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    videos.push({
      id: data[i][0],
      channel: data[i][1],
      guest: data[i][2],
      title: data[i][3],
      dropbox: data[i][4],
      deadline: data[i][5] instanceof Date
        ? Utilities.formatDate(data[i][5], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(data[i][5]),
      status: data[i][6] || '대기'
    });
  }
  return {videos: videos};
}

function getMyWork(editorName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var appData = ss.getSheetByName('편집신청').getDataRange().getValues();
  var vidData = ss.getSheetByName('영상목록').getDataRange().getValues();

  // 영상 정보 맵
  var vidMap = {};
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0]) vidMap[vidData[i][0]] = {
      channel: vidData[i][1], guest: vidData[i][2], title: vidData[i][3],
      dropbox: vidData[i][4],
      deadline: vidData[i][5] instanceof Date
        ? Utilities.formatDate(vidData[i][5], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(vidData[i][5]),
      status: vidData[i][6]
    };
  }

  var myWork = [];
  for (var i = 1; i < appData.length; i++) {
    if (appData[i][2] !== editorName) continue;
    var vid = vidMap[appData[i][1]] || {};
    myWork.push({
      applyDate: appData[i][0] instanceof Date
        ? Utilities.formatDate(appData[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(appData[i][0]),
      videoId: appData[i][1],
      plan: appData[i][3],
      expectedDate: appData[i][4] instanceof Date
        ? Utilities.formatDate(appData[i][4], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(appData[i][4]),
      approvalStatus: appData[i][5] || '대기',
      memo: appData[i][7] || '',
      video: vid
    });
  }
  return {work: myWork};
}

function getEditors() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = ss.getSheetByName('편집자').getDataRange().getValues();
  var editors = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][3] === '비활성') continue;
    editors.push({name: data[i][0], specialty: data[i][2] || ''});
  }
  return {editors: editors};
}

function getAdminData() {
  var videos = getVideos();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var appData = ss.getSheetByName('편집신청').getDataRange().getValues();

  var applications = [];
  for (var i = 1; i < appData.length; i++) {
    if (!appData[i][1]) continue;
    applications.push({
      applyDate: appData[i][0] instanceof Date
        ? Utilities.formatDate(appData[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(appData[i][0]),
      videoId: appData[i][1],
      editor: appData[i][2],
      plan: appData[i][3],
      expectedDate: appData[i][4] instanceof Date
        ? Utilities.formatDate(appData[i][4], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(appData[i][4]),
      approvalStatus: appData[i][5] || '대기',
      approvalDate: appData[i][6] instanceof Date
        ? Utilities.formatDate(appData[i][6], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(appData[i][6]),
      memo: appData[i][7] || ''
    });
  }
  return {videos: videos.videos, applications: applications};
}

// ══════════════════════════════════════════════════
// ── 편집자 액션 ──────────────────────────────────
// ══════════════════════════════════════════════════

function applyForVideo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var appSheet = ss.getSheetByName('편집신청');
  var vidSheet = ss.getSheetByName('영상목록');

  var videoId = data.videoId;
  var editor = data.editor;
  var plan = data.plan;
  var expectedDate = data.expectedDate;

  // 중복 신청 체크
  var existing = appSheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][1] === videoId && existing[i][2] === editor) {
      return {success: false, message: '이미 신청한 영상입니다.'};
    }
  }

  // 신청 추가
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  appSheet.appendRow([today, videoId, editor, plan, expectedDate, '대기', '', '']);

  // 영상 상태 업데이트
  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId && vidData[i][6] === '대기') {
      vidSheet.getRange(i + 1, 7).setValue('신청됨');
      break;
    }
  }

  // 관리자에게 이메일
  var vidInfo = getVideoInfo(ss, videoId);
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[편집 신청] ' + editor + ' → ' + (vidInfo.title || videoId),
      htmlBody: buildApplyEmail(editor, videoId, plan, expectedDate, vidInfo)
    });
  } catch(e) {}

  return {success: true, message: '신청 완료! 승인을 기다려주세요.'};
}

function submitDraft(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var vidSheet = ss.getSheetByName('영상목록');
  var videoId = data.videoId;
  var editor = data.editor;
  var draftLink = data.draftLink || '';

  // 영상 상태 업데이트
  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId) {
      vidSheet.getRange(i + 1, 7).setValue('초안제출');
      break;
    }
  }

  // 편집신청 메모에 초안 링크 추가
  var appSheet = ss.getSheetByName('편집신청');
  var appData = appSheet.getDataRange().getValues();
  for (var i = 1; i < appData.length; i++) {
    if (appData[i][1] === videoId && appData[i][2] === editor) {
      var memo = appData[i][7] || '';
      memo += (memo ? '\n' : '') + '초안: ' + draftLink;
      appSheet.getRange(i + 1, 8).setValue(memo);
      break;
    }
  }

  // 관리자에게 이메일
  var vidInfo = getVideoInfo(ss, videoId);
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[초안 제출] ' + editor + ' → ' + (vidInfo.title || videoId),
      htmlBody: '<h2>' + editor + '님이 초안을 제출했습니다</h2>' +
        '<p>영상: ' + (vidInfo.title || videoId) + '</p>' +
        (draftLink ? '<p>초안 링크: <a href="' + draftLink + '">' + draftLink + '</a></p>' : '') +
        '<p>편집자 보드에서 확인해주세요.</p>'
    });
  } catch(e) {}

  return {success: true, message: '초안 제출 완료!'};
}

// ══════════════════════════════════════════════════
// ── 관리자 액션 ──────────────────────────────────
// ══════════════════════════════════════════════════

function approveApplication(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var appSheet = ss.getSheetByName('편집신청');
  var vidSheet = ss.getSheetByName('영상목록');
  var videoId = data.videoId;
  var editor = data.editor;

  // 승인 상태 업데이트
  var appData = appSheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  for (var i = 1; i < appData.length; i++) {
    if (appData[i][1] === videoId && appData[i][2] === editor) {
      appSheet.getRange(i + 1, 6).setValue('승인');
      appSheet.getRange(i + 1, 7).setValue(today);
      break;
    }
  }

  // 영상 상태 → 편집중
  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId) {
      vidSheet.getRange(i + 1, 7).setValue('편집중');
      break;
    }
  }

  // 편집자에게 승인 + 코치 가이드 이메일
  var editorEmail = getEditorEmail(ss, editor);
  var vidInfo = getVideoInfo(ss, videoId);
  var guide = getCoachGuide(ss, vidInfo.channel);

  if (editorEmail) {
    try {
      MailApp.sendEmail({
        to: editorEmail,
        subject: '[승인] ' + (vidInfo.title || videoId) + ' 편집 시작하세요!',
        htmlBody: buildApprovalEmail(editor, vidInfo, guide)
      });
    } catch(e) {}
  }

  return {success: true, message: editor + ' 승인 완료'};
}

function rejectApplication(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var appSheet = ss.getSheetByName('편집신청');
  var vidSheet = ss.getSheetByName('영상목록');
  var videoId = data.videoId;
  var editor = data.editor;
  var reason = data.reason || '';

  // 반려
  var appData = appSheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  for (var i = 1; i < appData.length; i++) {
    if (appData[i][1] === videoId && appData[i][2] === editor) {
      appSheet.getRange(i + 1, 6).setValue('반려');
      appSheet.getRange(i + 1, 7).setValue(today);
      appSheet.getRange(i + 1, 8).setValue(reason);
      break;
    }
  }

  // 영상 상태 → 대기 (다른 편집자가 신청 가능)
  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId) {
      // 다른 승인된 신청이 있는지 확인
      var hasOther = false;
      for (var j = 1; j < appData.length; j++) {
        if (appData[j][1] === videoId && appData[j][2] !== editor && appData[j][5] === '승인') {
          hasOther = true; break;
        }
      }
      if (!hasOther) vidSheet.getRange(i + 1, 7).setValue('대기');
      break;
    }
  }

  // 편집자에게 반려 이메일
  var editorEmail = getEditorEmail(ss, editor);
  if (editorEmail) {
    try {
      MailApp.sendEmail({
        to: editorEmail,
        subject: '[반려] ' + videoId + ' 편집 신청',
        htmlBody: '<h2>편집 신청이 반려되었습니다</h2>' +
          '<p>영상: ' + videoId + '</p>' +
          (reason ? '<p>사유: ' + reason + '</p>' : '') +
          '<p>다른 영상에 다시 신청해주세요!</p>'
      });
    } catch(e) {}
  }

  return {success: true, message: editor + ' 반려 완료'};
}

function sendFeedback(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var videoId = data.videoId;
  var editor = data.editor;
  var feedback = data.feedback || '';

  // 영상 상태 → 수정중
  var vidSheet = ss.getSheetByName('영상목록');
  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId) {
      vidSheet.getRange(i + 1, 7).setValue('수정중');
      break;
    }
  }

  // 메모에 피드백 추가
  var appSheet = ss.getSheetByName('편집신청');
  var appData = appSheet.getDataRange().getValues();
  for (var i = 1; i < appData.length; i++) {
    if (appData[i][1] === videoId && appData[i][2] === editor) {
      var memo = appData[i][7] || '';
      memo += (memo ? '\n' : '') + '피드백: ' + feedback;
      appSheet.getRange(i + 1, 8).setValue(memo);
      break;
    }
  }

  // 편집자에게 이메일
  var editorEmail = getEditorEmail(ss, editor);
  if (editorEmail) {
    try {
      MailApp.sendEmail({
        to: editorEmail,
        subject: '[피드백] ' + videoId + ' 수정 요청',
        htmlBody: '<h2>피드백이 도착했습니다</h2>' +
          '<p>영상: ' + videoId + '</p>' +
          '<div style="background:#f5f5f5;padding:16px;border-radius:8px;margin:12px 0;">' + feedback + '</div>' +
          '<p>수정 후 다시 초안을 제출해주세요!</p>'
      });
    } catch(e) {}
  }

  return {success: true, message: '피드백 발송 완료'};
}

function completeVideo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var vidSheet = ss.getSheetByName('영상목록');
  var videoId = data.videoId;

  var vidData = vidSheet.getDataRange().getValues();
  for (var i = 1; i < vidData.length; i++) {
    if (vidData[i][0] === videoId) {
      vidSheet.getRange(i + 1, 7).setValue('완료');
      break;
    }
  }

  return {success: true, message: videoId + ' 완료 처리됨'};
}

// ══════════════════════════════════════════════════
// ── 헬퍼 ─────────────────────────────────────────
// ══════════════════════════════════════════════════

function getVideoInfo(ss, videoId) {
  var data = ss.getSheetByName('영상목록').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === videoId) {
      return {
        channel: data[i][1], guest: data[i][2], title: data[i][3],
        dropbox: data[i][4],
        deadline: data[i][5] instanceof Date
          ? Utilities.formatDate(data[i][5], 'Asia/Seoul', 'yyyy-MM-dd')
          : String(data[i][5])
      };
    }
  }
  return {};
}

function getEditorEmail(ss, name) {
  var data = ss.getSheetByName('편집자').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) return data[i][1];
  }
  return '';
}

function getCoachGuide(ss, channel) {
  var data = ss.getSheetByName('편집코치가이드').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === channel) return data[i][1];
  }
  return '편집 가이드가 아직 등록되지 않았습니다.';
}

// ── 이메일 템플릿 ────────────────────────────────

function buildApplyEmail(editor, videoId, plan, expectedDate, vidInfo) {
  return '<div style="font-family:-apple-system,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#0A1628;color:#fff;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;">편집 신청이 들어왔습니다</h2></div>' +
    '<div style="padding:20px;border:1px solid #e5e5e5;border-top:none;border-radius:0 0 8px 8px;">' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:8px 0;color:#888;width:100px;">편집자</td><td style="padding:8px 0;font-weight:700;">' + editor + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">영상</td><td style="padding:8px 0;">' + (vidInfo.title || videoId) + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">채널</td><td style="padding:8px 0;">' + (vidInfo.channel || '-') + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">예상 완료일</td><td style="padding:8px 0;">' + expectedDate + '</td></tr>' +
    '</table>' +
    '<div style="background:#f8f9fa;padding:16px;border-radius:8px;margin:16px 0;">' +
    '<strong>편집 기획서:</strong><br><br>' + plan.replace(/\n/g, '<br>') + '</div>' +
    '<p style="color:#888;font-size:13px;">편집자 보드 관리자 뷰에서 승인/반려해주세요.</p>' +
    '</div></div>';
}

function buildApprovalEmail(editor, vidInfo, guide) {
  return '<div style="font-family:-apple-system,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#217a38;color:#fff;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;">편집이 승인되었습니다!</h2></div>' +
    '<div style="padding:20px;border:1px solid #e5e5e5;border-top:none;border-radius:0 0 8px 8px;">' +
    '<p>' + editor + '님, 아래 영상의 편집을 시작해주세요.</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:8px 0;color:#888;width:100px;">영상</td><td style="padding:8px 0;font-weight:700;">' + (vidInfo.title || '-') + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">채널</td><td style="padding:8px 0;">' + (vidInfo.channel || '-') + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">출연자</td><td style="padding:8px 0;">' + (vidInfo.guest || '-') + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:#888;">마감일</td><td style="padding:8px 0;font-weight:700;color:#c01a1a;">' + (vidInfo.deadline || '-') + '</td></tr>' +
    '</table>' +
    (vidInfo.dropbox ? '<p><a href="' + vidInfo.dropbox + '" style="display:inline-block;background:#0061fe;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;margin:12px 0;">드롭박스에서 원본 다운로드</a></p>' : '') +
    '<div style="background:#f0fdf4;padding:16px;border-radius:8px;margin:16px 0;border-left:4px solid #217a38;">' +
    '<strong>편집 가이드:</strong><br><br>' + guide.replace(/\n/g, '<br>') + '</div>' +
    '<p style="color:#888;font-size:13px;">편집 완료 후 편집자 보드에서 초안을 제출해주세요.</p>' +
    '</div></div>';
}

// ── 마감 알림 ────────────────────────────────────

function sendDeadlineAlerts() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var vidData = ss.getSheetByName('영상목록').getDataRange().getValues();
  var appData = ss.getSheetByName('편집신청').getDataRange().getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var sent = 0;

  for (var i = 1; i < vidData.length; i++) {
    var videoId = vidData[i][0];
    var deadline = vidData[i][5];
    var status = vidData[i][6];
    if (!videoId || !deadline || !(deadline instanceof Date)) continue;
    if (status === '완료' || status === '대기') continue;

    var dl = new Date(deadline);
    dl.setHours(0, 0, 0, 0);
    var dDay = Math.ceil((dl - today) / (1000 * 60 * 60 * 24));

    if (dDay !== 3 && dDay !== 1 && dDay !== 0) continue;

    // 이 영상의 담당 편집자 찾기
    for (var j = 1; j < appData.length; j++) {
      if (appData[j][1] === videoId && appData[j][5] === '승인') {
        var editor = appData[j][2];
        var email = getEditorEmail(ss, editor);
        if (!email) continue;

        var subject = dDay === 0
          ? '[오늘 마감] ' + (vidData[i][3] || videoId)
          : '[D-' + dDay + '] ' + (vidData[i][3] || videoId) + ' 마감 임박';

        try {
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: '<h2>' + (dDay === 0 ? '오늘이 마감일입니다!' : 'D-' + dDay + ' 마감 임박') + '</h2>' +
              '<p>영상: ' + (vidData[i][3] || videoId) + '</p>' +
              '<p>마감일: ' + Utilities.formatDate(dl, 'Asia/Seoul', 'yyyy-MM-dd (E)') + '</p>' +
              '<p>편집자 보드에서 초안을 제출해주세요.</p>'
          });
          sent++;
        } catch(e) {}
      }
    }
  }

  try {
    SpreadsheetApp.getUi().alert('마감 알림 ' + sent + '건 발송');
  } catch(e) {}
}

function setupTriggers() {
  // 기존 트리거 제거
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDeadlineAlerts') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 매일 09:00
  ScriptApp.newTrigger('sendDeadlineAlerts')
    .timeBased().atHour(9).everyDays(1).inTimezone('Asia/Seoul').create();

  try {
    SpreadsheetApp.getUi().alert('매일 오전 9시 마감 알림 트리거 설정 완료');
  } catch(e) {}
}

// ── 시트 메뉴: 전체 현황 ─────────────────────────
function showSummary() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var vidData = ss.getSheetByName('영상목록').getDataRange().getValues();

  var stats = {};
  for (var i = 1; i < vidData.length; i++) {
    if (!vidData[i][0]) continue;
    var s = vidData[i][6] || '대기';
    stats[s] = (stats[s] || 0) + 1;
  }

  var msg = '영상 현황:\n';
  for (var key in stats) {
    msg += '  ' + key + ': ' + stats[key] + '개\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}
