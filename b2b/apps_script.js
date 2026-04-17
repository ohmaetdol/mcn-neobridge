// B2B 미팅 요청 폼 → Google Sheets 연동
// 시트: https://docs.google.com/spreadsheets/d/1E-33NJAxQikgopXKMZg94T53vr6MizecJOahZo5sNho
//
// 배포 방법:
// 1. 위 시트 열기 → 확장 프로그램 → Apps Script
// 2. 아래 코드 전체 복사 → 붙여넣기
// 3. 배포 → 새 배포 → 유형: 웹 앱
//    - 실행 주체: 본인 (ceo@flowmus.com)
//    - 액세스 권한: 모든 사용자
// 4. 배포 → URL 복사 → index.html의 APPS_SCRIPT_URL_HERE 교체

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("미팅요청");
  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date().toLocaleString("ko-KR", {timeZone: "Asia/Seoul"}),
    data.name || "",
    data.company || "",
    data.phone || "",
    data.email || "",
    data.instructors || "",
    data.message || ""
  ]);

  // 알림 이메일 발송
  try {
    MailApp.sendEmail({
      to: "ceo@flowmus.com",
      subject: "[B2B 미팅 요청] " + data.company + " - " + data.name,
      body: "담당자: " + data.name + "\n" +
            "회사: " + data.company + "\n" +
            "연락처: " + data.phone + "\n" +
            "이메일: " + (data.email || "-") + "\n" +
            "강사 수: " + (data.instructors || "-") + "\n" +
            "문의: " + (data.message || "-")
    });
  } catch(err) {}

  return ContentService.createTextOutput(JSON.stringify({result: "success"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput("B2B Meeting Form API is running.");
}
