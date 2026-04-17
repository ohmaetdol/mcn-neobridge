// UTM 트래커 + 쿠폰 관리 Apps Script
// 시트: https://docs.google.com/spreadsheets/d/1qRANdVfdy5Q05zI0yoSc15cPldX9m2OGXMg5tfyjZes
//
// 배포 방법:
// 1. 위 시트 열기 → 확장 프로그램 → Apps Script
// 2. 아래 코드 전체 복사 → 붙여넣기
// 3. 배포 → 새 배포 → 유형: 웹 앱
//    - 실행 주체: 본인 (ceo@flowmus.com)
//    - 액세스 권한: 모든 사용자
// 4. 배포 URL 복사 → go/index.html과 tracker/index.html에 입력

// ── 클릭 로그 (POST) ──────────────────────────────
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("클릭로그");
  var data = e.parameter;

  sheet.appendRow([
    new Date().toLocaleString("ko-KR", {timeZone: "Asia/Seoul"}),
    data.campaign || "",
    data.referrer || "",
    data.ua || ""
  ]);

  return ContentService.createTextOutput(JSON.stringify({result: "logged"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 캠페인 조회 / 정산 데이터 (GET) ───────────────
function doGet(e) {
  var action = e.parameter.action || "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 리다이렉트: 캠페인 slug로 타겟 URL 조회
  if (action === "redirect") {
    var slug = e.parameter.c || "";
    var campaign = findCampaign(ss, slug);
    if (campaign) {
      return ContentService.createTextOutput(JSON.stringify({
        result: "found",
        url: campaign.url,
        coupon: campaign.coupon,
        slug: slug
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      result: "not_found", slug: slug
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 전체 캠페인 목록
  if (action === "campaigns") {
    var campaigns = getAllCampaigns(ss);
    return ContentService.createTextOutput(JSON.stringify({
      result: "ok", campaigns: campaigns
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 월별 정산 데이터
  if (action === "summary") {
    var summary = getSummary(ss);
    return ContentService.createTextOutput(JSON.stringify({
      result: "ok", summary: summary
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 클릭 통계 (캠페인별 클릭수 집계)
  if (action === "clicks") {
    var clicks = getClickStats(ss);
    return ContentService.createTextOutput(JSON.stringify({
      result: "ok", clicks: clicks
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("UTM Tracker API is running.");
}

// ── Helper: 캠페인 slug로 찾기 ────────────────────
function findCampaign(ss, slug) {
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === slug && data[i][8] === "활성") {
      return {
        slug: data[i][0],
        type: data[i][1],
        url: data[i][2],
        platform: data[i][3],
        videoId: data[i][4],
        coupon: data[i][5],
        discount: data[i][6],
        rs: data[i][7],
        status: data[i][8]
      };
    }
  }
  return null;
}

// ── Helper: 전체 캠페인 ───────────────────────────
function getAllCampaigns(ss) {
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({
        slug: data[i][0],
        type: data[i][1],
        url: data[i][2],
        platform: data[i][3],
        videoId: data[i][4],
        coupon: data[i][5],
        discount: data[i][6],
        rs: data[i][7],
        status: data[i][8],
        created: data[i][9]
      });
    }
  }
  return result;
}

// ── Helper: 월별 정산 ─────────────────────────────
function getSummary(ss) {
  var sheet = ss.getSheetByName("월별정산");
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({
        month: data[i][0],
        slug: data[i][1],
        type: data[i][2],
        platform: data[i][3],
        clicks: data[i][4],
        couponUsed: data[i][5],
        convRate: data[i][6],
        revenue: data[i][7],
        rs: data[i][8],
        rsAmount: data[i][9]
      });
    }
  }
  return result;
}

// ── Helper: 클릭 통계 (캠페인별 집계) ──────────────
function getClickStats(ss) {
  var sheet = ss.getSheetByName("클릭로그");
  var data = sheet.getDataRange().getValues();
  var stats = {};
  for (var i = 1; i < data.length; i++) {
    var slug = data[i][1];
    if (slug) {
      stats[slug] = (stats[slug] || 0) + 1;
    }
  }
  return stats;
}
