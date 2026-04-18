// UTM 트래커 + 쿠폰 관리 Apps Script
// 시트: https://docs.google.com/spreadsheets/d/1qRANdVfdy5Q05zI0yoSc15cPldX9m2OGXMg5tfyjZes
//
// 배포 방법:
// 1. 위 시트 열기 → 확장 프로그램 → Apps Script
// 2. 아래 코드 전체 복사 → 붙여넣기
// 3. 배포 → 새 배포 → 유형: 웹 앱
//    - 실행 주체: 본인 (ceo@flowmus.com)
//    - 액세스 권한: 모든 사용자
// 4. 배포 URL 복사 → go/index.html과 tracker/index.html의 TRACKER_API에 입력
//
// 시트 구조 (캠페인관리):
// A:캠페인slug  B:유형  C:타겟URL  D:플랫폼/브랜드  E:영상ID
// F:쿠폰코드  G:할인  H:RS%  I:상태  J:생성일  K:채널

// ── 플랫폼 약어 매핑 ──
var PLATFORM_ABBR = {
  "타이탄클래스": "TITAN", "무아클래스": "MUA", "무순클래스": "MUSUN",
  "패스트캠퍼스": "FAST", "클래스101": "C101", "콜로소": "COLOSO",
  "인프런": "INFLEARN", "직방프랜차이즈": "REAL",
  "BBQ": "BBQ", "교촌": "GYOCHON", "맘스터치": "MOMS",
  "bhc": "BHC", "메가커피": "MEGA", "컴포즈커피": "COMPOSE"
};

function getPlatformAbbr(platformName) {
  if (PLATFORM_ABBR[platformName]) return PLATFORM_ABBR[platformName];
  // 매핑에 없으면 앞 3글자 대문자
  return platformName.substring(0, 3).toUpperCase();
}

// ── 시트 메뉴 ─────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('채널 관리')
    .addItem('전체 보기', 'filterAll')
    .addSeparator()
    .addItem('사찍남만 보기', 'filterSajjiknam')
    .addItem('현장속으로만 보기', 'filterHyunjang')
    .addItem('머니로드만 보기', 'filterMoneyroad')
    .addSeparator()
    .addItem('채널별 현황 요약', 'showChannelSummary')
    .addSeparator()
    .addItem('새 캠페인 처리 (쿠폰 자동생성)', 'processNewCampaigns')
    .addSeparator()
    .addSubMenu(ui.createMenu('쿠폰 관리')
      .addItem('캠페인별 쿠폰 보기', 'filterCouponByCampaign')
      .addItem('전체 쿠폰 보기', 'filterCouponAll')
      .addSeparator()
      .addItem('종료 캠페인 쿠폰 삭제', 'deleteEndedCoupons')
      .addItem('선택한 쿠폰 삭제', 'deleteSelectedCoupons'))
    .addSeparator()
    .addItem('자동 트리거 설정', 'setupTriggers')
    .addToUi();
}

function filterAll() { applyChannelFilter(''); }
function filterSajjiknam() { applyChannelFilter('sajjiknam'); }
function filterHyunjang() { applyChannelFilter('hyunjang'); }
function filterMoneyroad() { applyChannelFilter('moneyroad'); }

function applyChannelFilter(channel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('캠페인관리');
  ss.setActiveSheet(sheet);

  var range = sheet.getDataRange();
  var filter = sheet.getFilter();
  if (filter) filter.remove();

  if (!channel) {
    SpreadsheetApp.getUi().alert('전체 캠페인을 표시합니다.');
    return;
  }

  filter = range.createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(channel)
    .build();
  filter.setColumnFilterCriteria(11, criteria); // K열 = 채널
}

function showChannelSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campSheet = ss.getSheetByName('캠페인관리');
  var campData = campSheet.getDataRange().getValues();
  var chSheet = ss.getSheetByName('채널관리');
  var chData = chSheet.getDataRange().getValues();

  var stats = {};
  for (var i = 1; i < campData.length; i++) {
    var ch = campData[i][10] || '미지정';
    if (!stats[ch]) stats[ch] = {total: 0, active: 0};
    stats[ch].total++;
    if (campData[i][8] === '활성') stats[ch].active++;
  }

  // 쿠폰 현황
  var poolSheet = ss.getSheetByName('쿠폰풀');
  var poolData = poolSheet.getDataRange().getValues();
  var couponStats = {};
  for (var j = 1; j < poolData.length; j++) {
    var slug = poolData[j][1];
    // slug에서 채널 찾기
    for (var k = 1; k < campData.length; k++) {
      if (campData[k][0] === slug) {
        var ch2 = campData[k][10] || '미지정';
        if (!couponStats[ch2]) couponStats[ch2] = {total: 0, used: 0, remaining: 0};
        couponStats[ch2].total++;
        if (poolData[j][2] === '발급완료') couponStats[ch2].used++;
        else couponStats[ch2].remaining++;
        break;
      }
    }
  }

  var msg = '--- 채널별 현황 ---\n\n';
  for (var c = 1; c < chData.length; c++) {
    var chSlug = chData[c][0];
    var chName = chData[c][1];
    var s = stats[chSlug] || {total: 0, active: 0};
    var cs = couponStats[chSlug] || {total: 0, used: 0, remaining: 0};
    msg += chName + ' (' + chSlug + ')\n';
    msg += '  캠페인: ' + s.active + '개 활성 / ' + s.total + '개 전체\n';
    msg += '  쿠폰: ' + cs.remaining + '개 잔여 / ' + cs.total + '개 전체 (발급 ' + cs.used + '개)\n\n';
  }

  SpreadsheetApp.getUi().alert(msg);
}

// ══════════════════════════════════════════════════
// ── 시트에서 새 캠페인 처리 (slug + 쿠폰 자동생성) ─
// ══════════════════════════════════════════════════
// 사용법: 캠페인관리 탭에 아래만 채우고 메뉴 클릭
//   B(유형) C(플랫폼 결제URL) D(플랫폼명) E(YouTube URL) G(할인) H(RS%) K(채널)
//   → A(slug) E(영상ID로 변환) F(쿠폰코드) I(상태) J(생성일) 자동 채움 + 쿠폰풀 100개 생성

function extractVideoId(input) {
  input = String(input || "").trim();
  // YouTube URL에서 ID 추출
  var match = input.match(/(?:v=|youtu\.be\/|\/embed\/|\/v\/)([a-zA-Z0-9_-]{11})/);
  if (match) return match[1];
  // 이미 11자 ID면 그대로
  if (/^[a-zA-Z0-9_-]{11}$/.test(input)) return input;
  return input;
}

function processNewCampaigns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  var poolSheet = ss.getSheetByName("쿠폰풀");

  var processed = 0;

  for (var i = 1; i < data.length; i++) {
    var slug = data[i][0];
    var channel = data[i][10] || "";
    var rawVideoInput = String(data[i][4] || "");

    // slug이 비어있고 채널+영상 정보가 있으면 새 캠페인
    if (slug || !channel || !rawVideoInput) continue;

    // YouTube URL → 영상 ID 추출
    var videoId = extractVideoId(rawVideoInput);
    var vid6 = videoId.substring(0, 6);
    var platform = data[i][3] || "";
    var platAbbr = getPlatformAbbr(platform);
    var newSlug = channel + "-" + vid6;

    // E열을 영상 ID로 교체 (URL → ID)
    var row = i + 1;
    if (videoId !== rawVideoInput) {
      sheet.getRange(row, 5).setValue(videoId);
    }

    // slug 중복 체크
    var counter = 2;
    var baseSlug = newSlug;
    var exists = true;
    while (exists) {
      exists = false;
      for (var j = 1; j < data.length; j++) {
        if (data[j][0] === newSlug) { exists = true; break; }
      }
      if (exists) {
        newSlug = baseSlug + "-" + counter;
        counter++;
      }
    }

    var couponCode = "FLOW-" + platAbbr + "-" + vid6 + "-01";
    var today = new Date().toISOString().split("T")[0];
    var row = i + 1; // 시트 행 번호 (1-based)

    // A: slug, F: 쿠폰코드, I: 상태, J: 생성일
    sheet.getRange(row, 1).setValue(newSlug);
    sheet.getRange(row, 6).setValue(couponCode);
    sheet.getRange(row, 9).setValue("활성");
    sheet.getRange(row, 10).setValue(today);

    // 쿠폰풀에 100개 생성
    var poolRows = [];
    for (var k = 1; k <= 100; k++) {
      var code = "FLOW-" + platAbbr + "-" + vid6 + "-" + ("0000" + k).slice(-4);
      poolRows.push([code, newSlug, "미발급", "", "", "", ""]);
    }
    poolSheet.getRange(poolSheet.getLastRow() + 1, 1, poolRows.length, 7).setValues(poolRows);

    // data 배열도 업데이트 (중복 체크용)
    data[i][0] = newSlug;
    processed++;
  }

  if (processed === 0) {
    SpreadsheetApp.getUi().alert("처리할 새 캠페인이 없습니다.\n\n캠페인관리 탭에 새 행을 추가하세요:\n  B(유형) C(플랫폼 결제URL) D(플랫폼명)\n  E(YouTube URL) G(할인) H(RS%) K(채널)\n\nA(slug)는 비워두세요 — 자동 생성됩니다.");
  } else {
    SpreadsheetApp.getUi().alert(processed + "개 캠페인 처리 완료!\n\n각 캠페인에 쿠폰 100개씩 생성됨.\n쿠폰풀 탭에서 확인하세요.");
  }
}

// ══════════════════════════════════════════════════
// ── 쿠폰 관리 (필터 / 삭제) ─────────────────────
// ══════════════════════════════════════════════════

// ── 캠페인별 쿠폰 보기 (필터) ──
function filterCouponByCampaign() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campSheet = ss.getSheetByName("캠페인관리");
  var campData = campSheet.getDataRange().getValues();

  // 캠페인 목록 만들기
  var campaigns = [];
  for (var i = 1; i < campData.length; i++) {
    if (campData[i][0]) {
      campaigns.push(campData[i][0] + " (" + campData[i][3] + ")");
    }
  }

  if (campaigns.length === 0) {
    SpreadsheetApp.getUi().alert("등록된 캠페인이 없습니다.");
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "캠페인별 쿠폰 보기",
    "번호를 입력하세요:\n\n" + campaigns.map(function(c, i) { return (i + 1) + ". " + c; }).join("\n"),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  var idx = parseInt(response.getResponseText()) - 1;
  if (isNaN(idx) || idx < 0 || idx >= campaigns.length) return;

  var selectedSlug = campData[idx + 1][0];

  // 쿠폰풀 탭으로 이동 + 필터 적용
  var poolSheet = ss.getSheetByName("쿠폰풀");
  ss.setActiveSheet(poolSheet);

  var filter = poolSheet.getFilter();
  if (filter) filter.remove();

  var range = poolSheet.getDataRange();
  filter = range.createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(selectedSlug)
    .build();
  filter.setColumnFilterCriteria(2, criteria); // B열 = 캠페인slug
}

// ── 전체 쿠폰 보기 (필터 해제) ──
function filterCouponAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var poolSheet = ss.getSheetByName("쿠폰풀");
  ss.setActiveSheet(poolSheet);

  var filter = poolSheet.getFilter();
  if (filter) filter.remove();

  SpreadsheetApp.getUi().alert("전체 쿠폰을 표시합니다.");
}

// ── 종료 캠페인 쿠폰 삭제 ──
function deleteEndedCoupons() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campSheet = ss.getSheetByName("캠페인관리");
  var campData = campSheet.getDataRange().getValues();
  var poolSheet = ss.getSheetByName("쿠폰풀");
  var poolData = poolSheet.getDataRange().getValues();

  // 종료된 캠페인 slug 수집
  var endedSlugs = {};
  for (var i = 1; i < campData.length; i++) {
    if (campData[i][8] === "종료") {
      endedSlugs[campData[i][0]] = campData[i][3];
    }
  }

  if (Object.keys(endedSlugs).length === 0) {
    SpreadsheetApp.getUi().alert("종료된 캠페인이 없습니다.\n\n캠페인관리 탭에서 I열(상태)을 '종료'로 변경한 후 다시 실행하세요.");
    return;
  }

  // 삭제 대상 확인
  var toDelete = [];
  var details = {};
  for (var j = 1; j < poolData.length; j++) {
    var slug = poolData[j][1];
    if (endedSlugs[slug]) {
      toDelete.push(j + 1); // 시트 행 번호
      details[slug] = (details[slug] || 0) + 1;
    }
  }

  if (toDelete.length === 0) {
    SpreadsheetApp.getUi().alert("삭제할 쿠폰이 없습니다.");
    return;
  }

  // 확인 프롬프트
  var msg = "아래 쿠폰을 삭제합니다:\n\n";
  for (var s in details) {
    msg += endedSlugs[s] + " (" + s + "): " + details[s] + "개\n";
  }
  msg += "\n총 " + toDelete.length + "개 삭제됩니다. 계속할까요?";

  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert("쿠폰 삭제 확인", msg, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // 아래에서 위로 삭제 (인덱스 밀림 방지)
  toDelete.sort(function(a, b) { return b - a; });
  for (var k = 0; k < toDelete.length; k++) {
    poolSheet.deleteRow(toDelete[k]);
  }

  SpreadsheetApp.getUi().alert("삭제 완료!\n\n총 " + toDelete.length + "개 쿠폰이 삭제되었습니다.");
}

// ── 선택한 쿠폰 삭제 ──
function deleteSelectedCoupons() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var poolSheet = ss.getSheetByName("쿠폰풀");

  if (ss.getActiveSheet().getName() !== "쿠폰풀") {
    SpreadsheetApp.getUi().alert("쿠폰풀 탭에서 삭제할 행을 선택한 후 실행하세요.");
    return;
  }

  var selection = poolSheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();

  if (startRow <= 1) {
    SpreadsheetApp.getUi().alert("헤더 행은 삭제할 수 없습니다.\n2행 이하를 선택하세요.");
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "선택 삭제 확인",
    startRow + "행 ~ " + (startRow + numRows - 1) + "행 (" + numRows + "개) 삭제할까요?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  poolSheet.deleteRows(startRow, numRows);
  SpreadsheetApp.getUi().alert(numRows + "개 행 삭제 완료!");
}

// ── 에디터에서 실행할 테스트 함수 (권한 승인용) ──
function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ["캠페인관리", "클릭로그", "월별정산", "고객DB", "쿠폰풀", "채널관리", "채널암호"];
  tabs.forEach(function(t) {
    var sheet = ss.getSheetByName(t);
    if (sheet) {
      Logger.log(t + ": " + (sheet.getDataRange().getValues().length - 1) + "행");
    } else {
      Logger.log(t + ": 탭 없음!");
    }
  });
  Logger.log("테스트 성공!");
}

// ── 클릭 로그 (POST) ──────────────────────────────
function doPost(e) {
  if (!e || !e.parameter) return jsonResponse({result: "no_params"});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = (e.parameter && e.parameter.action) || "log";

  // 클릭 로그 기록
  if (action === "log" || !e.parameter.action) {
    var sheet = ss.getSheetByName("클릭로그");
    var data = e.parameter;
    sheet.appendRow([
      new Date().toLocaleString("ko-KR", {timeZone: "Asia/Seoul"}),
      data.campaign || "",
      data.referrer || "",
      data.ua || ""
    ]);
    return jsonResponse({result: "logged"});
  }

  // 쿠폰 수령 (고객 정보 저장)
  if (action === "coupon_claim") {
    return saveCouponClaim(ss, e.parameter);
  }

  // 캠페인 추가 (관리자용)
  if (action === "add_campaign") {
    return addCampaign(ss, e.parameter);
  }

  // 로그인
  if (action === "login") {
    return handleLogin(ss, e.parameter);
  }

  return jsonResponse({result: "unknown_action"});
}

// ── 캠페인 조회 / 정산 데이터 (GET) ───────────────
function doGet(e) {
  if (!e || !e.parameter) return jsonResponse({result: "ok", message: "UTM Tracker API is running."});
  var action = e.parameter.action || "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 리다이렉트: 캠페인 slug로 타겟 URL 조회 (쿠폰 랜딩페이지용)
  if (action === "redirect") {
    var slug = e.parameter.slug || "";
    var campaign = findCampaign(ss, slug);
    if (campaign) {
      return jsonResponse({
        result: "found",
        url: campaign.url,
        coupon: campaign.coupon,
        discount: campaign.discount,
        platform: campaign.platform,
        type: campaign.type,
        slug: slug
      });
    }
    return jsonResponse({result: "not_found", slug: slug});
  }

  // 전체 캠페인 목록
  if (action === "campaigns") {
    var channel = e.parameter.channel || "";
    var campaigns = getAllCampaigns(ss, channel);
    return jsonResponse({result: "ok", campaigns: campaigns});
  }

  // 월별 정산 데이터
  if (action === "summary") {
    var channel = e.parameter.channel || "";
    var summary = getSummary(ss, channel);
    return jsonResponse({result: "ok", summary: summary});
  }

  // 클릭 통계 (캠페인별 클릭수 집계)
  if (action === "clicks") {
    var clicks = getClickStats(ss);
    return jsonResponse({result: "ok", clicks: clicks});
  }

  // 쿠폰 조회
  if (action === "coupon") {
    var code = (e.parameter.code || "").toUpperCase();
    var found = findByCoupon(ss, code);
    if (found) {
      return jsonResponse({result: "found", campaign: found});
    }
    return jsonResponse({result: "not_found", code: code});
  }

  // 쿠폰 잔여 수량
  if (action === "coupon_stock") {
    var slug = e.parameter.slug || "";
    var stock = getCouponStock(ss, slug);
    return jsonResponse({result: "ok", remaining: stock.remaining, total: stock.total});
  }

  // 토큰 검증
  if (action === "verify_token") {
    return handleVerifyToken(e.parameter);
  }

  return ContentService.createTextOutput("UTM Tracker API is running.");
}

// ── JSON 응답 헬퍼 ────────────────────────────────
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 캠페인 slug로 찾기 ────────────────────────────
function findCampaign(ss, slug) {
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === slug && data[i][8] === "활성") {
      return parseCampaignRow(data[i]);
    }
  }
  return null;
}

// ── 쿠폰 코드로 찾기 ─────────────────────────────
function findByCoupon(ss, code) {
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] && data[i][5].toString().toUpperCase() === code) {
      var campaign = parseCampaignRow(data[i]);
      // 월별정산에서 실적 가져오기
      var perf = getPerformance(ss, data[i][0]);
      campaign.clicks = perf.clicks;
      campaign.couponUsed = perf.couponUsed;
      campaign.revenue = perf.revenue;
      campaign.rsAmount = perf.rsAmount;
      campaign.convRate = perf.convRate;
      return campaign;
    }
  }
  return null;
}

// ── 캠페인 행 파싱 ────────────────────────────────
function parseCampaignRow(row) {
  return {
    slug: row[0],
    type: row[1],
    url: row[2],
    platform: row[3],
    videoId: row[4],
    coupon: row[5],
    discount: row[6],
    rs: row[7],
    status: row[8],
    created: row[9],
    channel: row[10] || ""
  };
}

// ── 캠페인별 실적 (월별정산 + 클릭로그) ───────────
function getPerformance(ss, slug) {
  // 클릭수: 클릭로그에서 집계
  var logSheet = ss.getSheetByName("클릭로그");
  var logData = logSheet.getDataRange().getValues();
  var clicks = 0;
  for (var i = 1; i < logData.length; i++) {
    if (logData[i][1] === slug) clicks++;
  }

  // 쿠폰사용/매출: 월별정산에서 합산
  var sumSheet = ss.getSheetByName("월별정산");
  var sumData = sumSheet.getDataRange().getValues();
  var couponUsed = 0, revenue = 0, rsAmount = 0;
  for (var j = 1; j < sumData.length; j++) {
    if (sumData[j][1] === slug) {
      couponUsed += Number(sumData[j][5]) || 0;
      revenue += Number(sumData[j][7]) || 0;
      rsAmount += Number(sumData[j][9]) || 0;
    }
  }

  var convRate = clicks > 0 ? (couponUsed / clicks * 100).toFixed(1) + "%" : "0%";

  return {
    clicks: clicks,
    couponUsed: couponUsed,
    revenue: revenue,
    rsAmount: rsAmount,
    convRate: convRate
  };
}

// ── 전체 캠페인 (채널 필터 지원) ──────────────────
function getAllCampaigns(ss, channelFilter) {
  var sheet = ss.getSheetByName("캠페인관리");
  var data = sheet.getDataRange().getValues();
  var clicks = getClickStats(ss);
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var channel = data[i][10] || "";
      if (channelFilter && channel !== channelFilter) continue;
      var camp = parseCampaignRow(data[i]);
      camp.clicks = clicks[camp.slug] || 0;
      result.push(camp);
    }
  }
  return result;
}

// ── 월별 정산 (채널 필터 지원) ────────────────────
function getSummary(ss, channelFilter) {
  var sheet = ss.getSheetByName("월별정산");
  var data = sheet.getDataRange().getValues();

  // 채널 필터가 있으면 해당 채널의 slug 목록 구하기
  var validSlugs = null;
  if (channelFilter) {
    validSlugs = {};
    var campSheet = ss.getSheetByName("캠페인관리");
    var campData = campSheet.getDataRange().getValues();
    for (var c = 1; c < campData.length; c++) {
      if (campData[c][10] === channelFilter) {
        validSlugs[campData[c][0]] = true;
      }
    }
  }

  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      if (validSlugs && !validSlugs[data[i][1]]) continue;
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

// ── 클릭 통계 (캠페인별 집계) ─────────────────────
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

// ── 1인 1쿠폰: 고유 쿠폰 배정 + 고객 저장 ───────
function saveCouponClaim(ss, params) {
  var slug = params.campaign || "";
  var name = params.name || "";
  var phone = params.phone || "";
  var email = params.email || "";
  var now = new Date().toLocaleString("ko-KR", {timeZone: "Asia/Seoul"});

  // 캠페인 정보
  var campaign = findCampaign(ss, slug);
  var channel = campaign ? campaign.channel : "";
  var platform = campaign ? campaign.platform : "";

  // ── 1인 1쿠폰: 쿠폰풀에서 미발급 코드 배정 ──
  var poolSheet = ss.getSheetByName("쿠폰풀");
  var poolData = poolSheet.getDataRange().getValues();
  var assignedCoupon = "";
  var assignedRow = -1;

  // 이미 발급받았는지 체크 (같은 연락처 + 같은 캠페인)
  for (var i = 1; i < poolData.length; i++) {
    if (poolData[i][1] === slug && poolData[i][2] === "발급완료" && poolData[i][5] === phone) {
      // 이미 발급받음 → 기존 코드 반환
      return jsonResponse({result: "saved", coupon: poolData[i][0]});
    }
  }

  // 미발급 코드 찾기
  for (var j = 1; j < poolData.length; j++) {
    if (poolData[j][1] === slug && poolData[j][2] === "미발급") {
      assignedCoupon = poolData[j][0];
      assignedRow = j + 1; // 시트 행 번호 (1-based)
      break;
    }
  }

  // 쿠폰 소진
  if (!assignedCoupon) {
    return jsonResponse({result: "no_stock"});
  }

  // 쿠폰풀 업데이트: 미발급 → 발급완료
  poolSheet.getRange(assignedRow, 3).setValue("발급완료");  // 상태
  poolSheet.getRange(assignedRow, 4).setValue(now);          // 발급일
  poolSheet.getRange(assignedRow, 5).setValue(name);         // 고객명
  poolSheet.getRange(assignedRow, 6).setValue(phone);        // 연락처
  poolSheet.getRange(assignedRow, 7).setValue(email);        // 이메일

  // 고객DB에도 저장
  var dbSheet = ss.getSheetByName("고객DB");
  dbSheet.appendRow([now, slug, channel, platform, assignedCoupon, name, phone, email, params.referrer || "", "쿠폰발급"]);

  // 클릭로그 기록
  var logSheet = ss.getSheetByName("클릭로그");
  logSheet.appendRow([now, slug, params.referrer || "", params.ua || ""]);

  return jsonResponse({result: "saved", coupon: assignedCoupon});
}

// ── 쿠폰 잔여 수량 ──────────────────────────────
function getCouponStock(ss, slug) {
  var sheet = ss.getSheetByName("쿠폰풀");
  var data = sheet.getDataRange().getValues();
  var total = 0, remaining = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === slug) {
      total++;
      if (data[i][2] === "미발급") remaining++;
    }
  }
  return {total: total, remaining: remaining};
}

// ── 캠페인 추가 ──────────────────────────────────
function addCampaign(ss, params) {
  var sheet = ss.getSheetByName("캠페인관리");

  var slug = params.slug || "";
  var type = params.type || "강의";
  var url = params.url || "";
  var platform = params.platform || "";
  var videoId = params.videoId || "";
  var coupon = params.coupon || "";
  var discount = params.discount || "";
  var rs = params.rs || (type === "강의" ? "20%" : "10%");
  var channel = params.channel || "";
  var today = new Date().toISOString().split("T")[0];

  // 중복 체크
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === slug) {
      return jsonResponse({result: "error", message: "중복된 slug: " + slug});
    }
    if (data[i][5] && data[i][5].toString().toUpperCase() === coupon.toUpperCase()) {
      return jsonResponse({result: "error", message: "중복된 쿠폰: " + coupon});
    }
  }

  sheet.appendRow([slug, type, url, platform, videoId, coupon, discount, rs, "활성", today, channel]);
  return jsonResponse({result: "added", slug: slug, coupon: coupon});
}

// ══════════════════════════════════════════════════
// ── 비밀번호 로그인 ──────────────────────────────
// ══════════════════════════════════════════════════

function handleLogin(ss, params) {
  var channel = params.channel || "";
  var pw = params.pw || "";

  var pwSheet = ss.getSheetByName("채널암호");
  if (!pwSheet) return jsonResponse({result: "error", message: "채널암호 탭이 없습니다"});

  var data = pwSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === channel && String(data[i][1]) === pw) {
      var token = generateToken(channel);
      return jsonResponse({result: "ok", token: token, channel: channel});
    }
  }
  return jsonResponse({result: "fail", message: "비밀번호가 일치하지 않습니다"});
}

function handleVerifyToken(params) {
  var token = params.token || "";
  var channel = verifyToken(token);
  if (channel) {
    return jsonResponse({result: "ok", channel: channel});
  }
  return jsonResponse({result: "invalid"});
}

function generateToken(channel) {
  var ts = new Date().getTime().toString();
  var secret = getSecretKey();
  var raw = channel + "." + ts;
  var sig = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw + "." + secret)
  ).substring(0, 24);
  return raw + "." + sig;
}

function verifyToken(token) {
  if (!token) return null;
  var parts = token.split(".");
  if (parts.length !== 3) return null;

  var channel = parts[0];
  var ts = parseInt(parts[1]);
  var sig = parts[2];

  // 24시간 만료
  var now = new Date().getTime();
  if (now - ts > 24 * 60 * 60 * 1000) return null;

  // 서명 검증
  var secret = getSecretKey();
  var raw = channel + "." + parts[1];
  var expectedSig = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw + "." + secret)
  ).substring(0, 24);

  if (sig !== expectedSig) return null;
  return channel;
}

function getSecretKey() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty("TOKEN_SECRET");
  if (!key) {
    key = Utilities.getUuid();
    props.setProperty("TOKEN_SECRET", key);
  }
  return key;
}

// ══════════════════════════════════════════════════
// ── 쿠폰 소진 알림 (매일 오전 9시 트리거) ───────
// ══════════════════════════════════════════════════

function dailyStockAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campSheet = ss.getSheetByName("캠페인관리");
  var campData = campSheet.getDataRange().getValues();
  var poolSheet = ss.getSheetByName("쿠폰풀");
  var poolData = poolSheet.getDataRange().getValues();
  var chSheet = ss.getSheetByName("채널관리");
  var chData = chSheet ? chSheet.getDataRange().getValues() : [];

  // 캠페인별 잔여 쿠폰 집계
  var stockBySlug = {};
  for (var i = 1; i < poolData.length; i++) {
    var slug = poolData[i][1];
    if (!stockBySlug[slug]) stockBySlug[slug] = {total: 0, remaining: 0};
    stockBySlug[slug].total++;
    if (poolData[i][2] === "미발급") stockBySlug[slug].remaining++;
  }

  // 10개 미만인 활성 캠페인
  var alerts = [];
  for (var j = 1; j < campData.length; j++) {
    if (campData[j][8] !== "활성") continue;
    var cSlug = campData[j][0];
    var stock = stockBySlug[cSlug];
    if (stock && stock.remaining < 10) {
      alerts.push({
        slug: cSlug,
        platform: campData[j][3],
        channel: campData[j][10] || "",
        remaining: stock.remaining,
        total: stock.total
      });
    }
  }

  if (alerts.length === 0) return; // 알림 불필요

  // HTML 이메일 본문
  var html = '<div style="font-family:-apple-system,sans-serif;max-width:600px;margin:0 auto;">';
  html += '<div style="background:#1a1d23;color:#fff;padding:20px 24px;border-radius:12px 12px 0 0;">';
  html += '<h2 style="margin:0;font-size:18px;">쿠폰 잔여 알림</h2></div>';
  html += '<div style="padding:20px 24px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px;">';
  html += '<p style="color:#64748b;margin:0 0 16px;">아래 캠페인의 쿠폰이 <strong>10개 미만</strong>입니다:</p>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:14px;">';
  html += '<tr style="background:#f8fafc;"><th style="padding:10px;text-align:left;border-bottom:2px solid #e2e8f0;">캠페인</th>';
  html += '<th style="padding:10px;text-align:left;border-bottom:2px solid #e2e8f0;">플랫폼</th>';
  html += '<th style="padding:10px;text-align:left;border-bottom:2px solid #e2e8f0;">채널</th>';
  html += '<th style="padding:10px;text-align:right;border-bottom:2px solid #e2e8f0;">잔여</th>';
  html += '<th style="padding:10px;text-align:right;border-bottom:2px solid #e2e8f0;">전체</th></tr>';

  alerts.forEach(function(a) {
    var color = a.remaining <= 3 ? '#dc2626' : '#ea580c';
    html += '<tr>';
    html += '<td style="padding:10px;border-bottom:1px solid #e2e8f0;">' + a.slug + '</td>';
    html += '<td style="padding:10px;border-bottom:1px solid #e2e8f0;">' + a.platform + '</td>';
    html += '<td style="padding:10px;border-bottom:1px solid #e2e8f0;">' + a.channel + '</td>';
    html += '<td style="padding:10px;text-align:right;border-bottom:1px solid #e2e8f0;color:' + color + ';font-weight:bold;">' + a.remaining + '개</td>';
    html += '<td style="padding:10px;text-align:right;border-bottom:1px solid #e2e8f0;">' + a.total + '개</td>';
    html += '</tr>';
  });

  html += '</table>';
  html += '<p style="margin:16px 0 0;font-size:12px;color:#94a3b8;">Flowmus MCN UTM 트래커 자동 발송</p>';
  html += '</div></div>';

  // 수신자: 관리자 + 채널관리 탭 이메일 (C열)
  var recipients = ['ceo@flowmus.com'];
  for (var k = 1; k < chData.length; k++) {
    var email = chData[k][2]; // C열 = 이메일
    if (email && recipients.indexOf(email) < 0) {
      recipients.push(email);
    }
  }

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: '[Flowmus MCN] 쿠폰 잔여 알림 — ' + alerts.length + '개 캠페인',
    htmlBody: html
  });
}

// ══════════════════════════════════════════════════
// ── 월별 RS 리포트 (매월 1일 오전 10시 트리거) ──
// ══════════════════════════════════════════════════

function monthlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chSheet = ss.getSheetByName("채널관리");
  if (!chSheet) return;
  var chData = chSheet.getDataRange().getValues();

  // 지난 달 계산
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var monthStr = lastMonth.getFullYear() + '-' + String(lastMonth.getMonth() + 1).padStart(2, '0');

  for (var c = 1; c < chData.length; c++) {
    var chSlug = chData[c][0];
    var chName = chData[c][1];
    var chEmail = chData[c][2]; // C열 = 이메일

    if (!chEmail) continue;

    // 해당 채널의 정산 데이터
    var summary = getSummary(ss, chSlug);
    var monthData = summary.filter(function(s) { return s.month === monthStr; });

    if (monthData.length === 0) continue;

    // 합계 계산
    var totalClicks = 0, totalCoupons = 0, totalRevenue = 0, totalRS = 0;
    monthData.forEach(function(s) {
      totalClicks += Number(s.clicks) || 0;
      totalCoupons += Number(s.couponUsed) || 0;
      totalRevenue += Number(s.revenue) || 0;
      totalRS += Number(s.rsAmount) || 0;
    });

    // HTML 이메일
    var html = '<div style="font-family:-apple-system,sans-serif;max-width:640px;margin:0 auto;background:#fff;">';

    // 헤더
    html += '<div style="background:#1a1d23;color:#fff;padding:24px 32px;border-radius:12px 12px 0 0;">';
    html += '<h1 style="margin:0;font-size:20px;">Flowmus MCN — ' + monthStr + ' RS 리포트</h1>';
    html += '<p style="margin:8px 0 0;font-size:14px;opacity:0.8;">' + chName + ' 채널</p>';
    html += '</div>';

    // 요약 수치
    html += '<div style="padding:24px 32px;border:1px solid #e2e8f0;border-top:none;">';
    html += '<table style="width:100%;margin-bottom:24px;" cellpadding="0" cellspacing="0"><tr>';
    html += '<td style="text-align:center;padding:16px;background:#f8fafc;border-radius:8px;"><div style="font-size:24px;font-weight:800;color:#2563EB;">' + totalClicks.toLocaleString() + '</div><div style="font-size:12px;color:#64748b;margin-top:4px;">총 클릭</div></td>';
    html += '<td style="width:12px;"></td>';
    html += '<td style="text-align:center;padding:16px;background:#f8fafc;border-radius:8px;"><div style="font-size:24px;font-weight:800;color:#1a1d23;">' + totalCoupons.toLocaleString() + '</div><div style="font-size:12px;color:#64748b;margin-top:4px;">쿠폰 사용</div></td>';
    html += '<td style="width:12px;"></td>';
    html += '<td style="text-align:center;padding:16px;background:#f8fafc;border-radius:8px;"><div style="font-size:24px;font-weight:800;color:#ea580c;">' + formatKRW(totalRevenue) + '</div><div style="font-size:12px;color:#64748b;margin-top:4px;">추적 매출</div></td>';
    html += '<td style="width:12px;"></td>';
    html += '<td style="text-align:center;padding:16px;background:rgba(22,163,74,0.06);border-radius:8px;border:1px solid rgba(22,163,74,0.15);"><div style="font-size:24px;font-weight:800;color:#16a34a;">' + formatKRW(totalRS) + '</div><div style="font-size:12px;color:#64748b;margin-top:4px;">RS 정산</div></td>';
    html += '</tr></table>';

    // 캠페인별 테이블
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#1a1d23;color:#fff;">';
    html += '<th style="padding:10px 12px;text-align:left;">캠페인</th>';
    html += '<th style="padding:10px 12px;text-align:left;">플랫폼</th>';
    html += '<th style="padding:10px 12px;text-align:right;">클릭</th>';
    html += '<th style="padding:10px 12px;text-align:right;">쿠폰</th>';
    html += '<th style="padding:10px 12px;text-align:right;">전환율</th>';
    html += '<th style="padding:10px 12px;text-align:right;">매출</th>';
    html += '<th style="padding:10px 12px;text-align:center;">RS%</th>';
    html += '<th style="padding:10px 12px;text-align:right;">RS 금액</th>';
    html += '</tr>';

    monthData.forEach(function(s, idx) {
      var bg = idx % 2 === 0 ? '#fff' : '#f8fafc';
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;">' + s.slug + '</td>';
      html += '<td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;">' + s.platform + '</td>';
      html += '<td style="padding:10px 12px;text-align:right;border-bottom:1px solid #e2e8f0;">' + Number(s.clicks).toLocaleString() + '</td>';
      html += '<td style="padding:10px 12px;text-align:right;border-bottom:1px solid #e2e8f0;">' + Number(s.couponUsed).toLocaleString() + '</td>';
      html += '<td style="padding:10px 12px;text-align:right;border-bottom:1px solid #e2e8f0;">' + (s.convRate || '-') + '</td>';
      html += '<td style="padding:10px 12px;text-align:right;border-bottom:1px solid #e2e8f0;">' + Number(s.revenue).toLocaleString() + '원</td>';
      html += '<td style="padding:10px 12px;text-align:center;border-bottom:1px solid #e2e8f0;">' + s.rs + '</td>';
      html += '<td style="padding:10px 12px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:700;color:#16a34a;">' + Number(s.rsAmount).toLocaleString() + '원</td>';
      html += '</tr>';
    });

    // 합계 행
    html += '<tr style="background:#f0fdf4;">';
    html += '<td colspan="2" style="padding:10px 12px;font-weight:700;">합계</td>';
    html += '<td style="padding:10px 12px;text-align:right;font-weight:700;">' + totalClicks.toLocaleString() + '</td>';
    html += '<td style="padding:10px 12px;text-align:right;font-weight:700;">' + totalCoupons.toLocaleString() + '</td>';
    html += '<td style="padding:10px 12px;text-align:right;font-weight:700;">' + (totalClicks > 0 ? (totalCoupons / totalClicks * 100).toFixed(1) + '%' : '-') + '</td>';
    html += '<td style="padding:10px 12px;text-align:right;font-weight:700;">' + totalRevenue.toLocaleString() + '원</td>';
    html += '<td style="padding:10px 12px;text-align:center;">-</td>';
    html += '<td style="padding:10px 12px;text-align:right;font-weight:700;color:#16a34a;">' + totalRS.toLocaleString() + '원</td>';
    html += '</tr>';
    html += '</table></div>';

    // 푸터
    html += '<div style="padding:16px 32px;background:#f8fafc;text-align:center;font-size:12px;color:#94a3b8;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px;">';
    html += 'Powered by Flowmus MCN | 문의: ceo@flowmus.com</div>';
    html += '</div>';

    MailApp.sendEmail({
      to: chEmail,
      subject: '[Flowmus MCN] ' + monthStr + ' RS 정산 리포트 — ' + chName,
      htmlBody: html
    });
  }
}

function formatKRW(amount) {
  if (amount >= 100000000) return (amount / 100000000).toFixed(1) + '억';
  if (amount >= 10000) return (amount / 10000).toFixed(0) + '만원';
  return amount.toLocaleString() + '원';
}

// ══════════════════════════════════════════════════
// ── 트리거 자동 설정 ────────────────────────────
// ══════════════════════════════════════════════════

function setupTriggers() {
  // 기존 트리거 제거 (중복 방지)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'dailyStockAlert' || fn === 'monthlyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 매일 오전 9시: 쿠폰 소진 알림
  ScriptApp.newTrigger('dailyStockAlert')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();

  // 매월 1일 오전 10시: RS 리포트
  ScriptApp.newTrigger('monthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .inTimezone('Asia/Seoul')
    .create();

  SpreadsheetApp.getUi().alert(
    '트리거 설정 완료!\n\n'
    + '1. 쿠폰 소진 알림: 매일 오전 9시\n'
    + '2. 월별 RS 리포트: 매월 1일 오전 10시\n\n'
    + '설정 확인: 확장 프로그램 → Apps Script → 트리거'
  );
}
