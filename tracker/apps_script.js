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

// ── 에디터에서 실행할 테스트 함수 (권한 승인용) ──
function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ["캠페인관리", "클릭로그", "월별정산", "고객DB", "쿠폰풀"];
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
