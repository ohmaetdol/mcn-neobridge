/**
 * Flowmus 크리에이터 리드 대시보드 API
 *
 * 이 스크립트를 "크리에이터 리드 관리 대시보드" 구글시트의
 * 확장 프로그램 > Apps Script에 붙여넣기 하세요.
 *
 * 배포: 배포 > 새 배포 > 웹 앱 > 모든 사용자 접근
 */

// ── 설정 ──
const LEAD_SHEET_ID = '1LYrBEoJVJejXskhfkEyy7Dn_73HjpUTOAlZZm6RSd5Y';
const LEAD_SHEET_NAME = '시트1';
const FREE_LEADS = 50;
const PRICE_PER_LEAD = 500;

// ── 계정 (MVP: 여기에 직접 등록) ──
const ACCOUNTS = {
  'aisaju': { password: 'flowmus2026', channel: 'AI사주 당근마켓 + 자사몰 강의', partner: '십억채널 × 사장찍어주는남자' },
  // 새 클라이언트 추가 시 여기에 한 줄 추가
};

function doGet(e) {
  var params = e.parameter;
  var action = params.action || 'data';

  // CORS 헤더 포함 JSON 응답
  if (action === 'login') {
    return handleLogin(params.id, params.pw);
  }

  if (action === 'data') {
    return handleData(params.key);
  }

  return jsonResponse({ error: '잘못된 요청' }, 400);
}

function handleLogin(id, pw) {
  if (!id || !pw) return jsonResponse({ error: '아이디와 비밀번호를 입력하세요' });

  var account = ACCOUNTS[id];
  if (!account || account.password !== pw) {
    return jsonResponse({ error: '아이디 또는 비밀번호가 올바르지 않습니다' });
  }

  // 간단한 토큰 생성
  var token = Utilities.base64Encode(id + ':' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + pw + 'flowmus').toString());

  return jsonResponse({
    success: true,
    token: token,
    channel: account.channel,
    partner: account.partner
  });
}

function handleData(key) {
  if (!key) return jsonResponse({ error: '인증이 필요합니다' });

  // 토큰 검증
  var valid = false;
  var channelInfo = null;
  for (var id in ACCOUNTS) {
    var acc = ACCOUNTS[id];
    var expected = Utilities.base64Encode(id + ':' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + acc.password + 'flowmus').toString());
    if (expected === key) {
      valid = true;
      channelInfo = { id: id, channel: acc.channel, partner: acc.partner };
      break;
    }
  }

  if (!valid) return jsonResponse({ error: '인증 실패' });

  // 리드 데이터 가져오기
  try {
    var ss = SpreadsheetApp.openById(LEAD_SHEET_ID);
    var sheet = ss.getSheetByName(LEAD_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);

    // 날짜 인덱스
    var dateIdx = 0;  // 접수일시
    var nameIdx = 1;  // 이름
    var interestIdx = 4;  // 관심분야
    var utmSourceIdx = 5; // utm_source

    var now = new Date();
    var thisMonth = now.getMonth();
    var thisYear = now.getFullYear();

    // 월별 집계
    var monthlyData = {};
    var utmCounts = {};
    var recentLeads = [];
    var totalLeads = rows.length;
    var thisMonthLeads = 0;
    var lastMonthLeads = 0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var dateVal = row[dateIdx];
      var d = new Date(dateVal);
      var monthKey = d.getFullYear() + '.' + String(d.getMonth() + 1).padStart(2, '0');

      monthlyData[monthKey] = (monthlyData[monthKey] || 0) + 1;

      if (d.getMonth() === thisMonth && d.getFullYear() === thisYear) {
        thisMonthLeads++;
      }
      if (d.getMonth() === (thisMonth - 1 + 12) % 12) {
        lastMonthLeads++;
      }

      // UTM 집계
      var source = row[utmSourceIdx] || '기타';
      utmCounts[source] = (utmCounts[source] || 0) + 1;

      // 최근 리드 (최대 10개)
      recentLeads.push({
        date: formatDate(d),
        name: maskName(row[nameIdx] || ''),
        interest: row[interestIdx] || '',
        source: source
      });
    }

    // 최근순 정렬
    recentLeads.sort(function(a, b) { return b.date > a.date ? 1 : -1; });
    recentLeads = recentLeads.slice(0, 10);

    // UTM 정렬 (많은 순)
    var utmArray = [];
    for (var src in utmCounts) {
      utmArray.push({ name: src, count: utmCounts[src] });
    }
    utmArray.sort(function(a, b) { return b.count - a.count; });

    // 전월 대비 변화율
    var changeRate = lastMonthLeads > 0 ? Math.round((thisMonthLeads - lastMonthLeads) / lastMonthLeads * 100) : 0;

    // 일 평균
    var dayOfMonth = now.getDate();
    var dailyAvg = dayOfMonth > 0 ? (thisMonthLeads / dayOfMonth).toFixed(1) : 0;

    // 과금 계산
    var billableLeads = Math.max(0, thisMonthLeads - FREE_LEADS);
    var billingAmount = billableLeads * PRICE_PER_LEAD;

    // 월별 차트 데이터 (올해 1~12월)
    var chartData = [];
    for (var m = 1; m <= 12; m++) {
      var key = thisYear + '.' + String(m).padStart(2, '0');
      chartData.push({ month: m + '월', count: monthlyData[key] || 0 });
    }

    // 정산 내역 (월별)
    var billingHistory = [];
    var sortedMonths = Object.keys(monthlyData).sort().reverse();
    for (var j = 0; j < sortedMonths.length; j++) {
      var mk = sortedMonths[j];
      var cnt = monthlyData[mk];
      var bill = Math.max(0, cnt - FREE_LEADS);
      var isCurrent = (mk === thisYear + '.' + String(thisMonth + 1).padStart(2, '0'));
      billingHistory.push({
        month: mk,
        leads: cnt,
        billable: bill,
        amount: bill * PRICE_PER_LEAD,
        status: isCurrent ? '진행중' : '입금완료'
      });
    }

    return jsonResponse({
      success: true,
      channel: channelInfo.channel,
      partner: channelInfo.partner,
      kpi: {
        thisMonth: thisMonthLeads,
        total: totalLeads,
        changeRate: changeRate,
        dailyAvg: parseFloat(dailyAvg)
      },
      billing: {
        totalLeads: thisMonthLeads,
        freeLeads: FREE_LEADS,
        billableLeads: billableLeads,
        pricePerLead: PRICE_PER_LEAD,
        amount: billingAmount
      },
      chart: chartData,
      utm: utmArray,
      recentLeads: recentLeads,
      billingHistory: billingHistory,
      lastUpdate: new Date().toISOString()
    });

  } catch (err) {
    return jsonResponse({ error: '데이터를 불러올 수 없습니다: ' + err.message });
  }
}

function maskName(name) {
  if (!name || name.length < 2) return name;
  return name.charAt(0) + '*' + name.substring(2);
}

function formatDate(d) {
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  var hh = String(d.getHours()).padStart(2, '0');
  var mi = String(d.getMinutes()).padStart(2, '0');
  return mm + '.' + dd + ' ' + hh + ':' + mi;
}

function jsonResponse(data, code) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 테스트용 (에디터에서 실행)
function testAPI() {
  var result = handleLogin('aisaju', 'flowmus2026');
  Logger.log(result.getContent());
}
