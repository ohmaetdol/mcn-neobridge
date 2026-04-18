#!/usr/bin/env python3
"""
캠페인 + 쿠폰 관리 CLI
사용법:
  python3 tracker/manage_campaigns.py add        # 새 캠페인 추가 (대화형)
  python3 tracker/manage_campaigns.py list       # 캠페인 목록
  python3 tracker/manage_campaigns.py add --channel sajjiknam --type 강의 --platform 타이탄클래스 --url https://... --discount 5% --rs 20

쿠폰 코드 자동 생성 규칙:
  {채널접두어}-{플랫폼약어}-{순번}
  예: SAJJIK-TITAN-01, GOGO-BBQ-01, MONEY-FAST-01
"""
import argparse
import re
import sys
from datetime import date

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = '/Users/millstone/Desktop/Flowmus/scripts/token.json'
SHEET_ID = '1qRANdVfdy5Q05zI0yoSc15cPldX9m2OGXMg5tfyjZes'
TAB_NAME = '캠페인관리'
GO_BASE = 'https://ohmaetdol.github.io/mcn-neobridge/go/?c='
TRACKER_BASE = 'https://ohmaetdol.github.io/mcn-neobridge/tracker/'

# ── 채널 접두어 매핑 ──
CHANNEL_PREFIX = {
    'sajjiknam': 'SAJJIK',
    'hyunjang': 'HYUNJANG',
    'moneyroad': 'MONEY',
}

# ── 플랫폼 약어 매핑 (자동 확장) ──
PLATFORM_ABBR = {
    '타이탄클래스': 'TITAN',
    '무아클래스': 'MUA',
    '패스트캠퍼스': 'FAST',
    '클래스101': 'CLASS101',
    '콜로소': 'COLOSO',
    '인프런': 'INFLEARN',
    '직방프랜차이즈': 'REAL',
    'BBQ': 'BBQ',
    '교촌': 'GYOCHON',
    '맘스터치': 'MOMS',
    'bhc': 'BHC',
    '메가커피': 'MEGA',
    '컴포즈커피': 'COMPOSE',
}


def get_sheets():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    return build('sheets', 'v4', credentials=creds)


def read_campaigns(sheets):
    """시트에서 전체 캠페인 읽기"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f'{TAB_NAME}!A:K'
    ).execute()
    rows = result.get('values', [])
    if len(rows) < 2:
        return [], rows[0] if rows else []
    return rows[1:], rows[0]


def make_slug(channel, video_id):
    """캠페인 slug 생성: channel-videoID앞6자"""
    vid = video_id[:6] if video_id else 'novid'
    return f"{channel}-{vid}"


def make_coupon(video_id, existing_coupons):
    """쿠폰 코드 자동 생성: FLOW-videoID앞6자-순번"""
    vid = video_id[:6] if video_id else 'NOvid'
    prefix = f"FLOW-{vid}-"

    # 기존 쿠폰 중 같은 접두어 찾아서 순번 결정
    max_seq = 0
    for c in existing_coupons:
        if c.startswith(prefix):
            try:
                seq = int(c[len(prefix):])
                max_seq = max(max_seq, seq)
            except ValueError:
                pass

    return f"{prefix}{max_seq + 1:02d}"


def cmd_list(args):
    """캠페인 목록 출력"""
    sheets = get_sheets()
    rows, headers = read_campaigns(sheets)

    if not rows:
        print("  등록된 캠페인이 없습니다.")
        return

    print(f"\n  {'캠페인 slug':<20} {'채널':<12} {'유형':<6} {'플랫폼':<16} {'쿠폰 코드':<22} {'RS':<6} {'상태':<6}")
    print("  " + "─" * 90)

    for r in rows:
        # 열: slug(0) type(1) url(2) platform(3) videoId(4) coupon(5) discount(6) rs(7) status(8) created(9) channel(10)
        slug = r[0] if len(r) > 0 else ''
        typ = r[1] if len(r) > 1 else ''
        plat = r[3] if len(r) > 3 else ''
        coupon = r[5] if len(r) > 5 else ''
        rs = r[7] if len(r) > 7 else ''
        status = r[8] if len(r) > 8 else ''
        channel = r[10] if len(r) > 10 else ''

        status_mark = '●' if status == '활성' else '○'
        print(f"  {slug:<20} {channel:<12} {typ:<6} {plat:<16} {coupon:<22} {rs:<6} {status_mark} {status}")

    print(f"\n  총 {len(rows)}개 캠페인\n")


def cmd_add(args):
    """새 캠페인 추가"""
    sheets = get_sheets()
    rows, _ = read_campaigns(sheets)

    existing_coupons = []
    existing_slugs = set()
    for r in rows:
        if len(r) > 5 and r[5]:
            existing_coupons.append(r[5])
        if len(r) > 0 and r[0]:
            existing_slugs.add(r[0])

    # 대화형 or 인자 입력
    channel = args.channel or input_with_options('채널 선택', list(CHANNEL_PREFIX.keys()))
    camp_type = args.type or input_with_options('캠페인 유형', ['강의', '가맹'])
    platform = args.platform or input('  플랫폼/브랜드명: ').strip()
    url = args.url or input('  타겟 URL: ').strip()
    video_id = args.video_id or extract_video_id(url) or input('  영상 ID (없으면 Enter): ').strip()
    discount = args.discount or input('  할인 (예: 5%, 가맹비 50만 할인): ').strip()
    rs = args.rs or ('20%' if camp_type == '강의' else '10%')

    # 자동 생성 (영상 ID 기반)
    slug = make_slug(channel, video_id)
    # slug 중복 체크
    base_slug = slug
    counter = 2
    while slug in existing_slugs:
        slug = f"{base_slug}-{counter}"
        counter += 1

    coupon = make_coupon(video_id, existing_coupons)
    today = date.today().isoformat()

    # 확인
    print(f"\n  ┌─────────────────────────────────────────────┐")
    print(f"  │  새 캠페인 등록 확인                         │")
    print(f"  ├─────────────────────────────────────────────┤")
    print(f"  │  캠페인 slug : {slug:<28} │")
    print(f"  │  채널        : {channel:<28} │")
    print(f"  │  유형        : {camp_type:<28} │")
    print(f"  │  플랫폼      : {platform:<28} │")
    print(f"  │  쿠폰 코드   : {coupon:<28} │")
    print(f"  │  할인        : {discount:<28} │")
    print(f"  │  RS%         : {rs:<28} │")
    print(f"  └─────────────────────────────────────────────┘")

    confirm = input("\n  등록하시겠습니까? (Y/n): ").strip().lower()
    if confirm == 'n':
        print("  취소되었습니다.")
        return

    # Google Sheet에 추가
    new_row = [slug, camp_type, url, platform, video_id, coupon, discount, rs, '활성', today, channel]

    sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=f'{TAB_NAME}!A:K',
        valueInputOption='RAW',
        body={'values': [new_row]}
    ).execute()

    go_link = GO_BASE + slug
    tracker_link = TRACKER_BASE + '?id=' + channel

    print(f"\n  ✅ 캠페인 등록 완료!")
    print(f"\n  📋 쿠폰 코드 : {coupon}")
    print(f"  🔗 리다이렉트 : {go_link}")
    print(f"  📊 대시보드   : {tracker_link}")
    print(f"\n  ⚠️  쿠폰 코드 '{coupon}'을 플랫폼({platform})에 등록해주세요.")
    print(f"      할인 조건: {discount}\n")


def input_with_options(prompt, options):
    """옵션 목록에서 선택"""
    print(f"\n  {prompt}:")
    for i, opt in enumerate(options, 1):
        print(f"    {i}. {opt}")
    while True:
        choice = input(f"  선택 (1-{len(options)}) 또는 직접 입력: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return options[int(choice) - 1]
        if choice:
            return choice


def extract_video_id(url):
    """YouTube URL에서 비디오 ID 추출"""
    match = re.search(r'(?:v=|youtu\.be/)([a-zA-Z0-9_-]{11})', url or '')
    return match.group(1) if match else ''


def cmd_generate(args):
    """캠페인에 대한 쿠폰 풀 벌크 생성"""
    sheets = get_sheets()
    rows, _ = read_campaigns(sheets)

    slug = args.slug
    count = args.count

    # 캠페인 존재 확인
    campaign = None
    for r in rows:
        if r[0] == slug:
            campaign = r
            break

    if not campaign:
        print(f"  캠페인 '{slug}'을 찾을 수 없습니다.")
        print("  사용 가능한 캠페인:")
        for r in rows:
            print(f"    - {r[0]}")
        return

    video_id = campaign[4] if len(campaign) > 4 else ''
    platform = campaign[3] if len(campaign) > 3 else ''
    vid = video_id[:6] if video_id else 'NOvid'

    # 기존 쿠폰풀 확인
    pool_result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range='쿠폰풀!A:B'
    ).execute()
    pool_rows = pool_result.get('values', [])[1:]  # 헤더 제외

    existing_for_slug = [r[0] for r in pool_rows if len(r) > 1 and r[1] == slug]
    start_num = len(existing_for_slug) + 1

    # 쿠폰 생성 (FLOW-영상ID-순번)
    new_coupons = []
    for i in range(start_num, start_num + count):
        code = f"FLOW-{vid}-{i:04d}"
        new_coupons.append([code, slug, '미발급', '', '', '', ''])

    sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID, range='쿠폰풀!A:G',
        valueInputOption='RAW', body={'values': new_coupons}
    ).execute()

    print(f"\n  쿠폰 {count}개 생성 완료!")
    print(f"  캠페인: {slug}")
    print(f"  코드 범위: {new_coupons[0][0]} ~ {new_coupons[-1][0]}")
    print(f"  기존: {len(existing_for_slug)}개 → 총: {len(existing_for_slug) + count}개")
    print(f"\n  이 코드를 플랫폼({platform})에 벌크 등록해주세요.\n")


def cmd_stock(args):
    """쿠폰 풀 현황"""
    sheets = get_sheets()

    pool_result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range='쿠폰풀!A:C'
    ).execute()
    pool_rows = pool_result.get('values', [])[1:]

    stats = {}
    for r in pool_rows:
        if len(r) < 3:
            continue
        slug = r[1]
        if slug not in stats:
            stats[slug] = {'total': 0, 'remaining': 0, 'used': 0}
        stats[slug]['total'] += 1
        if r[2] == '미발급':
            stats[slug]['remaining'] += 1
        else:
            stats[slug]['used'] += 1

    if not stats:
        print("  쿠폰풀이 비어있습니다. generate 명령어로 쿠폰을 생성하세요.")
        return

    print(f"\n  {'캠페인':<20} {'총 쿠폰':<10} {'발급됨':<10} {'잔여':<10} {'소진율'}")
    print("  " + "─" * 65)
    for slug, s in stats.items():
        rate = f"{s['used'] / s['total'] * 100:.0f}%" if s['total'] > 0 else '-'
        bar = '█' * int(s['used'] / max(s['total'], 1) * 10) + '░' * (10 - int(s['used'] / max(s['total'], 1) * 10))
        print(f"  {slug:<20} {s['total']:<10} {s['used']:<10} {s['remaining']:<10} {bar} {rate}")
    print()


def main():
    parser = argparse.ArgumentParser(description='캠페인 + 쿠폰 관리')
    sub = parser.add_subparsers(dest='command')

    # add 명령어
    add_p = sub.add_parser('add', help='새 캠페인 추가')
    add_p.add_argument('--channel', help='채널 slug (sajjiknam, hyunjang, moneyroad)')
    add_p.add_argument('--type', help='캠페인 유형 (강의, 가맹)')
    add_p.add_argument('--platform', help='플랫폼/브랜드명')
    add_p.add_argument('--url', help='타겟 URL')
    add_p.add_argument('--video-id', help='영상 ID', dest='video_id')
    add_p.add_argument('--discount', help='할인 조건')
    add_p.add_argument('--rs', help='RS 비율 (기본: 강의 20%%, 가맹 10%%)')

    # list 명령어
    sub.add_parser('list', help='캠페인 목록')

    # generate 명령어
    gen_p = sub.add_parser('generate', help='쿠폰 벌크 생성')
    gen_p.add_argument('--slug', required=True, help='캠페인 slug')
    gen_p.add_argument('--count', type=int, default=100, help='생성할 쿠폰 수 (기본 100)')

    # stock 명령어
    sub.add_parser('stock', help='쿠폰 풀 현황')

    args = parser.parse_args()

    if args.command == 'add':
        cmd_add(args)
    elif args.command == 'list':
        cmd_list(args)
    elif args.command == 'generate':
        cmd_generate(args)
    elif args.command == 'stock':
        cmd_stock(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
