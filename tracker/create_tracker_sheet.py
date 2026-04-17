#!/usr/bin/env python3
"""
UTM 트래커 + 쿠폰 관리 + RS 정산 시트 생성
실행: /tmp/gdrive_env/bin/python3 tracker/create_tracker_sheet.py
"""
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = '/Users/millstone/Desktop/Flowmus/scripts/token.json'
DRIVE_ID = '0ALF8EAkiiYy2Uk9PVA'

# 색상
NAVY  = {'red': 0.06, 'green': 0.11, 'blue': 0.23}
WHITE = {'red': 1, 'green': 1, 'blue': 1}
BLUE  = {'red': 0.15, 'green': 0.39, 'blue': 0.92}   # #2563EB
LBLUE = {'red': 0.94, 'green': 0.96, 'blue': 1}       # #EFF4FF
GRAY  = {'red': 0.95, 'green': 0.95, 'blue': 0.95}
GREEN = {'red': 0.22, 'green': 0.66, 'blue': 0.36}
ORANGE = {'red': 0.92, 'green': 0.35, 'blue': 0.05}   # #ea580c

TABS = [
    {'title': '캠페인관리', 'id': 0, 'color': BLUE},
    {'title': '클릭로그',   'id': 1, 'color': GREEN},
    {'title': '월별정산',   'id': 2, 'color': ORANGE},
]

CAMPAIGN_HEADERS = ['캠페인 slug', '유형', '타겟 URL', '플랫폼/브랜드', '영상 ID', '쿠폰 코드', '할인', 'RS%', '상태', '생성일']
LOG_HEADERS = ['타임스탬프', '캠페인 slug', 'referrer', 'user-agent']
SETTLEMENT_HEADERS = ['월', '캠페인 slug', '유형', '플랫폼', '클릭수', '쿠폰사용', '전환율', '매출(원)', 'RS%', 'RS금액(원)']

SAMPLE_CAMPAIGNS = [
    ['titan-blog', '강의', 'https://www.youtube.com/watch?v=M0nYqv_SFJI', '타이탄클래스', 'M0nYqv_SFJI', 'SAJJIK-BLOG-01', '5%', '20%', '활성', '2026-04-17'],
    ['mua-invest', '강의', 'https://www.youtube.com/watch?v=iynwLv5LJuY', '무아클래스', 'iynwLv5LJuY', 'SAJJIK-INVEST-01', '5%', '20%', '활성', '2026-04-17'],
]


def get_services():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    drive = build('drive', 'v3', credentials=creds)
    sheets = build('sheets', 'v4', credentials=creds)
    return drive, sheets


def get_folder_id(drive, name_contains):
    r = drive.files().list(
        q=f"'{DRIVE_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        corpora='drive', driveId=DRIVE_ID,
        includeItemsFromAllDrives=True, supportsAllDrives=True,
        fields='files(id, name)'
    ).execute()
    for f in r.get('files', []):
        if name_contains in f['name']:
            return f['id']
    return DRIVE_ID


def create_sheet(drive, name, parent_id):
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [parent_id],
        'driveId': DRIVE_ID
    }
    f = drive.files().create(body=meta, supportsAllDrives=True, fields='id,name,webViewLink').execute()
    print(f"  시트 생성: {f['name']}")
    print(f"  링크: {f['webViewLink']}")
    return f['id'], f['webViewLink']


def cell(value, bg=None, fg=None, bold=False):
    c = {'userEnteredValue': {'stringValue': str(value)}}
    fmt = {}
    if bg:
        fmt['backgroundColor'] = bg
    tf = {}
    if fg:
        tf['foregroundColor'] = fg
    if bold:
        tf['bold'] = True
    if tf:
        fmt['textFormat'] = tf
    if fmt:
        c['userEnteredFormat'] = fmt
    return c


def header_row(headers, bg=NAVY, fg=WHITE):
    return {'values': [cell(h, bg=bg, fg=fg, bold=True) for h in headers]}


def data_row(values, bg=None):
    return {'values': [cell(v, bg=bg) for v in values]}


def setup_tabs(sheets, sheet_id):
    # 탭은 이미 생성됨 (id: 100, 101, 102). 헤더 + 데이터만 입력
    # 캠페인관리 탭
    campaign_rows = [header_row(CAMPAIGN_HEADERS)]
    for row in SAMPLE_CAMPAIGNS:
        campaign_rows.append(data_row(row))

    requests2 = [
        {
            'updateCells': {
                'rows': campaign_rows,
                'fields': 'userEnteredValue,userEnteredFormat',
                'start': {'sheetId': 100, 'rowIndex': 0, 'columnIndex': 0}
            }
        },
        # 클릭로그 탭 헤더
        {
            'updateCells': {
                'rows': [header_row(LOG_HEADERS)],
                'fields': 'userEnteredValue,userEnteredFormat',
                'start': {'sheetId': 101, 'rowIndex': 0, 'columnIndex': 0}
            }
        },
        # 월별정산 탭 헤더
        {
            'updateCells': {
                'rows': [header_row(SETTLEMENT_HEADERS)],
                'fields': 'userEnteredValue,userEnteredFormat',
                'start': {'sheetId': 102, 'rowIndex': 0, 'columnIndex': 0}
            }
        },
        # 헤더 행 고정 (캠페인관리)
        {
            'updateSheetProperties': {
                'properties': {'sheetId': 100, 'gridProperties': {'frozenRowCount': 1}},
                'fields': 'gridProperties.frozenRowCount'
            }
        },
        {
            'updateSheetProperties': {
                'properties': {'sheetId': 101, 'gridProperties': {'frozenRowCount': 1}},
                'fields': 'gridProperties.frozenRowCount'
            }
        },
        {
            'updateSheetProperties': {
                'properties': {'sheetId': 102, 'gridProperties': {'frozenRowCount': 1}},
                'fields': 'gridProperties.frozenRowCount'
            }
        },
        # 열 너비 조정 (캠페인관리)
        {
            'updateDimensionProperties': {
                'range': {'sheetId': 100, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
                'properties': {'pixelSize': 140}, 'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {'sheetId': 100, 'dimension': 'COLUMNS', 'startIndex': 2, 'endIndex': 3},
                'properties': {'pixelSize': 300}, 'fields': 'pixelSize'
            }
        },
    ]

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id, body={'requests': requests2}
    ).execute()


def main():
    print("UTM 트래커 시트 설정 시작...")
    drive, sheets = get_services()

    # 이미 생성된 시트 사용
    sheet_id = '1qRANdVfdy5Q05zI0yoSc15cPldX9m2OGXMg5tfyjZes'
    print(f"  기존 시트 사용: {sheet_id}")

    setup_tabs(sheets, sheet_id)
    print(f"\n  시트 ID: {sheet_id}")
    print(f"  완료!")


if __name__ == '__main__':
    main()
