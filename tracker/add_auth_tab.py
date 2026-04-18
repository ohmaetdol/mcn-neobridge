#!/usr/bin/env python3
"""
채널암호 탭 추가 (UTM 트래커 시트)
실행: /tmp/gdrive_env/bin/python3 tracker/add_auth_tab.py
"""
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = '/Users/millstone/Desktop/Flowmus/scripts/token.json'
SHEET_ID = '1qRANdVfdy5Q05zI0yoSc15cPldX9m2OGXMg5tfyjZes'

NAVY = {'red': 0.06, 'green': 0.11, 'blue': 0.23}
WHITE = {'red': 1, 'green': 1, 'blue': 1}
RED = {'red': 0.86, 'green': 0.15, 'blue': 0.15}


def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    sheets = build('sheets', 'v4', credentials=creds)

    # 1. 채널암호 탭 추가
    try:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'requests': [{
                'addSheet': {
                    'properties': {
                        'title': '채널암호',
                        'sheetId': 105,
                        'tabColorStyle': {'rgbColor': RED},
                        'gridProperties': {'frozenRowCount': 1, 'columnCount': 3}
                    }
                }
            }]}
        ).execute()
        print("  채널암호 탭 생성 완료")
    except Exception as e:
        if 'already exists' in str(e):
            print("  채널암호 탭 이미 존재")
        else:
            raise

    # 2. 헤더 + 초기 비밀번호 입력
    header = ['채널slug', '비밀번호', '마지막변경일']
    sample = [
        ['sajjiknam', 'flowmus2026', '2026-04-17'],
        ['gogo', 'flowmus2026', '2026-04-17'],
        ['moneyroad', 'flowmus2026', '2026-04-17'],
    ]

    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range='채널암호!A1:C4',
        valueInputOption='RAW',
        body={'values': [header] + sample}
    ).execute()

    # 3. 헤더 스타일 + 열 너비
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': 105, 'startRowIndex': 0, 'endRowIndex': 1},
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': NAVY,
                            'textFormat': {'foregroundColor': WHITE, 'bold': True}
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                }
            },
            {
                'updateDimensionProperties': {
                    'range': {'sheetId': 105, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 3},
                    'properties': {'pixelSize': 160},
                    'fields': 'pixelSize'
                }
            },
        ]}
    ).execute()

    print("  채널암호 데이터 입력 완료")
    print(f"\n  ⚠️  반드시 비밀번호를 변경하세요! (기본값: flowmus2026)")
    print(f"  시트: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


if __name__ == '__main__':
    main()
