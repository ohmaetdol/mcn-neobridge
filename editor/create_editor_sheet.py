#!/usr/bin/env python3
"""
편집자 보드 전용 Google Sheet 생성
실행: /tmp/gdrive_env/bin/python3 editor/create_editor_sheet.py
"""
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
TOKEN_FILE = '/Users/millstone/Desktop/Flowmus/scripts/token.json'
DRIVE_ID = '0ALF8EAkiiYy2Uk9PVA'  # Flowmus 공유 드라이브

NAVY = {'red': 0.06, 'green': 0.11, 'blue': 0.23}
WHITE = {'red': 1, 'green': 1, 'blue': 1}
BLUE = {'red': 0.15, 'green': 0.39, 'blue': 0.92}
GREEN = {'red': 0.13, 'green': 0.48, 'blue': 0.22}
ORANGE = {'red': 0.92, 'green': 0.35, 'blue': 0.05}


def main():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    sheets = build('sheets', 'v4', credentials=creds)
    drive = build('drive', 'v3', credentials=creds)

    # 1. 새 스프레드시트 생성
    body = {
        'properties': {'title': '편집자 보드'},
        'sheets': [
            {
                'properties': {
                    'sheetId': 0, 'title': '영상목록',
                    'gridProperties': {'frozenRowCount': 1, 'columnCount': 7},
                    'tabColorStyle': {'rgbColor': ORANGE}
                }
            },
            {
                'properties': {
                    'sheetId': 1, 'title': '편집신청',
                    'gridProperties': {'frozenRowCount': 1, 'columnCount': 8},
                    'tabColorStyle': {'rgbColor': BLUE}
                }
            },
            {
                'properties': {
                    'sheetId': 2, 'title': '편집자',
                    'gridProperties': {'frozenRowCount': 1, 'columnCount': 4},
                    'tabColorStyle': {'rgbColor': GREEN}
                }
            },
            {
                'properties': {
                    'sheetId': 3, 'title': '편집코치가이드',
                    'gridProperties': {'frozenRowCount': 1, 'columnCount': 2},
                    'tabColorStyle': {'rgbColor': NAVY}
                }
            },
        ]
    }

    result = sheets.spreadsheets().create(body=body).execute()
    sheet_id = result['spreadsheetId']
    print(f"  시트 생성 완료: {sheet_id}")

    # 공유 드라이브로 이동
    try:
        file = drive.files().get(fileId=sheet_id, fields='parents').execute()
        prev_parents = ",".join(file.get('parents', []))

        # 00_Management_Admin 폴더 찾기
        query = f"name='00_Management_Admin' and mimeType='application/vnd.google-apps.folder' and '{DRIVE_ID}' in parents"
        folders = drive.files().list(
            q=query, driveId=DRIVE_ID,
            includeItemsFromAllDrives=True, supportsAllDrives=True,
            corpora='drive', fields='files(id,name)'
        ).execute().get('files', [])

        if folders:
            drive.files().update(
                fileId=sheet_id,
                addParents=folders[0]['id'],
                removeParents=prev_parents,
                supportsAllDrives=True,
                fields='id,parents'
            ).execute()
            print(f"  공유 드라이브 이동 완료")
    except Exception as e:
        print(f"  드라이브 이동 실패 (수동으로 이동하세요): {e}")

    # 2. 헤더 입력
    headers = {
        '영상목록!A1:G1': [['영상ID', '채널', '출연자', '제목', '드롭박스링크', '마감일', '상태']],
        '편집신청!A1:H1': [['신청일', '영상ID', '편집자', '편집기획서', '예상완료일', '승인상태', '승인일', '메모']],
        '편집자!A1:D1': [['이름', '이메일', '전문분야', '상태']],
        '편집코치가이드!A1:B1': [['채널', '가이드내용']],
    }

    for range_name, values in headers.items():
        sheets.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=range_name,
            valueInputOption='RAW', body={'values': values}
        ).execute()

    # 3. 편집코치 기본 가이드
    guide_data = [
        ['사장찍어주는남자', '1. Never-Cut 씬은 절대 삭제 금지\n2. 오프닝 40초가 전부 — 가장 강한 씬으로 시작\n3. 모든 촬영 장소 최소 1씬 포함\n4. BGM은 감정 변화 포인트에만\n5. 자막은 핵심 발언만 강조'],
        ['부업소개하는남자', '1. Never-Cut 씬은 절대 삭제 금지\n2. 오프닝 40초가 전부 — 가장 강한 씬으로 시작\n3. 모든 촬영 장소 최소 1씬 포함\n4. BGM은 감정 변화 포인트에만\n5. 자막은 핵심 발언만 강조'],
    ]

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range='편집코치가이드!A2:B3',
        valueInputOption='RAW', body={'values': guide_data}
    ).execute()

    # 4. 헤더 스타일 + 드롭다운
    requests = []

    # 헤더 스타일 (모든 탭)
    for sid in [0, 1, 2, 3]:
        requests.append({
            'repeatCell': {
                'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': NAVY,
                        'textFormat': {'foregroundColor': WHITE, 'bold': True, 'fontSize': 10}
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }
        })

    # 영상목록 G열(상태) 드롭다운
    requests.append({
        'setDataValidation': {
            'range': {'sheetId': 0, 'startRowIndex': 1, 'endRowIndex': 100, 'startColumnIndex': 6, 'endColumnIndex': 7},
            'rule': {
                'condition': {
                    'type': 'ONE_OF_LIST',
                    'values': [
                        {'userEnteredValue': '대기'},
                        {'userEnteredValue': '신청됨'},
                        {'userEnteredValue': '편집중'},
                        {'userEnteredValue': '초안제출'},
                        {'userEnteredValue': '수정중'},
                        {'userEnteredValue': '완료'},
                    ]
                },
                'showCustomUi': True, 'strict': True
            }
        }
    })

    # 편집신청 F열(승인상태) 드롭다운
    requests.append({
        'setDataValidation': {
            'range': {'sheetId': 1, 'startRowIndex': 1, 'endRowIndex': 100, 'startColumnIndex': 5, 'endColumnIndex': 6},
            'rule': {
                'condition': {
                    'type': 'ONE_OF_LIST',
                    'values': [
                        {'userEnteredValue': '대기'},
                        {'userEnteredValue': '승인'},
                        {'userEnteredValue': '반려'},
                    ]
                },
                'showCustomUi': True, 'strict': True
            }
        }
    })

    # 편집자 D열(상태) 드롭다운
    requests.append({
        'setDataValidation': {
            'range': {'sheetId': 2, 'startRowIndex': 1, 'endRowIndex': 20, 'startColumnIndex': 3, 'endColumnIndex': 4},
            'rule': {
                'condition': {
                    'type': 'ONE_OF_LIST',
                    'values': [
                        {'userEnteredValue': '활성'},
                        {'userEnteredValue': '비활성'},
                    ]
                },
                'showCustomUi': True, 'strict': True
            }
        }
    })

    # 열 너비 설정
    col_widths = [
        (0, 0, 1, 100), (0, 1, 2, 140), (0, 2, 3, 100), (0, 3, 4, 200),
        (0, 4, 5, 300), (0, 5, 6, 100), (0, 6, 7, 80),
        (1, 0, 1, 100), (1, 1, 2, 100), (1, 2, 3, 100), (1, 3, 4, 400),
        (1, 4, 5, 100), (1, 5, 6, 80), (1, 6, 7, 100), (1, 7, 8, 200),
    ]
    for sid, start, end, width in col_widths:
        requests.append({
            'updateDimensionProperties': {
                'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': start, 'endIndex': end},
                'properties': {'pixelSize': width},
                'fields': 'pixelSize'
            }
        })

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests}
    ).execute()

    print(f"\n  편집자 보드 시트 생성 완료!")
    print(f"  시트 ID: {sheet_id}")
    print(f"  링크: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    print(f"\n  다음 단계:")
    print(f"  1. 편집자 탭에 편집자 이름/이메일 입력")
    print(f"  2. 확장 프로그램 > Apps Script > 코드 붙여넣기 > 배포")


if __name__ == '__main__':
    main()
