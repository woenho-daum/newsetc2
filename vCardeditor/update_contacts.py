"""
구글 연락처 이름 일괄 변경 스크립트
------------------------------------
- mapping.csv 의 전화번호를 기준으로 구글 연락처를 찾아 이름을 새 이름으로 변경합니다.
- CSV 형식 (헤더 필수): name,phone
    name  = 새로 설정할 이름
    phone = 매칭 기준이 되는 기존 전화번호

사용 전 준비물 (같은 폴더에 위치):
    1. credentials.json  (Google Cloud OAuth 클라이언트 ID, 데스크톱 앱)
    2. mapping.csv        (새이름-전화번호 목록)

설치:
    pip install google-auth google-auth-oauthlib google-api-python-client

실행:
    python update_contacts.py

주의:
    - 실행 전 반드시 contacts.google.com 에서 연락처를 CSV로 백업하세요.
    - 처음 실행 시 브라우저가 열리며 구글 로그인/권한 승인이 필요합니다.
    - 승인 후 token.json 이 생성되어 다음부터는 재인증 없이 실행됩니다.
"""

import csv
import re
import time
import os

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/contacts']
CSV_FILE = 'mapping.csv'
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

# 실제로 변경을 적용할지 여부. 먼저 True(미리보기)로 실행해 매칭 결과를 확인한 뒤
# 문제 없으면 False로 바꿔서 다시 실행하는 것을 권장합니다.
DRY_RUN = True


def normalize_phone(num: str) -> str:
    """전화번호에서 숫자만 남기고, 국가코드(82)를 제거한 뒤 뒤 10자리로 비교 기준을 통일합니다."""
    digits = re.sub(r'\D', '', num or '')
    if digits.startswith('0082'):
        digits = '0' + digits[4:]
    elif digits.startswith('82'):
        digits = '0' + digits[2:]
    return digits[-10:] if len(digits) >= 10 else digits


def get_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w', encoding='utf-8') as f:
            f.write(creds.to_json())
    return build('people', 'v1', credentials=creds)


def load_mapping(csv_file: str) -> dict:
    mapping = {}
    with open(csv_file, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        if 'name' not in reader.fieldnames or 'phone' not in reader.fieldnames:
            raise ValueError("CSV 헤더는 'name,phone' 이어야 합니다. 현재 헤더: " + str(reader.fieldnames))
        for row in reader:
            phone = normalize_phone(row['phone'])
            if not phone:
                continue
            mapping[phone] = row['name'].strip()
    return mapping


def fetch_all_connections(service):
    connections = []
    page_token = None
    while True:
        resp = service.people().connections().list(
            resourceName='people/me',
            pageSize=1000,
            personFields='names,phoneNumbers',
            pageToken=page_token
        ).execute()
        connections.extend(resp.get('connections', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return connections


def main():
    mapping = load_mapping(CSV_FILE)
    print(f"[매핑] CSV에서 {len(mapping)}건 로드")

    service = get_service()
    connections = fetch_all_connections(service)
    print(f"[구글 연락처] 총 {len(connections)}건 로드\n")

    matched_phones = set()
    updated = 0
    skipped_no_match = 0

    for person in connections:
        resource_name = person['resourceName']
        etag = person['etag']
        current_names = person.get('names', [])
        current_name = current_names[0].get('displayName') if current_names else '(이름없음)'
        phone_numbers = person.get('phoneNumbers', [])

        new_name = None
        matched_phone = None
        for p in phone_numbers:
            norm = normalize_phone(p.get('value', ''))
            if norm in mapping:
                new_name = mapping[norm]
                matched_phone = norm
                break

        if not new_name:
            skipped_no_match += 1
            continue

        matched_phones.add(matched_phone)

        if current_name == new_name:
            print(f"[동일] {current_name} (변경 불필요)")
            continue

        if DRY_RUN:
            print(f"[미리보기] {current_name}  ->  {new_name}")
            updated += 1
            continue

        body = {
            'etag': etag,
            'names': [{'givenName': new_name}]
        }
        try:
            service.people().updateContact(
                resourceName=resource_name,
                updatePersonFields='names',
                body=body
            ).execute()
            print(f"[변경완료] {current_name}  ->  {new_name}")
            updated += 1
            time.sleep(0.3)  # API 속도 제한 보호
        except Exception as e:
            print(f"[실패] {current_name} -> {new_name} : {e}")

    unmatched_csv_entries = set(mapping.keys()) - matched_phones
    print(f"\n=== 결과 요약 ===")
    print(f"처리(또는 미리보기) 건수: {updated}")
    print(f"매칭 안 된 연락처(스킵): {skipped_no_match}")
    if unmatched_csv_entries:
        print(f"CSV에는 있지만 폰 연락처에서 못 찾은 번호: {len(unmatched_csv_entries)}건")
        for ph in unmatched_csv_entries:
            print(f"  - {ph}")

    if DRY_RUN:
        print("\n※ 현재 DRY_RUN=True 상태로 실제 변경은 이루어지지 않았습니다.")
        print("   위 미리보기 결과가 맞으면 스크립트 상단의 DRY_RUN을 False로 바꾸고 다시 실행하세요.")


if __name__ == '__main__':
    main()
