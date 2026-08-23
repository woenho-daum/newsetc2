"""
구글 연락처 이름 일괄 변경 스크립트
------------------------------------
- mapping.csv 를 기준으로 구글 연락처를 찾아 이름을 새 이름으로 변경합니다.
- "전화번호"와 "예전 이름(old_name)"이 모두 일치하는 경우에만 변경합니다.
  (전화번호만 같고 이름이 다르면 오매칭 방지를 위해 변경하지 않습니다.)
- CSV 형식 (헤더 필수): old_name,phone,new_name
    old_name = 매칭 기준이 되는 기존 이름
    phone    = 매칭 기준이 되는 기존 전화번호
    new_name = 새로 설정할 이름

사용 전 준비물 (같은 폴더에 위치):
    1. credentials.json  (Google Cloud OAuth 클라이언트 ID, 데스크톱 앱)
    2. mapping.csv        (예전이름-전화번호-새이름 목록)

설치:
    pip install google-auth google-auth-oauthlib google-api-python-client

실행:
    python update_contacts.py

결과:
    실행이 끝나면 결과 요약과 함께,
    변경/미변경 각 항목에 대해 "이유"가 표시된 전체 목록이 출력되고,
    result_report.csv 파일로도 저장됩니다.

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
REPORT_FILE = 'result_report.csv'

# 실제로 변경을 적용할지 여부. 먼저 True(미리보기)로 실행해 매칭 결과를 확인한 뒤
# 문제 없으면 False로 바꿔서 다시 실행하는 것을 권장합니다.
DRY_RUN = True


def normalize_phone(num: str) -> str:
    """전화번호에서 숫자만 남기고, 국가코드(+82/0082)만 국내 형식(0으로 시작)으로 변환합니다."""
    digits = re.sub(r'\D', '', num or '')
    if digits.startswith('0082'):
        digits = '0' + digits[4:]
    elif digits.startswith('82') and not digits.startswith('0'):
        digits = '0' + digits[2:]
    return digits


def normalize_name(name: str) -> str:
    """이름 비교 시 앞뒤 공백/중복 공백 차이로 인한 오매칭을 줄이기 위한 정규화."""
    return re.sub(r'\s+', ' ', (name or '').strip())


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
    """반환값: { 정규화된 전화번호: {'old_name': str, 'new_name': str} }"""
    mapping = {}
    with open(csv_file, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        required = {'old_name', 'phone', 'new_name'}
        if not required.issubset(set(reader.fieldnames or [])):
            raise ValueError(
                "CSV 헤더는 'old_name,phone,new_name' 이어야 합니다. 현재 헤더: "
                + str(reader.fieldnames)
            )
        for row in reader:
            phone = normalize_phone(row['phone'])
            if not phone:
                continue
            mapping[phone] = {
                'old_name': normalize_name(row['old_name']),
                'new_name': row['new_name'].strip(),
            }
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
    report_rows = []  # 각 행: dict(old_name, phone, new_name, current_name, status, reason)

    def add_report(old_name, phone, new_name, current_name, status, reason):
        report_rows.append({
            'old_name': old_name,
            'phone': phone,
            'new_name': new_name,
            'current_name': current_name,
            'status': status,
            'reason': reason,
        })

    for person in connections:
        resource_name = person['resourceName']
        etag = person['etag']
        current_names = person.get('names', [])
        current_name = current_names[0].get('displayName') if current_names else '(이름없음)'
        current_name_norm = normalize_name(current_name)
        phone_numbers = person.get('phoneNumbers', [])

        entry = None
        matched_phone = None
        for p in phone_numbers:
            norm = normalize_phone(p.get('value', ''))
            if norm in mapping:
                entry = mapping[norm]
                matched_phone = norm
                break

        # 1) 전화번호 자체가 CSV에 없는 연락처는 애초에 대상이 아니므로 리포트에 넣지 않음
        if not entry:
            continue

        matched_phones.add(matched_phone)
        old_name_expected = entry['old_name']
        new_name = entry['new_name']

        # 2) 전화번호는 일치하지만 예전 이름이 다른 경우 -> 변경하지 않음
        if current_name_norm != old_name_expected:
            add_report(
                old_name_expected, matched_phone, new_name, current_name,
                status='미변경',
                reason=f"이름 불일치 (CSV old_name='{old_name_expected}' / 실제='{current_name}')"
            )
            continue

        # 3) 이미 새 이름과 동일한 경우 -> 변경 불필요
        if current_name == new_name:
            add_report(
                old_name_expected, matched_phone, new_name, current_name,
                status='미변경',
                reason='이미 새 이름과 동일함 (변경 불필요)'
            )
            continue

        # 4) 전화번호+예전이름 모두 일치 -> 변경 대상
        if DRY_RUN:
            add_report(
                old_name_expected, matched_phone, new_name, current_name,
                status='변경예정(DRY_RUN)',
                reason='전화번호/이름 모두 일치'
            )
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
            add_report(
                old_name_expected, matched_phone, new_name, current_name,
                status='변경완료',
                reason='전화번호/이름 모두 일치'
            )
            time.sleep(0.3)  # API 속도 제한 보호
        except Exception as e:
            add_report(
                old_name_expected, matched_phone, new_name, current_name,
                status='실패',
                reason=f'API 오류: {e}'
            )

    # CSV에는 있지만 구글 연락처에서 전화번호 자체를 못 찾은 경우
    unmatched_csv_entries = set(mapping.keys()) - matched_phones
    for phone in unmatched_csv_entries:
        entry = mapping[phone]
        add_report(
            entry['old_name'], phone, entry['new_name'], '(해당없음)',
            status='미변경',
            reason='구글 연락처에서 해당 전화번호를 찾을 수 없음'
        )

    # ---- 결과 출력 ----
    print("\n=== 상세 결과 목록 ===")
    header = f"{'상태':12} {'전화번호':14} {'CSV old_name':15} {'실제 이름':15} {'new_name':15} 이유"
    print(header)
    print('-' * len(header))
    for row in report_rows:
        print(f"{row['status']:12} {row['phone']:14} {row['old_name']:15} "
              f"{row['current_name']:15} {row['new_name']:15} {row['reason']}")

    # ---- 리포트 CSV 저장 ----
    with open(REPORT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'status', 'phone', 'old_name', 'current_name', 'new_name', 'reason'
        ])
        writer.writeheader()
        for row in report_rows:
            writer.writerow({
                'status': row['status'],
                'phone': row['phone'],
                'old_name': row['old_name'],
                'current_name': row['current_name'],
                'new_name': row['new_name'],
                'reason': row['reason'],
            })

    changed_count = sum(1 for r in report_rows if r['status'] in ('변경완료', '변경예정(DRY_RUN)'))
    unchanged_count = sum(1 for r in report_rows if r['status'] == '미변경')
    failed_count = sum(1 for r in report_rows if r['status'] == '실패')

    print(f"\n=== 결과 요약 ===")
    print(f"전체 CSV 매핑 건수: {len(mapping)}")
    print(f"변경(완료/예정): {changed_count}")
    print(f"미변경: {unchanged_count}")
    print(f"실패: {failed_count}")
    print(f"상세 리포트 저장됨: {REPORT_FILE}")

    if DRY_RUN:
        print("\n※ 현재 DRY_RUN=True 상태로 실제 변경은 이루어지지 않았습니다.")
        print("   위 미리보기 결과가 맞으면 스크립트 상단의 DRY_RUN을 False로 바꾸고 다시 실행하세요.")


if __name__ == '__main__':
    main()
