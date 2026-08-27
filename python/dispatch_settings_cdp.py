#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
크롬(디버그 포트 9223)에 CDP로 연결하여
https://bumil.mobilhi.com/a3/view_dispatch_settings/ 페이지를 열고,
영업소(본사 -> 출퇴근 -> 심야) 순서로 선택 후 '조회' 버튼을 눌러
데이터를 조회하고, 조회가 끝날 때마다 같은 폴더의
collect_dispatch_settings.py 의 main() 함수를 호출한다.

collect_dispatch_settings.py 를 별도 프로세스(subprocess)로 띄우지 않고
import 해서 같은 파이썬 프로세스 안에서 함수로 호출한다.
-> VSCode 에서 이 파일을 디버그로 실행하면 collect_dispatch_settings.py
   내부의 브레이크포인트도 같은 디버그 세션에서 그대로 잡힌다.
"""

import os
import sys

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# 이 스크립트 파일이 있는 폴더로 작업 디렉토리를 강제 고정 (디버그 실행시 작업디렉토리를 환경파일 폴더로 한다. 이를 소스폴더로 변경하려면...)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)   # -> 현재소스의 경로를 cwd 고정

# collect_dispatch_settings.py 를 같은 폴더에서 import 할 수 있도록 경로 추가
# 같은 폴더의 모듈을 import할 수 있도록 경로 추가, 설정이 없다면 시스템 폴더, 기본 임포트는 위에다가 해야한다. 아님 꼬이겠지
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import collect_dispatch_settings  # noqa: E402  (경로 추가 이후에 import 해야 함)

# ----------------------------------------------------------------------
# 설정값
# ----------------------------------------------------------------------
CDP_URL = "http://localhost:9223"
TARGET_URL = "https://bumil.mobilhi.com/a3/view_dispatch_settings/"

# 조회할 영업소 순서 (select 태그의 <option> 텍스트와 동일해야 함)
OFFICE_OPTIONS = ["본사", "출퇴근", "심야"]

# 조회 결과가 서버에서 내려오는 htmx 요청 URL (form action)
SEARCH_RESPONSE_URL_PART = "view_dispatch_settings_tab1"

SELECT_ID = "#id_search_office_false"
SEARCH_BUTTON_SELECTOR = "form.search_box button.btn-outline-primary"


def run_collect_script():
    """collect_dispatch_settings.main() 을 같은 프로세스 안에서 직접 호출한다."""
    print("[정보] collect_dispatch_settings.main() 호출...")
    try:
        collect_dispatch_settings.main()
    except Exception as e:
        print(f"[경고] collect_dispatch_settings.main() 실행 중 예외 발생: {e}")
        raise
    print("[정보] collect_dispatch_settings.main() 실행 완료")


def select_office_and_search(page, office_name: str):
    """영업소 select 값을 office_name 으로 바꾸고 조회 버튼을 눌러 결과를 기다린다."""
    print(f"[정보] 영업소 '{office_name}' 선택 중...")

    # select 값 변경 (option 의 표시 텍스트로 선택 -> change 이벤트도 자동 발생)
    page.select_option(SELECT_ID, label=office_name)

    # select 의 change 핸들러가 노선(select_office_false) 옵션을 갱신하므로 잠시 대기
    page.wait_for_timeout(300)

    print("[정보] '조회' 버튼 클릭...")
    try:
        # htmx 가 폼 submit 을 가로채서 GET 요청을 보내므로, 그 응답을 기다린다.
        with page.expect_response(
            lambda resp: SEARCH_RESPONSE_URL_PART in resp.url and resp.request.method == "GET",
            timeout=15000,
        ):
            page.click(SEARCH_BUTTON_SELECTOR)
    except PWTimeoutError:
        print("[경고] 조회 응답을 기다리는 중 타임아웃 발생. 계속 진행합니다.")

    # htmx 인디케이터(#modal_loading)가 사라질 때까지 대기 (있다면)
    try:
        page.wait_for_selector("#modal_loading", state="hidden", timeout=10000)
    except PWTimeoutError:
        pass

    # DOM 스왑 이후 테이블 렌더링 여유시간
    page.wait_for_timeout(800)

    print(f"[정보] 영업소 '{office_name}' 조회 완료")


def main():
    with sync_playwright() as p:
        print(f"[정보] CDP({CDP_URL})로 크롬에 연결 중...")
        browser = p.chromium.connect_over_cdp(CDP_URL)

        # 기존 컨텍스트가 있으면 사용, 없으면 새로 생성
        context = browser.contexts[0] if browser.contexts else browser.new_context()

        # 새 탭 추가
        page = context.new_page()
        print(f"[정보] 페이지 이동: {TARGET_URL}")
        page.goto(TARGET_URL, wait_until="networkidle")

        for office_name in OFFICE_OPTIONS:
            select_office_and_search(page, office_name)
            run_collect_script()

        print("[정보] 모든 영업소 처리 완료")
        # 탭/연결은 유지하고 싶으면 아래 두 줄을 주석 처리하세요.
        # page.close()
        # browser.close()


if __name__ == "__main__":
    main()
