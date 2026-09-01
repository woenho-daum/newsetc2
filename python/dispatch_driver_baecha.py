#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기사별 배차현황 수집 프로그램
=============================
1) 이미 9223 포트로 원격 디버깅이 열려있는 크롬(CDP)에 Playwright로 접속
2) https://bumil.mobilhi.com/a3/view_dispatch_daily/ 페이지를 로드
3) 조회 조건(날짜/사업소/노선/평시-방학-공휴일/평일-토-일) 세팅 후 [조회] 클릭
4) "기사별 배차현황" 탭(htmx) 클릭 -> 결과 테이블 로드 대기
5) <table id="table"> 의 내용을 파싱하여 SQLite DB에 적재

필요 패키지
-----------
    pip install playwright beautifulsoup4
    playwright install chromium   # (CDP로 기존 크롬에 붙는 경우엔 브라우저 바이너리 설치는 필수 아님)

사용 예
-------
    python dispatch_driver_baeche.py --date 2026-08-31 --db dispatch.db
    python dispatch_driver_baeche.py                     # 날짜 생략 시 오늘 날짜, db 기본값 dispatch.db
"""

import argparse
import re
import sqlite3
import sys
from datetime import datetime, date
from pathlib import Path

from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# ----------------------------------------------------------------------
# 설정값
# ----------------------------------------------------------------------
TARGET_URL = "https://bumil.mobilhi.com/a3/view_dispatch_daily/"
CDP_PORT_DEFAULT = 9223

# 요청서에 명시된 폼 필드 select box 값
FORM_DEFAULTS = {
    "office": "1",          # 1=본사, 2=광명, 3=법원, 4=독산, 5=심야, 9=출퇴근
    "route": "",             # ""=전체, 1=5413, 2=5525, 3=5537, 4=5617, 5=5619, 6=5620
    "run_flag": "370",       # 370=평시, 371=방학, 372=공휴일, 373=특별편성
    "week_flag": "평일",      # 평일, 토요일, 일요일, 평일2, 토요2, 일요2
}

# 결과 테이블의 고정 컬럼 수 (앞: 사번/성명/휴무일, 뒤: 오전~근무일 요약 9개)
N_FIXED_HEAD = 3
SUMMARY_HEADERS = ["오전", "오후", "정상", "SH오전", "SH오후", "단축", "A->P", "초과", "근무일"]
N_SUMMARY = len(SUMMARY_HEADERS)


# ----------------------------------------------------------------------
# 1. 크롬 CDP 접속 & 폼 조작
# ----------------------------------------------------------------------
def connect_to_chrome(cdp_port: int):
    """이미 열려있는 크롬(원격 디버깅 포트)에 Playwright로 접속한다."""
    pw = sync_playwright().start()
    browser = pw.chromium.connect_over_cdp(f"http://localhost:{cdp_port}")
    context = browser.contexts[0] if browser.contexts else browser.new_context()
    page = context.new_page() if not context.pages else context.pages[0]
    return pw, browser, page


def fill_search_form(page, query_date: str, office: str, route: str,
                      run_flag: str, week_flag: str):
    """조회 조건 폼을 채우고 [조회] 버튼을 클릭한다."""
    page.goto(TARGET_URL, wait_until="domcontentloaded")

    # 2) 날짜
    page.wait_for_selector("#id_extra_date_required", timeout=20000)
    page.fill("#id_extra_date_required", query_date)

    # 3) 사업소 (본사=1)
    page.select_option("#id_search_office", value=office)

    # 4) 노선 (전체="")
    page.select_option("#id_empty_route_false", value=route)

    # 5) 평시/방학/공휴일/특별편성 (평시=370)
    page.select_option("#id_select_run_flag", value=run_flag)

    # 6) 평일/토요일/일요일 등
    page.select_option("#id_select_week_flag", value=week_flag)

    # 7) 조회 버튼 클릭 (날짜 입력창과 같은 form 내부의 submit 버튼을 찾는다)
    submit_btn = page.locator("#id_extra_date_required").locator(
        "xpath=ancestor::form[1]//button[@type='submit']"
    )
    submit_btn.first.click()

    # 폼 제출 후 화면 안정화 대기
    page.wait_for_load_state("networkidle", timeout=30000)


def open_driver_dispatch_tab(page):
    """'기사별 배차현황' 탭(htmx)을 클릭하고 결과 테이블이 뜰 때까지 대기한다."""
    tab_label = page.locator("label[for='btnradio2']")
    tab_label.wait_for(state="visible", timeout=20000)
    tab_label.click()

    # htmx 요청 완료 대기: 네트워크 idle + 실제 테이블 렌더링 대기
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_selector("#tab_load table#table tbody tr", timeout=30000)


def extract_table_html(page) -> str:
    """결과 테이블 전체(outerHTML)를 문자열로 가져온다."""
    table_el = page.locator("#tab_load table#table")
    return table_el.evaluate("el => el.outerHTML")


def fetch_dispatch_html(cdp_port: int, query_date: str, office: str,
                         route: str, run_flag: str, week_flag: str) -> str:
    """크롬 CDP에 접속하여 조회 -> 탭 클릭 -> 결과 테이블 HTML을 반환."""
    pw, browser, page = connect_to_chrome(cdp_port)
    try:
        fill_search_form(page, query_date, office, route, run_flag, week_flag)
        open_driver_dispatch_tab(page)
        return extract_table_html(page)
    finally:
        # CDP로 접속한 기존 크롬은 닫지 않고 연결만 해제한다.
        pw.stop()


# ----------------------------------------------------------------------
# 2. 테이블 HTML 파싱
# ----------------------------------------------------------------------
def parse_dispatch_table(table_html: str):
    """
    <table id="table"> HTML을 파싱하여
      - summary_rows: [{사번,성명,휴무일,오전,오후,정상,SH오전,SH오후,단축,'A->P',초과,근무일}, ...]
      - daily_rows:   [{사번,성명,day,weekday,value,bg_color}, ...]
    를 반환한다.
    """
    soup = BeautifulSoup(table_html, "html.parser")
    table = soup if soup.name == "table" else soup.find("table")
    if table is None:
        raise ValueError("테이블(<table id='table'>)을 찾을 수 없습니다.")

    thead = table.find("thead")
    headers = [th.get_text(strip=True) for th in thead.find_all("th")]

    n_days = len(headers) - N_FIXED_HEAD - N_SUMMARY
    if n_days <= 0:
        raise ValueError(f"예상치 못한 컬럼 구성입니다. headers={headers}")

    day_headers = headers[N_FIXED_HEAD:N_FIXED_HEAD + n_days]
    # "01 토" -> (1, "토")
    day_defs = []
    for h in day_headers:
        m = re.match(r"^(\d{1,2})\s*(.*)$", h)
        if m:
            day_defs.append((int(m.group(1)), m.group(2)))
        else:
            day_defs.append((None, h))

    tbody = table.find("tbody")
    rows = tbody.find_all("tr", recursive=False)

    summary_rows = []
    daily_rows = []

    for tr in rows:
        tds = tr.find_all("td", recursive=False)
        if len(tds) < N_FIXED_HEAD + N_SUMMARY:
            continue  # 형식이 다른 행(합계행 등)은 건너뜀

        emp_no = tds[0].get_text(strip=True)
        name = tds[1].get_text(strip=True)
        day_off = tds[2].get_text(strip=True)

        # 일자별 값
        for i in range(n_days):
            td = tds[N_FIXED_HEAD + i]
            day_no, weekday = day_defs[i]
            value = td.get_text(strip=True)
            style = td.get("style", "")
            bg_match = re.search(r"background-color:\s*([^;]+)", style)
            bg_color = bg_match.group(1).strip() if bg_match else None
            daily_rows.append({
                "employee_no": emp_no,
                "name": name,
                "day": day_no,
                "weekday": weekday,
                "value": value,
                "bg_color": bg_color,
            })

        # 요약(오전~근무일) 값
        summary_vals = [td.get_text(strip=True)
                         for td in tds[N_FIXED_HEAD + n_days:N_FIXED_HEAD + n_days + N_SUMMARY]]
        row = {"employee_no": emp_no, "name": name, "day_off": day_off}
        row.update(dict(zip(SUMMARY_HEADERS, summary_vals)))
        summary_rows.append(row)

    return summary_rows, daily_rows


# ----------------------------------------------------------------------
# 3. SQLite 적재
# ----------------------------------------------------------------------
def init_db(db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS dispatch_summary (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        query_date    TEXT NOT NULL,
        office        TEXT,
        run_flag      TEXT,
        week_flag     TEXT,
        employee_no   TEXT NOT NULL,
        name          TEXT,
        day_off       TEXT,
        am_count      TEXT,
        pm_count      TEXT,
        normal_count  TEXT,
        sh_am         TEXT,
        sh_pm         TEXT,
        short_count   TEXT,
        a_to_p        TEXT,
        excess        TEXT,
        work_days     TEXT,
        created_at    TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(query_date, employee_no)
    );

    CREATE TABLE IF NOT EXISTS dispatch_daily (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        query_date    TEXT NOT NULL,
        employee_no   TEXT NOT NULL,
        name          TEXT,
        day           INTEGER,
        weekday       TEXT,
        value         TEXT,
        bg_color      TEXT,
        created_at    TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(query_date, employee_no, day)
    );

    CREATE INDEX IF NOT EXISTS idx_daily_emp ON dispatch_daily(employee_no);
    CREATE INDEX IF NOT EXISTS idx_daily_date ON dispatch_daily(query_date);
    """)
    conn.commit()
    return conn


def summary_row_exists(cur, query_date: str, employee_no: str) -> bool:
    cur.execute(
        "SELECT 1 FROM dispatch_summary WHERE query_date = ? AND employee_no = ? LIMIT 1",
        (query_date, employee_no),
    )
    return cur.fetchone() is not None


def insert_summary_row(cur, params: dict):
    cur.execute("""
        INSERT INTO dispatch_summary
            (query_date, office, run_flag, week_flag, employee_no, name, day_off,
             am_count, pm_count, normal_count, sh_am, sh_pm, short_count, a_to_p, excess, work_days)
        VALUES (:query_date, :office, :run_flag, :week_flag, :employee_no, :name, :day_off,
                :am_count, :pm_count, :normal_count, :sh_am, :sh_pm, :short_count, :a_to_p, :excess, :work_days)
    """, params)


def update_summary_row(cur, params: dict):
    cur.execute("""
        UPDATE dispatch_summary
           SET office = :office,
               run_flag = :run_flag,
               week_flag = :week_flag,
               name = :name,
               day_off = :day_off,
               am_count = :am_count,
               pm_count = :pm_count,
               normal_count = :normal_count,
               sh_am = :sh_am,
               sh_pm = :sh_pm,
               short_count = :short_count,
               a_to_p = :a_to_p,
               excess = :excess,
               work_days = :work_days,
               created_at = datetime('now','localtime')
         WHERE query_date = :query_date AND employee_no = :employee_no
    """, params)


def daily_row_exists(cur, query_date: str, employee_no: str, day) -> bool:
    cur.execute(
        "SELECT 1 FROM dispatch_daily WHERE query_date = ? AND employee_no = ? AND day IS ? LIMIT 1",
        (query_date, employee_no, day),
    )
    return cur.fetchone() is not None


def insert_daily_row(cur, params: dict):
    cur.execute("""
        INSERT INTO dispatch_daily (query_date, employee_no, name, day, weekday, value, bg_color)
        VALUES (:query_date, :employee_no, :name, :day, :weekday, :value, :bg_color)
    """, params)


def update_daily_row(cur, params: dict):
    cur.execute("""
        UPDATE dispatch_daily
           SET name = :name,
               weekday = :weekday,
               value = :value,
               bg_color = :bg_color,
               created_at = datetime('now','localtime')
         WHERE query_date = :query_date AND employee_no = :employee_no AND day IS :day
    """, params)


def save_to_sqlite(conn, query_date: str, office: str, run_flag: str, week_flag: str,
                    summary_rows, daily_rows):
    """
    행 단위로 존재 여부를 먼저 확인한 뒤,
      - 이미 있으면 update_* 함수를 호출
      - 없으면 insert_* 함수를 호출
    한다. (일괄 DELETE 후 재삽입 방식 대신 명시적 UPDATE/INSERT 분기)
    """
    cur = conn.cursor()
    stats = {"summary_inserted": 0, "summary_updated": 0, "daily_inserted": 0, "daily_updated": 0}

    for r in summary_rows:
        params = {
            "query_date": query_date, "office": office, "run_flag": run_flag, "week_flag": week_flag,
            "employee_no": r["employee_no"], "name": r["name"], "day_off": r["day_off"],
            "am_count": r["오전"], "pm_count": r["오후"], "normal_count": r["정상"],
            "sh_am": r["SH오전"], "sh_pm": r["SH오후"], "short_count": r["단축"],
            "a_to_p": r["A->P"], "excess": r["초과"], "work_days": r["근무일"],
        }
        if summary_row_exists(cur, query_date, params["employee_no"]):
            update_summary_row(cur, params)
            stats["summary_updated"] += 1
        else:
            insert_summary_row(cur, params)
            stats["summary_inserted"] += 1

    for r in daily_rows:
        params = {"query_date": query_date, **r}
        if daily_row_exists(cur, query_date, params["employee_no"], params["day"]):
            update_daily_row(cur, params)
            stats["daily_updated"] += 1
        else:
            insert_daily_row(cur, params)
            stats["daily_inserted"] += 1

    conn.commit()
    return stats


# ----------------------------------------------------------------------
# 4. 로컬 HTML 파일로부터 적재 (테스트/오프라인용)
# ----------------------------------------------------------------------
def load_from_html_file(html_path: str, query_date: str, office: str, run_flag: str,
                         week_flag: str, db_path: str):
    """이미 저장된 결과 HTML 파일을 읽어 바로 파싱/적재한다 (브라우저 없이 테스트할 때 사용)."""
    html = Path(html_path).read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="table")
    if table is None:
        raise ValueError(f"{html_path} 에서 <table id='table'> 를 찾지 못했습니다.")

    summary_rows, daily_rows = parse_dispatch_table(str(table))
    conn = init_db(db_path)
    stats = save_to_sqlite(conn, query_date[:7], office, run_flag, week_flag, summary_rows, daily_rows)
    conn.close()
    return stats


# ----------------------------------------------------------------------
# 5. CLI
# ----------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="기사별 배차현황 수집 -> SQLite 적재")
    parser.add_argument("--date", default=date.today().isoformat(),
                         help="조회 일자 (YYYY-MM-DD), 기본값: 오늘")
    parser.add_argument("--office", default=FORM_DEFAULTS["office"], help="사업소 코드 (기본: 1=본사)")
    parser.add_argument("--route", default=FORM_DEFAULTS["route"], help="노선 코드 (기본: 빈값=전체)")
    parser.add_argument("--run-flag", default=FORM_DEFAULTS["run_flag"], help="평시/방학/공휴일 코드 (기본: 370=평시)")
    parser.add_argument("--week-flag", default=FORM_DEFAULTS["week_flag"], help="평일/토요일/일요일 (기본: 평일)")
    parser.add_argument("--db", default="dispatch.db", help="SQLite 파일 경로")
    parser.add_argument("--cdp-port", type=int, default=CDP_PORT_DEFAULT, help="크롬 원격 디버깅 포트")
    parser.add_argument("--from-html", default=None,
                         help="(테스트용) 이미 저장된 결과 HTML 파일 경로. 지정 시 브라우저 접속 없이 이 파일을 파싱/적재")
    args = parser.parse_args()

    # 날짜 형식 검증
    try:
        datetime.strptime(args.date, "%Y-%m-%d")
    except ValueError:
        print(f"[오류] --date 형식이 올바르지 않습니다: {args.date} (예: 2026-08-31)", file=sys.stderr)
        sys.exit(1)

    if args.from_html:
        stats = load_from_html_file(
            args.from_html, args.date, args.office, args.run_flag, args.week_flag, args.db
        )
        print(f"[완료] HTML 파일에서 적재 -> {args.db}")
        print(f"  summary: insert {stats['summary_inserted']}건 / update {stats['summary_updated']}건")
        print(f"  daily  : insert {stats['daily_inserted']}건 / update {stats['daily_updated']}건")
        return

    print(f"[진행] 크롬(CDP:{args.cdp_port})에 접속하여 {args.date} 배차현황 조회 중...")
    try:
        table_html = fetch_dispatch_html(
            args.cdp_port, args.date, args.office, args.route, args.run_flag, args.week_flag
        )
    except PWTimeoutError as e:
        print(f"[오류] 페이지 로드/응답 대기 시간 초과: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[오류] 크롬 CDP 접속 또는 조회 중 문제가 발생했습니다: {e}", file=sys.stderr)
        print(f"  -> 크롬이 --remote-debugging-port={args.cdp_port} 옵션으로 실행 중인지 확인하세요.", file=sys.stderr)
        sys.exit(1)

    summary_rows, daily_rows = parse_dispatch_table(table_html)
    if not summary_rows:
        print("[경고] 파싱된 데이터가 없습니다. 조회 조건 또는 페이지 구조를 확인하세요.", file=sys.stderr)

    conn = init_db(args.db)
    stats = save_to_sqlite(conn, args.date[:7], args.office, args.run_flag, args.week_flag, summary_rows, daily_rows)
    conn.close()

    print(f"[완료] 적재 -> {args.db}")
    print(f"  summary: insert {stats['summary_inserted']}건 / update {stats['summary_updated']}건")
    print(f"  daily  : insert {stats['daily_inserted']}건 / update {stats['daily_updated']}건")


if __name__ == "__main__":
    main()
