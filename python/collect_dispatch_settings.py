"""
고정(쉬프트)지정 페이지(/a3/view_dispatch_settings_tab1/) 표를
Chrome DevTools Protocol(CDP)로 읽어서 baecha.db(SQLite)의
dispatch_settings 테이블에 저장하는 스크립트.

요구사항 반영:
  - B성명 / B사번은 컬럼이 아니라 "행(row)"으로 저장한다.
    즉 원본 한 행(A기사 + B기사)이 DB에서는 최대 2행(A행, B행)이 된다.
  - driver_type 컬럼을 추가해서 그 행이 A인지 B인지 구분한다.
  - 사번(driver_no)이 키다. 사번이 없는 자료(A/B 각각)는 버린다.
  - 진행 로그를 stdout으로 남긴다.

실행 인수 (추가됨):
    python collect_dispatch_settings.py {new|update} {search|param} [cdp연결페이지 정보]

  - 첫째 인수 mode
      new    : DB 연결 시 기존 dispatch_settings 테이블을 백업용 이름으로
               rename 하고, 새 테이블을 만들어서 처리한다.
      update : 기존 방식 그대로. 테이블이 없으면 만들고, 있으면 그대로 사용
               (UPSERT 방식으로 driver_no 기준 갱신/삽입).

  - 둘째 인수 source
      search : 기존 방식 그대로. CDP(9223)에 새로 연결해서 TAB_TITLE로
               탭을 찾아 그 탭에 붙는다.
      param  : 크롬에 새로 연결/탭 검색을 하지 않는다. 셋째 인수로 넘겨받은
               "cdp연결페이지 정보"(해당 탭의 webSocketDebuggerUrl)로 바로
               websocket 연결한다. dispatch_settings_cdp.py 처럼 이미 같은
               탭을 열어서 조회까지 마친 상태에서 호출할 때 사용한다.

  - 셋째 인수 page_info
      source=param 일 때만 사용. 대상 탭의 webSocketDebuggerUrl 문자열.
      (예: ws://127.0.0.1:9223/devtools/page/XXXXXXXX)

사전 준비
  1) pip install websocket-client requests
  2) 크롬이 --remote-debugging-port=9223 으로 이미 떠 있고, 고정(쉬프트)지정
     페이지가 이미 로그인된 상태로 열려서 표가 보이는 상태여야 한다.
     (source=search 인 경우에만 해당. source=param 이면 이미 연결된 탭을
      그대로 재사용하므로 이 조건은 호출하는 쪽에서 보장한다.)
  3) 필요하면 아래 TAB_TITLE 을 실제 탭 제목과 맞춰 조정한다.

실행 예시
  python collect_dispatch_settings.py new search
  python collect_dispatch_settings.py update search
  python collect_dispatch_settings.py update param "ws://127.0.0.1:9223/devtools/page/XXXX"
"""
import json
import logging
import sqlite3
import sys
import time
from typing import Optional

import requests
import websocket  # pip install websocket-client

# ------------------------------------------------------------------
# 로깅 (stdout)
# ------------------------------------------------------------------
logger = logging.getLogger("dispatch_settings")
logger.setLevel(logging.INFO)
_handler = logging.StreamHandler(stream=sys.stdout)
_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
logger.addHandler(_handler)

# ------------------------------------------------------------------
# 설정
# ------------------------------------------------------------------
CDP_HOST = "127.0.0.1"
CDP_PORT = 9223
# 찾을 탭의 제목(<title>고정(쉬프트)지정</title>). 부분 일치로 찾는다. (source=search 일 때만 사용)
TAB_TITLE = "고정(쉬프트)지정"
DB_PATH = "baecha.db"


# ------------------------------------------------------------------
# 아주 단순한 CDP 클라이언트 (websocket 기반, 외부 브라우저 자동화 라이브러리 없이)
# ------------------------------------------------------------------
class CDPClient:
    def __init__(self, host: str = CDP_HOST, port: int = CDP_PORT):
        self.host = host
        self.port = port
        self.ws: Optional[websocket.WebSocket] = None  # noqa: UP045
        self._msg_id = 0

    def connect_new_tab(self, url: str = "about:blank") -> "CDPClient":
        resp = requests.put(f"http://{self.host}:{self.port}/json/new?{url}")
        resp.raise_for_status()
        tab = resp.json()
        self.ws = websocket.create_connection(tab["webSocketDebuggerUrl"], timeout=30)
        return self

    def find_tab_by_title(self, title_contains: str) -> Optional[dict]:  # noqa: UP045
        """열려 있는 탭 목록 중 title에 title_contains가 포함된 첫 탭 정보를 반환한다."""
        resp = requests.get(f"http://{self.host}:{self.port}/json")
        resp.raise_for_status()
        for tab in resp.json():
            if tab.get("type") == "page" and title_contains in tab.get("title", ""):
                return tab
        return None

    def connect_existing_tab_by_title(self, title_contains: str) -> Optional["CDPClient"]:
        """탭 제목으로 탭을 찾아 그 탭에 붙는다. 못 찾으면 None."""
        tab = self.find_tab_by_title(title_contains)
        if tab is None:
            return None
        self.ws = websocket.create_connection(tab["webSocketDebuggerUrl"], timeout=30)
        return self

    def connect_ws_url(self, ws_url: str) -> "CDPClient":
        """이미 알고 있는 webSocketDebuggerUrl로 바로 연결한다.
        (탭 검색 없이 곧바로 붙을 때 사용. source=param 용)
        """
        self.ws = websocket.create_connection(ws_url, timeout=30)
        return self

    def send(self, method: str, params: Optional[dict] = None) -> dict:  # noqa: UP045
        self._msg_id += 1
        msg_id = self._msg_id
        self.ws.send(json.dumps({"id": msg_id, "method": method, "params": params or {}}))
        while True:
            raw = self.ws.recv()
            resp = json.loads(raw)
            if resp.get("id") == msg_id:
                return resp
            # 그 외에는 이벤트 메시지이므로 무시

    def eval_js(self, expression: str):
        result = self.send(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
        )
        value = result.get("result", {}).get("result", {})
        if "value" in value:
            return value["value"]
        return None

    def close(self):
        if self.ws:
            self.ws.close()


# 현재 열려 있는 페이지에 렌더링된 표를 그대로 JS에서 딕셔너리 배열로 뽑아오는 스크립트.
# 페이지를 새로 이동/조회하지 않고, 지금 화면에 보이는 데이터를 읽는다.
# 테이블의 <td> 순서: 노선, 순번, 예비차량여부, 차량번호, 휴무, 쉬프트,
#                     A성명, A사번, B성명, B사번, 적용일자
EXTRACT_JS = r"""
(() => {
  const officeSelect = document.querySelector('#id_search_office_false');
  const office = officeSelect && officeSelect.selectedOptions.length
    ? officeSelect.selectedOptions[0].textContent.trim()
    : '';
  const rows = document.querySelectorAll('#dispatch_settings_tbody tr[data-id]');
  const out = [];
  rows.forEach(tr => {
    const tds = tr.querySelectorAll('td');
    const get = (i) => (tds[i] ? tds[i].textContent.trim() : '');
    out.push({
      source_id: tr.dataset.id || '',
      route: get(0),
      idx: get(1),
      is_reserve: get(2),
      car_number: get(3),
      rest_day: get(4),
      shift: get(5),
      a_name: get(6),
      a_no: get(7),
      b_name: get(8),
      b_no: get(9),
      apply_date: get(10),
    });
  });
  return { office, rows: out };
})()
"""


# ------------------------------------------------------------------
# SQLite
# ------------------------------------------------------------------
def init_db(conn: sqlite3.Connection) -> None:
    """테이블이 없으면 만든다. 있으면 그대로 둔다 (update 모드에서 사용)."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS dispatch_settings (
            driver_no     TEXT PRIMARY KEY,   -- 사번 = 키. 사번 없는 자료는 저장하지 않음
            driver_type   TEXT NOT NULL CHECK (driver_type IN ('A', 'B')),
            driver_name   TEXT,
            office        TEXT,               -- 조회 시 선택한 영업소명
            route         TEXT,               -- 노선
            car_number    TEXT,               -- 차량번호
            idx           INTEGER,            -- 순번
            rest_day      TEXT,               -- 휴무
            shift         TEXT,               -- 쉬프트
            apply_date    DATE,
            source_id     INTEGER NOT NULL,   -- 원본 페이지의 data-id (A/B 두 행이 같은 값을 공유)
            is_reserve    TEXT,               -- 예비차량여부 (Y/'')
            collected_at  DATETIME NOT NULL
        )
        """
    )
    conn.commit()


def backup_and_recreate_table(conn: sqlite3.Connection) -> None:
    """mode=new 용.
    기존 dispatch_settings 테이블이 있으면 dispatch_settings_backup_YYYYMMDD_HHMMSS
    이름으로 rename(백업)하고, 이어서 새 dispatch_settings 테이블을 만든다.
    """
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='dispatch_settings'"
    )
    if cur.fetchone() is not None:
        backup_name = f"dispatch_settings_{time.strftime('%Y%m%d_%H%M%S')}"
        logger.info("mode=new -> 기존 dispatch_settings 테이블을 %r (으)로 백업합니다.", backup_name)
        conn.execute(f"ALTER TABLE dispatch_settings RENAME TO {backup_name}")
        conn.commit()
    else:
        logger.info("mode=new -> 기존 dispatch_settings 테이블이 없어 백업 없이 새로 만듭니다.")
    init_db(conn)


def row_exists(conn: sqlite3.Connection, driver_no: str) -> bool:
    """driver_no(사번)가 이미 저장되어 있는지 확인한다."""
    cur = conn.execute(
        "SELECT 1 FROM dispatch_settings WHERE driver_no = ? LIMIT 1",
        (driver_no,),
    )
    return cur.fetchone() is not None


def insert_row(conn: sqlite3.Connection, row: dict, driver_type: str,
                driver_name: str, driver_no: str, office: str, collected_at: str) -> None:
    """새 사번 -> INSERT"""
    idx_val = int(row["idx"]) if row["idx"].strip().isdigit() else None
    conn.execute(
        """
        INSERT INTO dispatch_settings
            (driver_no, source_id, office, route, idx, is_reserve, car_number, rest_day, shift,
             driver_type, driver_name, apply_date, collected_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            driver_no, int(row["source_id"]), office, row["route"], idx_val, row["is_reserve"],
            row["car_number"], row["rest_day"], row["shift"],
            driver_type, driver_name, row["apply_date"], collected_at,
        ),
    )


def update_row(conn: sqlite3.Connection, row: dict, driver_type: str,
                driver_name: str, driver_no: str, office: str, collected_at: str) -> None:
    """기존 사번 -> UPDATE"""
    idx_val = int(row["idx"]) if row["idx"].strip().isdigit() else None
    conn.execute(
        """
        UPDATE dispatch_settings
        SET source_id    = ?,
            office       = ?,
            route        = ?,
            idx          = ?,
            is_reserve   = ?,
            car_number   = ?,
            rest_day     = ?,
            shift        = ?,
            driver_type  = ?,
            driver_name  = ?,
            apply_date   = ?,
            collected_at = ?
        WHERE driver_no = ?
        """,
        (
            int(row["source_id"]), office, row["route"], idx_val, row["is_reserve"],
            row["car_number"], row["rest_day"], row["shift"],
            driver_type, driver_name, row["apply_date"], collected_at,
            driver_no,
        ),
    )


def save_rows(conn: sqlite3.Connection, rows: list, office: str) -> tuple:
    """반환값: (신규 삽입 행 수, 갱신된 행 수, 사번 없어서 버려진 행 수)"""
    collected_at = time.strftime("%Y-%m-%d %H:%M:%S")
    inserted = 0
    updated = 0
    dropped = 0
    for row in rows:
        if not row.get("source_id"):
            continue
        for driver_type, name_key, no_key in (("A", "a_name", "a_no"), ("B", "b_name", "b_no")):
            name = row[name_key]
            no = row[no_key].strip()
            if not name and not no:
                # A/B 둘 다 비어 있는 칸(예비차량 등)은 애초에 대상 아님
                continue
            if not no:
                dropped += 1
                logger.warning(
                    "사번 없음 -> 버림: source_id=%s route=%s idx=%s driver_type=%s name=%r",
                    row["source_id"], row["route"], row["idx"], driver_type, name,
                )
                continue

            if row_exists(conn, no):
                update_row(conn, row, driver_type, name, no, office, collected_at)
                updated += 1
            else:
                insert_row(conn, row, driver_type, name, no, office, collected_at)
                inserted += 1

    conn.commit()
    return inserted, updated, dropped


# ------------------------------------------------------------------
# 메인
# ------------------------------------------------------------------
def main(mode: str = "update", source: str = "search", page_info: Optional[str] = None):  # noqa: UP045
    """
    mode      : "new" | "update"
    source    : "search" | "param"
    page_info : source="param" 일 때 사용할 webSocketDebuggerUrl
    """
    if mode not in ("new", "update"):
        logger.error("잘못된 첫번째 인수 mode=%r ('new' 또는 'update' 여야 함)", mode)
        sys.exit(1)
    if source not in ("search", "param"):
        logger.error("잘못된 두번째 인수 source=%r ('search' 또는 'param' 이어야 함)", source)
        sys.exit(1)
    if source == "param" and not page_info:
        logger.error("source='param' 인데 셋째 인수(cdp연결페이지 정보)가 없습니다.")
        sys.exit(1)

    logger.info("시작: DB=%s mode=%s source=%s", DB_PATH, mode, source)
    conn = sqlite3.connect(DB_PATH)

    if mode == "new":
        backup_and_recreate_table(conn)
    else:
        init_db(conn)

    client = CDPClient()

    if source == "search":
        logger.info("탭 검색 중 (title 포함: %r, port=%d)", TAB_TITLE, CDP_PORT)
        if client.connect_existing_tab_by_title(TAB_TITLE) is None:
            logger.error(
                "탭 제목에 %r 이(가) 포함된 탭을 찾지 못했습니다. "
                "고정(쉬프트)지정 페이지가 열려 있는지 확인하세요.",
                TAB_TITLE,
            )
            sys.exit(1)
        logger.info("탭에 연결했습니다.")
    else:  # source == "param"
        logger.info("전달받은 cdp연결페이지 정보로 바로 연결합니다: %s", page_info)
        client.connect_ws_url(page_info)
        logger.info("탭에 연결했습니다.")

    inserted = updated = dropped = 0
    try:
        result = client.eval_js(EXTRACT_JS) or {}
        office = result.get("office", "")
        rows = result.get("rows", [])
        logger.info("현재 화면(영업소=%r)에서 %d행을 읽었습니다.", office, len(rows))
        inserted, updated, dropped = save_rows(conn, rows, office)
        print(f"신규 {inserted}건, 갱신 {updated}건, 버림 {dropped}건")
    except Exception:
        logger.exception("스크립트 실행 중 오류가 발생했습니다.")
        raise
    finally:
        client.close()
        conn.close()

    logger.info(
        "완료. 신규 %d행, 갱신 %d행, 사번없어 버려진 행 %d개 -> %s (table: dispatch_settings)",
        inserted, updated, dropped, DB_PATH,
    )


if __name__ == "__main__":
    _mode = sys.argv[1] if len(sys.argv) > 1 else "update"
    _source = sys.argv[2] if len(sys.argv) > 2 else "search"
    _page_info = sys.argv[3] if len(sys.argv) > 3 else None
    main(_mode, _source, _page_info)
