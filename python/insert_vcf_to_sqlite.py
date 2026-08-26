#!/usr/bin/env python3

import re
import quopri
import sqlite3
import sys
from datetime import datetime


# ============================================================
# 설정
# ============================================================

DEFAULT_DB_PATH = "contacts.db"


# ============================================================
# SQLite 테이블
# ============================================================

DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS contacts (
    tel_section TEXT NOT NULL,
    telnumber   TEXT NOT NULL,
    fn_old      TEXT NOT NULL,
    fn_new      TEXT,
    insert_date TEXT NOT NULL,
    update_date TEXT,
    PRIMARY KEY (tel_section, telnumber)
);
"""


# ============================================================
# vCard 값 디코딩
# ============================================================

def decode_vcard_value(value: str, params: str = "") -> str:
    """
    vCard 2.1의 값을 디코딩한다.

    예:
        FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:=EC=9D=B4...

    -> 이충열

    ENCODING=QUOTED-PRINTABLE이면 Quoted-Printable 디코딩 후
    CHARSET에 따라 문자열로 변환한다.

    ENCODING이 없으면 원문을 그대로 반환한다.
    """

    encoding_match = re.search(
        r"(?:^|;)ENCODING=([^;:]+)",
        params,
        re.IGNORECASE
    )

    charset_match = re.search(
        r"(?:^|;)CHARSET=([^;:]+)",
        params,
        re.IGNORECASE
    )

    encoding = (
        encoding_match.group(1).upper()
        if encoding_match
        else ""
    )

    charset = (
        charset_match.group(1)
        if charset_match
        else "utf-8"
    )

    # Quoted-Printable
    if encoding == "QUOTED-PRINTABLE":
        raw = quopri.decodestring(value)

        try:
            return raw.decode(charset)
        except (UnicodeDecodeError, LookupError):
            return raw.decode("utf-8", errors="replace")

    # 현재 목적에서는 그 외 encoding은 그대로 처리
    return value


# ============================================================
# 전화번호 정규화
# ============================================================

def normalize_phone(phone: str) -> str:
    """
    전화번호에서 숫자만 남긴다.

    예:
        010-1234-5678
        -> 01012345678
    """

    if not phone:
        return ""

    return re.sub(r"\D", "", phone)


# ============================================================
# vCard folded line 처리
# ============================================================

def parse_vcard_lines(text: str):
    """
    vCard의 folded line을 처리한다.

    vCard에서는 이전 줄의 내용이 다음 줄로 이어질 경우
    다음 줄이 공백 또는 TAB으로 시작할 수 있다.

    예:
        FN;...:=EC=9D=B4=...
         =EC=9A=B0=...

    -> 하나의 논리적인 줄로 합친다.
    """

    physical_lines = text.splitlines()

    logical_lines = []

    for line in physical_lines:

        if line.startswith((" ", "\t")) and logical_lines:
            logical_lines[-1] += line[1:]

        else:
            logical_lines.append(line)

    return logical_lines


# ============================================================
# VCF 연락처 분리
# ============================================================

def parse_vcards(text: str):
    """
    BEGIN:VCARD ~ END:VCARD 단위로 연락처를 분리한다.
    """

    lines = parse_vcard_lines(text)

    cards = []
    current = None

    for line in lines:

        line = line.rstrip("\r\n")

        if line.upper() == "BEGIN:VCARD":

            current = [line]

        elif line.upper() == "END:VCARD":

            if current is not None:

                current.append(line)
                cards.append(current)

                current = None

        elif current is not None:

            current.append(line)

    return cards


# ============================================================
# vCard property 분석
# ============================================================

def parse_property(line: str):
    """
    vCard property를 다음 세 부분으로 분리한다.

        FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:이름

    반환:
        property_name
        params
        value
    """

    if ":" not in line:
        return None, "", ""

    left, value = line.split(":", 1)

    parts = left.split(";")

    property_name = parts[0].upper()

    params = ";".join(parts[1:])

    return property_name, params, value


# ============================================================
# 전화번호 종류 추출
# ============================================================

def get_tel_section(params: str) -> str:
    """
    TEL property의 전화번호 종류를 반환한다.

    예:

        TEL;CELL;PREF
        -> cell

        TEL;WORK
        -> work

        TEL;HOME
        -> home

    PREF는 전화번호 종류가 아니므로 무시한다.
    """

    valid_sections = {
        "CELL",
        "WORK",
        "HOME",
        "FAX",
        "PAGER",
        "VOICE",
    }

    for param in params.split(";"):

        param_upper = param.upper()

        if param_upper in valid_sections:
            return param.lower()

    return "unknown"


# ============================================================
# 한 개의 VCF 연락처 추출
# ============================================================

def extract_contact(card_lines):
    """
    하나의 vCard에서 다음 정보를 추출한다.

        FN
        N
        TEL

    반환:

        fn
        n_value
        phones

    phones:
        [
            ("cell", "01012345678"),
            ("work", "0212345678")
        ]
    """

    fn = None
    n_value = None
    phones = []

    for line in card_lines:

        property_name, params, value = parse_property(line)

        if property_name == "FN":

            decoded = decode_vcard_value(
                value,
                params
            )

            fn = decoded.strip()

        elif property_name == "N":

            decoded = decode_vcard_value(
                value,
                params
            )

            n_value = decoded.strip()

        elif property_name == "TEL":

            decoded = decode_vcard_value(
                value,
                params
            )

            phone = decoded.strip()

            if not phone:
                continue

            tel_section = get_tel_section(params)

            phone = normalize_phone(phone)

            if phone:
                phones.append(
                    (
                        tel_section,
                        phone
                    )
                )

    return fn, n_value, phones


# ============================================================
# SQLite DB 초기화
# ============================================================

def init_database(db_path: str):

    conn = sqlite3.connect(db_path)

    conn.execute(DB_SCHEMA)

    conn.commit()

    return conn


# ============================================================
# FN 없는 연락처 로그
# ============================================================

def print_missing_fn_warning(
    n_value,
    phones
):
    """
    FN이 없는 연락처를 표준출력으로 기록한다.
    """

    print(
        "[WARNING] FN 없음",
        file=sys.stdout
    )

    print(
        f"    N   : {n_value!r}",
        file=sys.stdout
    )

    if phones:

        for tel_section, telnumber in phones:

            print(
                f"    TEL : {tel_section} {telnumber}",
                file=sys.stdout
            )

    else:

        print(
            "    TEL : 없음",
            file=sys.stdout
        )

    print(
        "-" * 60,
        file=sys.stdout
    )


# ============================================================
# TEL 없는 연락처 로그
# ============================================================

def print_missing_tel_warning(fn, n_value):
    """
    TEL이 없는 연락처를 표준출력으로 기록한다.
    """

    print(
        "[WARNING] TEL 없음",
        file=sys.stdout
    )

    print(
        f"    FN  : {fn!r}",
        file=sys.stdout
    )

    print(
        f"    N   : {n_value!r}",
        file=sys.stdout
    )

    print(
        "-" * 60,
        file=sys.stdout
    )


# ============================================================
# VCF -> SQLite
# ============================================================

def import_vcf(vcf_path: str, db_path: str):

    # --------------------------------------------------------
    # VCF 읽기
    # --------------------------------------------------------

    try:

        with open(
            vcf_path,
            "r",
            encoding="utf-8-sig"
        ) as f:

            text = f.read()

    except FileNotFoundError:

        print(
            f"오류: 파일을 찾을 수 없습니다: {vcf_path}",
            file=sys.stderr
        )

        sys.exit(1)

    except UnicodeDecodeError:

        print(
            f"오류: UTF-8로 읽을 수 없습니다: {vcf_path}",
            file=sys.stderr
        )

        sys.exit(1)


    # --------------------------------------------------------
    # VCF 분리
    # --------------------------------------------------------

    cards = parse_vcards(text)


    # --------------------------------------------------------
    # DB 연결
    # --------------------------------------------------------

    conn = init_database(db_path)


    # --------------------------------------------------------
    # 현재 일시
    # --------------------------------------------------------

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


    # --------------------------------------------------------
    # 통계
    # --------------------------------------------------------

    inserted = 0
    duplicated = 0
    missing_fn = 0
    missing_tel = 0


    # --------------------------------------------------------
    # 연락처 처리
    # --------------------------------------------------------

    for card_number, card in enumerate(
        cards,
        start=1
    ):

        fn, n_value, phones = extract_contact(card)


        # ----------------------------------------------------
        # FN 없음
        # ----------------------------------------------------

        if not fn:

            missing_fn += 1

            print(
                f"[VCARD #{card_number}]",
                file=sys.stdout
            )

            print_missing_fn_warning(
                n_value,
                phones
            )

            # FN 없는 연락처는 DB에 저장하지 않는다.
            continue


        # ----------------------------------------------------
        # TEL 없음
        # ----------------------------------------------------

        if not phones:

            missing_tel += 1

            print(
                f"[VCARD #{card_number}]",
                file=sys.stdout
            )

            print_missing_tel_warning(
                fn,
                n_value
            )

            # TEL 없는 연락처는 현재 DB 구조상 저장할 수 없다.
            continue


        # ----------------------------------------------------
        # 전화번호 저장
        # ----------------------------------------------------

        for tel_section, telnumber in phones:

            cursor = conn.execute(
                """
                INSERT OR IGNORE INTO contacts
                (
                    tel_section,
                    telnumber,
                    fn_old,
                    fn_new,
                    insert_date,
                    update_date
                )
                VALUES (?, ?, ?, NULL, ?, NULL)
                """,
                (
                    tel_section,
                    telnumber,
                    fn,
                    now
                )
            )


            if cursor.rowcount == 1:

                inserted += 1

            else:

                duplicated += 1


    # --------------------------------------------------------
    # 저장
    # --------------------------------------------------------

    conn.commit()

    conn.close()


    # --------------------------------------------------------
    # 결과
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("VCF → SQLite 처리 완료")
    print("=" * 60)

    print(
        f"VCF 연락처 수 : {len(cards):,}"
    )

    print(
        f"새로 저장     : {inserted:,}"
    )

    print(
        f"중복/기존     : {duplicated:,}"
    )

    print(
        f"FN 없음       : {missing_fn:,}"
    )

    print(
        f"TEL 없음      : {missing_tel:,}"
    )

    print(
        f"SQLite 파일   : {db_path}"
    )

    print("=" * 60)


# ============================================================
# 사용법
# ============================================================

def print_usage():

    print(
        "사용법:"
    )

    print()

    print(
        "    python vcf_to_sqlite.py 입력.vcf"
    )

    print()

    print(
        "또는"
    )

    print()

    print(
        "    python vcf_to_sqlite.py 입력.vcf contacts.db"
    )


# ============================================================
# 프로그램 시작
# ============================================================

def main():

    if len(sys.argv) < 2:

        print_usage()

        sys.exit(1)


    vcf_path = sys.argv[1]


    if len(sys.argv) >= 3:

        db_path = sys.argv[2]

    else:

        db_path = DEFAULT_DB_PATH


    import_vcf(
        vcf_path,
        db_path
    )


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":

    main()