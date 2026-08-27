#!/usr/bin/env python3

import quopri
import re
import sqlite3
import sys
from datetime import datetime

# ============================================================
# 기본 설정
# ============================================================

DEFAULT_DB_PATH = "baecha.db"


# ============================================================
# SQLite 테이블
# ============================================================

DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS contacts (
    tel_section TEXT NOT NULL,
    tel_number  TEXT NOT NULL,
    fn_front    TEXT NOT NULL,
    fn_old      TEXT NOT NULL,
    fn_new      TEXT,
    insert_date TEXT NOT NULL,
    update_date TEXT,
    PRIMARY KEY (tel_section, tel_number)
);
"""


# ============================================================
# vCard 값 디코딩
# ============================================================

def decode_vcard_value(value: str, params: str = "") -> str:
    """
    vCard 2.1의 property 값을 디코딩한다.

    예:
        FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:=EC=9D=B4...

    -> 이충열

    지원:
        ENCODING=QUOTED-PRINTABLE

    CHARSET이 없으면 UTF-8을 기본으로 사용한다.
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

    # --------------------------------------------------------
    # Quoted-Printable
    # --------------------------------------------------------

    if encoding == "QUOTED-PRINTABLE":

        raw = quopri.decodestring(value)

        try:
            return raw.decode(charset)

        except (UnicodeDecodeError, LookupError):

            return raw.decode(
                "utf-8",
                errors="replace"
            )

    # --------------------------------------------------------
    # 별도 encoding이 없는 경우
    # --------------------------------------------------------

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

        +82-10-1234-5678
        -> 821012345678
    """

    if not phone:
        return ""

    return re.sub(r"\D", "", phone)


# ============================================================
# vCard folded line + Quoted-Printable soft line break
# ============================================================

def parse_vcard_lines(text: str):
    """
    vCard의 물리적인 줄을 논리적인 property 줄로 합친다.

    1. 일반적인 vCard folded line

       FN:홍길
        동

       -> FN:홍길동


    2. Quoted-Printable soft line break

       FN;...:=EA=B8=88=EC=...
       =99=A9=EC=8B=A4

       -> FN;...:=EA=B8=88=EC=...=99=A9=EC=8B=A4

    Quoted-Printable에서 줄 마지막 '='는
    실제 데이터가 아니라 줄 연결 표시이므로 제거한다.
    """

    physical_lines = text.splitlines()

    logical_lines = []

    for line in physical_lines:

        # ----------------------------------------------------
        # 첫 번째 줄
        # ----------------------------------------------------

        if not logical_lines:

            logical_lines.append(line)

            continue


        # ----------------------------------------------------
        # vCard folded line
        #
        # 다음 줄이 공백 또는 TAB으로 시작
        # ----------------------------------------------------

        if line.startswith((" ", "\t")):

            logical_lines[-1] += line[1:]

            continue


        # ----------------------------------------------------
        # Quoted-Printable soft line break
        #
        # 이전 줄의 마지막 문자가 '='이면
        # 다음 줄과 연결한다.
        # ----------------------------------------------------

        if logical_lines[-1].endswith("="):

            logical_lines[-1] = (
                logical_lines[-1][:-1] + line
            )

            continue


        # ----------------------------------------------------
        # 일반적인 새로운 property
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # BEGIN:VCARD
        # ----------------------------------------------------

        if line.upper() == "BEGIN:VCARD":

            current = [line]

            continue


        # ----------------------------------------------------
        # END:VCARD
        # ----------------------------------------------------

        if line.upper() == "END:VCARD":

            if current is not None:

                current.append(line)

                cards.append(current)

                current = None

            continue


        # ----------------------------------------------------
        # 연락처 내부
        # ----------------------------------------------------

        if current is not None:

            current.append(line)


    return cards


# ============================================================
# vCard property 분리
# ============================================================

def parse_property(line: str):
    """
    vCard property를:

        이름
        parameter
        value

    로 분리한다.

    예:

        FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:=EC=9D...

    반환:

        "FN"
        "CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE"
        "=EC=9D..."
    """

    if ":" not in line:

        return None, "", ""


    left, value = line.split(":", 1)

    parts = left.split(";")

    property_name = parts[0].upper()

    params = ";".join(parts[1:])

    return property_name, params, value


# ============================================================
# TEL 종류 추출
# ============================================================

def get_tel_section(params: str) -> str:
    """
    TEL의 전화번호 종류를 추출한다.

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
    하나의 vCard에서:

        FN
        N
        TEL

    을 추출한다.

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


        # ----------------------------------------------------
        # FN
        # ----------------------------------------------------

        if property_name == "FN":

            decoded = decode_vcard_value(
                value,
                params
            )

            fn = decoded.strip()


        # ----------------------------------------------------
        # N
        # ----------------------------------------------------

        elif property_name == "N":

            decoded = decode_vcard_value(
                value,
                params
            )

            n_value = decoded.strip()


        # ----------------------------------------------------
        # TEL
        # ----------------------------------------------------

        elif property_name == "TEL":

            decoded = decode_vcard_value(
                value,
                params
            )

            phone = decoded.strip()


            if not phone:

                continue


            tel_section = get_tel_section(
                params
            )


            phone = normalize_phone(
                phone
            )


            if phone:

                phones.append(
                    (
                        tel_section,
                        phone
                    )
                )


    return fn, n_value, phones


# ============================================================
# SQLite 초기화
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
    card_number,
    n_value,
    phones
):
    """
    FN이 없는 연락처를 표준출력한다.
    """

    print(
        f"[WARNING] VCARD #{card_number} : FN 없음"
    )

    print(
        f"    N : {n_value!r}"
    )


    if phones:

        for tel_section, tel_number in phones:

            print(
                f"    TEL : "
                f"{tel_section} "
                f"{tel_number}"
            )

    else:

        print(
            "    TEL : 없음"
        )


    print(
        "-" * 60
    )


# ============================================================
# TEL 없는 연락처 로그
# ============================================================

def print_missing_tel_warning(
    card_number,
    fn,
    n_value
):
    """
    TEL이 없는 연락처를 표준출력한다.
    """

    print(
        f"[WARNING] VCARD #{card_number} : TEL 없음"
    )

    print(
        f"    FN : {fn!r}"
    )

    print(
        f"    N  : {n_value!r}"
    )

    print(
        "-" * 60
    )


# ============================================================
# VCF -> SQLite
# ============================================================

def import_vcf(
    vcf_path: str,
    db_path: str
):

    # --------------------------------------------------------
    # VCF 파일 읽기
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
            f"오류: 파일을 찾을 수 없습니다: "
            f"{vcf_path}",
            file=sys.stderr
        )

        sys.exit(1)


    except UnicodeDecodeError:

        print(
            f"오류: UTF-8로 읽을 수 없습니다: "
            f"{vcf_path}",
            file=sys.stderr
        )

        sys.exit(1)


    # --------------------------------------------------------
    # VCF 연락처 분리
    # --------------------------------------------------------

    cards = parse_vcards(text)


    # --------------------------------------------------------
    # SQLite 연결
    # --------------------------------------------------------

    conn = init_database(
        db_path
    )


    # --------------------------------------------------------
    # 현재 날짜/시간
    #
    # INSERT 시에만 사용
    # --------------------------------------------------------

    now = datetime.now().strftime(  # noqa: DTZ005
        "%Y-%m-%d %H:%M:%S"
    )


    # --------------------------------------------------------
    # 통계
    # --------------------------------------------------------

    inserted = 0

    updated = 0

    missing_fn = 0

    missing_tel = 0


    # --------------------------------------------------------
    # 연락처 처리
    # --------------------------------------------------------

    for card_number, card in enumerate(
        cards,
        start=1
    ):

        fn, n_value, phones = extract_contact(
            card
        )


        # ====================================================
        # FN 없음
        # ====================================================

        if not fn:

            missing_fn += 1

            print_missing_fn_warning(
                card_number,
                n_value,
                phones
            )

            # FN 없는 연락처는 저장하지 않는다.

            continue
        else:
            fn_front = fn.split()[0]  # FN의 첫 번째 단어를 fn_front로 사용


        # ====================================================
        # TEL 없음
        # ====================================================

        if not phones:

            missing_tel += 1

            print_missing_tel_warning(
                card_number,
                fn,
                n_value
            )

            # TEL 없는 연락처는 저장하지 않는다.

            continue


        # ====================================================
        # 전화번호 저장
        # ====================================================

        for tel_section, tel_number in phones:

            # ------------------------------------------------
            # 기존 Key 확인
            # ------------------------------------------------

            existing = conn.execute(
                """
                SELECT 1
                FROM contacts
                WHERE tel_section = ?
                  AND tel_number = ?
                """,
                (
                    tel_section,
                    tel_number
                )
            ).fetchone()


            # =================================================
            # 신규
            # =================================================

            if existing is None:

                conn.execute(
                    """
                    INSERT INTO contacts
                    (
                        tel_section,
                        tel_number,
                        fn_front,
                        fn_old,
                        fn_new,
                        insert_date,
                        update_date
                    )
                    VALUES
                    (
                        ?,
                        ?,
                        ?,
                        ?,
                        NULL,
                        ?,
                        NULL
                    )
                    """,
                    (
                        tel_section,
                        tel_number,
                        fn_front,
                        fn,
                        now
                    )
                )

                inserted += 1


            # =================================================
            # 기존
            # =================================================

            else:

                # ------------------------------------------------
                # 중요:
                #
                # fn_front, fn_old만 최신 VCF 기준으로 변경한다.
                #
                # fn_new       -> 보존
                # update_date  -> 보존
                # insert_date  -> 보존
                # ------------------------------------------------

                conn.execute(
                    """
                    UPDATE contacts
                    SET fn_front = ?,
                        fn_old = ?
                    WHERE tel_section = ?
                      AND tel_number = ?
                    """,
                    (
                        fn_front,
                        fn,
                        tel_section,
                        tel_number
                    )
                )

                updated += 1


    # --------------------------------------------------------
    # DB 저장
    # --------------------------------------------------------

    conn.commit()

    conn.close()


    # --------------------------------------------------------
    # 결과 출력
    # --------------------------------------------------------

    print()

    print(
        "=" * 60
    )

    print(
        "VCF -> SQLite 처리 완료"
    )

    print(
        "=" * 60
    )

    print(
        f"VCF 연락처 수 : {len(cards):,}"
    )

    print(
        f"신규 INSERT   : {inserted:,}"
    )

    print(
        f"기존 UPDATE   : {updated:,}"
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

    print(
        "=" * 60
    )


# ============================================================
# 사용법 출력
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
        "    python vcf_to_sqlite.py "
        "입력.vcf contacts.db"
    )


# ============================================================
# main
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