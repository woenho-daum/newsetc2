import getpass
import sys
# sqlcipher3 또는 pysqlcipher3 모두 호환
try:
    from sqlcipher3 import dbapi2 as sqlite3
except ImportError:
    from pysqlcipher3 import dbapi2 as sqlite3


def migrate_drivers_data():
    # -------------------------------------------------------------
    # 1. DB 연결 및 암호화 설정
    # -------------------------------------------------------------
    # 배차 DB (암호화가 안 된 일반 DB일 경우)
    conn_baecha = sqlite3.connect('baecha.db')

    # 설문 DB 암호 입력받기 (화면에 마스킹 처리됨)
    db_password = getpass.getpass("survey.db 암호를 입력하세요: ")
    if not db_password:
        print("❌ 암호가 입력되지 않았습니다. 프로그램을 종료합니다.")
        sys.exit(1)

    conn_servey = sqlite3.connect('survey.db')

    cursor_baecha = conn_baecha.cursor()
    cursor_servey = conn_servey.cursor()

    # survey.db 암호 설정 (PRAGMA key는 다른 쿼리 실행 전에 가장 먼저 실행해야 합니다)
    cursor_servey.execute(f"PRAGMA key = '{db_password}';")

    # 암호 정상 여부 및 연결 검증
    try:
        cursor_servey.execute("SELECT count(*) FROM sqlite_master;")
    except Exception:
        print("❌ survey.db 암호가 틀렸거나 데이터베이스 파일이 손상되었습니다.")
        conn_baecha.close()
        conn_servey.close()
        sys.exit(1)

    print("✅ survey.db 암호 인증 완료. 마이그레이션을 시작합니다.")

    # -------------------------------------------------------------
    # 2. 테이블 및 인덱스 생성
    # -------------------------------------------------------------
    cursor_servey.executescript('''
        DROP TABLE IF EXISTS drivers;

        CREATE TABLE IF NOT EXISTS drivers (
            name            TEXT PRIMARY KEY,
            phone           TEXT NOT NULL UNIQUE,
            shift_day       TEXT NOT NULL CHECK (shift_day IN ('월','화','수','목','금','토','일')),
            off_day         TEXT NOT NULL CHECK (off_day IN ('월','화','수','목','금','토','일')),
            childcare_day   TEXT CHECK (
                childcare_day IS NULL OR childcare_day IN ('월','화','수','목','금','토','일')
            ),
            childcare_start TEXT NULL CHECK (
                childcare_start IS NULL OR (
                    length(childcare_start) = 10 AND 
                    strftime('%Y-%m-%d', childcare_start) = childcare_start
                )
            ),
            childcare_end   TEXT NULL CHECK (
                childcare_end IS NULL OR (
                    length(childcare_end) = 10 AND 
                    strftime('%Y-%m-%d', childcare_end) = childcare_end
                )
            ),
            CHECK (shift_day <> off_day)
        );

        CREATE INDEX IF NOT EXISTS idx_drivers_off_day ON drivers(off_day);
        CREATE INDEX IF NOT EXISTS idx_drivers_shift_day ON drivers(shift_day);
        CREATE INDEX IF NOT EXISTS idx_drivers_childcare_day ON drivers(childcare_day);
    ''')

    # -------------------------------------------------------------
    # 3. SELECT 쿼리
    # -------------------------------------------------------------
    query = '''
        SELECT 
            d.driver_name AS name,
            c.tel_number AS phone,
            REPLACE(d.shift, '요일', '') AS shift_day,
            REPLACE(d.rest_day, '요일', '') AS off_day,
            c.tel_section
        FROM dispatch_settings d
        INNER JOIN contacts c 
            ON d.driver_name = c.fn_front 
           AND c.fn_order = '고정/쉬프트'
        WHERE d.driver_name IS NOT NULL 
          AND d.driver_name != ''
          AND d.shift != d.rest_day;
    '''

    cursor_baecha.execute(query)
    rows = cursor_baecha.fetchall()

    insert_sql = '''
        INSERT INTO drivers (name, phone, shift_day, off_day)
        VALUES (?, ?, ?, ?);
    '''

    update_sql = '''
        UPDATE drivers 
        SET phone = ?, shift_day = ?, off_day = ?
        WHERE name = ?;
    '''

    insert_cnt = 0
    update_cnt = 0
    skip_cnt = 0

    # -------------------------------------------------------------
    # 4. 명시적 중복 검증 후 처리
    # -------------------------------------------------------------
    for name, phone, shift_day, off_day, tel_section in rows:
        cursor_servey.execute(
            "SELECT 1 FROM drivers WHERE name = ?", (name,)
        )
        exists = cursor_servey.fetchone()

        if not exists:
            cursor_servey.execute(
                insert_sql, (name, phone, shift_day, off_day)
            )
            insert_cnt += 1
        else:
            if tel_section == 'work':
                cursor_servey.execute(
                    update_sql, (phone, shift_day, off_day, name)
                )
                update_cnt += 1
                print(
                    f"업데이트: {name} - 전화번호: {phone}, 근무일: {shift_day}, 휴무일: {off_day}"
                )
            else:
                skip_cnt += 1

    conn_servey.commit()

    print(
        f"\n처리 완료 - 신규 추가: {insert_cnt}건, 'work' 업데이트: {update_cnt}건, 무시됨: {skip_cnt}건"
    )

    conn_baecha.close()
    conn_servey.close()


if __name__ == '__main__':
    migrate_drivers_data()