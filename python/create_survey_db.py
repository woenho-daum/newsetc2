import sqlite3

def migrate_drivers_data():
    conn_baecha = sqlite3.connect('baecha.db')
    conn_servey = sqlite3.connect('survey.db')
    
    cursor_baecha = conn_baecha.cursor()
    cursor_servey = conn_servey.cursor()

    # 1. 테이블 및 인덱스 생성
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

    # 2. SELECT 쿼리 (tel_section 값도 가져옴)
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

    # 3. 명시적 중복 검증 후 처리
    for name, phone, shift_day, off_day, tel_section in rows:
        # 1) 기존 데이터 존재 여부 명시적 확인
        cursor_servey.execute("SELECT 1 FROM drivers WHERE name = ?", (name,))
        exists = cursor_servey.fetchone()

        if not exists:
            # 존재하지 않으면 신규 INSERT
            cursor_servey.execute(insert_sql, (name, phone, shift_day, off_day))
            insert_cnt += 1
        else:
            # 이미 존재(중복)하면서 tel_section이 'work'일 때만 명시적 UPDATE
            if tel_section == 'work':
                cursor_servey.execute(update_sql, (phone, shift_day, off_day, name))
                update_cnt += 1
                print(f"업데이트: {name} - 전화번호: {phone}, 근무일: {shift_day}, 휴무일: {off_day}")
            else:
                skip_cnt += 1

    conn_servey.commit()

    print(f"처리 완료 - 신규 추가: {insert_cnt}건, 'work' 업데이트: {update_cnt}건, 무시됨: {skip_cnt}건")

    conn_baecha.close()
    conn_servey.close()

if __name__ == '__main__':
    migrate_drivers_data()