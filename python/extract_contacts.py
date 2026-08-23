"""
구글 연락처 내보내기 CSV에서 이름(old_name)과 전화번호(phone)만 추출하여
새로운 CSV 파일로 저장하는 스크립트.

사용법:
    python extract_contacts.py 입력파일.csv [출력파일.csv]

동작:
- 이름은 'First Name' 컬럼을 사용하고, 비어있으면 'File As' 컬럼을 사용함.
- 전화번호는 'Phone 1 - Value' ~ 'Phone 4 - Value' 중 값이 있는 항목을
  모두 각각 한 행씩 출력함 (한 사람이 번호를 여러 개 가지고 있으면
  old_name이 같은 행이 여러 개 생성됨).
- 이름/전화번호가 모두 없는 행은 건너뜀.
"""

import csv
import sys


def extract_contacts(input_path: str, output_path: str) -> None:
    rows_out = []

    with open(input_path, encoding="utf-8-sig", newline="") as f_in:
        reader = csv.DictReader(f_in)

        for row in reader:
            name = (row.get("First Name") or "").strip()
            if not name:
                name = (row.get("File As") or "").strip()

            phones = []
            for i in range(1, 5):
                value = (row.get(f"Phone {i} - Value") or "").strip()
                if value:
                    phones.append(value)

            if not name or not phones:
                continue

            for phone in phones:
                rows_out.append({"old_name": name, "phone": phone})

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f_out:
        writer = csv.DictWriter(f_out, fieldnames=["old_name", "phone"])
        writer.writeheader()
        writer.writerows(rows_out)

    print(f"완료: {len(rows_out)}개 행을 '{output_path}' 에 저장했습니다.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python extract_contacts.py 입력파일.csv [출력파일.csv]")
        sys.exit(1)

    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "contacts_extracted.csv"

    extract_contacts(in_path, out_path)
