import sys


def sort_vcf_by_fn(input_vcf, output_vcf):
    # VCard 단위로 읽기
    with open(input_vcf, "r", encoding="utf-8", newline="") as f:
        text = f.read()

    cards = text.split("BEGIN:VCARD")
    cards = [card for card in cards if card.strip()]

    vcards = []

    for card in cards:
        card = "BEGIN:VCARD" + card

        # FN 값 추출
        fn = ""
        for line in card.splitlines():
            if line.startswith("FN;") or line.startswith("FN:"):
                fn = line.split(":", 1)[1]
                break

        vcards.append((fn, card))

    # FN 기준 정렬
    vcards.sort(key=lambda x: x[0])

    # 저장
    with open(output_vcf, "w", encoding="utf-8", newline="") as f:
        for fn, card in vcards:
            f.write(card.rstrip("\r\n") + "\r\n")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("사용법:")
        print("python sort_vcf.py 입력.vcf 출력.vcf")
        sys.exit(1)

    input_vcf = sys.argv[1]
    output_vcf = sys.argv[2]

    sort_vcf_by_fn(input_vcf, output_vcf)

    print(f"정렬 완료: {output_vcf}")
