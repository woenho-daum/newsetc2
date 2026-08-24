import asyncio
import json
from playwright.async_api import async_playwright


CDP_URL = "http://127.0.0.1:9222"


async def get_dispatch_settings():

    async with async_playwright() as p:

        # Chrome CDP 연결
        browser = await p.chromium.connect_over_cdp(CDP_URL)

        # 모든 탭 확인
        pages = []

        for context in browser.contexts:
            pages.extend(context.pages)

        target_page = None

        # '고정(쉬프트)지정' 페이지 찾기
        for page in pages:
            try:
                title = await page.title()

                if "고정(쉬프트)지정" in title:
                    target_page = page
                    break

            except Exception:
                pass

        if target_page is None:
            raise RuntimeError(
                "'고정(쉬프트)지정' 페이지를 찾지 못했습니다."
            )

        print("페이지 :", await target_page.title())
        print("URL    :", target_page.url)

        # 테이블이 존재할 때까지 대기
        await target_page.locator(
            "#dispatch_settings_table"
        ).wait_for(
            state="attached",
            timeout=10000
        )

        # --------------------------------------------------
        # Chrome 안에서 table 전체를 한 번에 추출
        # --------------------------------------------------

        result = await target_page.locator(
            "#dispatch_settings_table"
        ).evaluate("""
        table => {

            // -----------------------------
            // 컬럼
            // -----------------------------

            const columns = Array.from(
                table.querySelectorAll("thead th")
            ).map((th, index) => ({
                index: index,
                name: th.innerText.trim()
            }));


            // -----------------------------
            // ROW
            // -----------------------------

            const rows = Array.from(
                table.querySelectorAll("tbody tr")
            ).map((tr, rowIndex) => {

                const cells = Array.from(
                    tr.querySelectorAll(":scope > td")
                );

                const values = cells.map(
                    td => td.innerText.trim()
                );

                // data-* 속성
                const attributes = {};

                for (const attr of tr.attributes) {

                    if (attr.name.startsWith("data-")) {
                        attributes[attr.name] = attr.value;
                    }
                }

                return {
                    index: rowIndex,
                    attributes: attributes,
                    values: values
                };
            });


            // -----------------------------
            // 컬럼명 → 값 형태의 record
            // -----------------------------

            const records = rows.map(row => {

                const record = {};

                columns.forEach((column, index) => {

                    record[column.name] =
                        row.values[index] ?? "";

                });

                // data-* 속성도 추가
                Object.assign(
                    record,
                    row.attributes
                );

                return record;
            });


            return {
                columns: columns,
                rows: rows,
                records: records
            };
        }
        """)

        return result


async def main():

    data = await get_dispatch_settings()

    # -----------------------------------------
    # 컬럼
    # -----------------------------------------

    print("\n===== COLUMN =====")

    for column in data["columns"]:
        print(
            column["index"],
            column["name"]
        )


    # -----------------------------------------
    # ROW
    # -----------------------------------------

    print("\n===== ROW =====")

    for row in data["rows"]:

        print(
            row["index"],
            row["values"]
        )


    # -----------------------------------------
    # RECORD
    # -----------------------------------------

    print("\n===== RECORD =====")

    for record in data["records"]:

        print(record)


    # -----------------------------------------
    # JSON 저장
    # -----------------------------------------

    with open(
        "dispatch_settings.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )


if __name__ == "__main__":
    asyncio.run(main())
