import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright

URL = "https://oge.fipi.ru/bank/questions.php?proj=DE0E276E497AB3784C3FC4CC20248DC0&init_filter_themes=1"
TARGET_IDS = ["45B944", "426640", "FADB4B"]

OUT = Path("data/debug_fipi")
OUT.mkdir(parents=True, exist_ok=True)

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={"width": 1600, "height": 1200})
        print("[debug] open", URL)
        await page.goto(URL, wait_until="domcontentloaded", timeout=180_000)
        await page.wait_for_timeout(10_000)

        html = await page.content()
        (OUT / "page.html").write_text(html, encoding="utf-8")
        await page.screenshot(path=str(OUT / "page.png"), full_page=True)

        data = await page.evaluate(
            """(targetIds) => {
                function shortText(el) {
                    return (el.innerText || el.textContent || "")
                        .replace(/\\s+/g, " ")
                        .trim()
                        .slice(0, 3000);
                }

                function elemInfo(el) {
                    return {
                        tag: el.tagName,
                        id: el.id || "",
                        className: el.className ? String(el.className) : "",
                        text: shortText(el),
                        html: el.outerHTML.slice(0, 50000)
                    };
                }

                const result = {};
                for (const taskId of targetIds) {
                    const all = [...document.querySelectorAll("body *")];
                    const hits = all.filter(el => shortText(el).includes(taskId));

                    result[taskId] = hits.slice(0, 20).map(el => {
                        const ancestors = [];
                        let cur = el;
                        for (let i = 0; i < 8 && cur; i++) {
                            ancestors.push(elemInfo(cur));
                            cur = cur.parentElement;
                        }
                        return ancestors;
                    });
                }

                return result;
            }""",
            TARGET_IDS,
        )

        (OUT / "target_cards.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        print("[debug] saved to", OUT)
        await browser.close()

asyncio.run(main())
