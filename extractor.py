"""Playwright-парсер: HTML qblock → структурированный Task.

Каждый Task — это упорядоченный поток inline-элементов:
  text (str), math (MathML), image (PNG bytes), break (\n)
Плюс метаданные: qid, guid, тип задачи (short/choice/extended).
"""
from __future__ import annotations

import asyncio
import base64
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from playwright.async_api import async_playwright


FIPI_URL_TEMPLATE = "https://{host}/bank/questions.php?proj={proj}&init_filter_themes=1"

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
VIEWPORT = {"width": 1400, "height": 2400}


@dataclass
class Chunk:
    kind: Literal["text", "math", "image", "break"]
    value: str = ""
    mathml: str = ""
    image_bytes: bytes = b""
    image_ext: str = "png"


@dataclass
class Task:
    qid: str
    guid: str
    answer_type: Literal["short", "choice", "extended"]
    content: list[Chunk] = field(default_factory=list)
    options: list[list[Chunk]] = field(default_factory=list)  # для choice — варианты

    @property
    def first_words(self) -> str:
        """Первые 4 слова текстовой части — для группировки."""
        text = " ".join(c.value for c in self.content if c.kind == "text").strip()
        return " ".join(text.split()[:4]).lower()


# JS-код для извлечения всех qblock'ов с MathML формулами
EXTRACT_JS = r"""
() => {
    const allMjx = [...document.querySelectorAll('mjx-container')];
    let formulas = [];
    try {
        formulas = MathJax.startup.document.math.toArray
            ? MathJax.startup.document.math.toArray().map(m => m.math)
            : [...MathJax.startup.document.math].map(m => m.math);
    } catch (e) {}
    const formulaMap = new Map();
    allMjx.forEach((el, i) => formulaMap.set(el, formulas[i] || ''));

    function walk(node, out) {
        if (node.nodeType === 3) {  // TEXT_NODE
            const t = node.textContent;
            if (t) out.push({kind: 'text', value: t});
            return;
        }
        if (node.nodeType !== 1) return;  // не ELEMENT_NODE
        const tag = node.tagName;
        if (tag === 'MJX-CONTAINER') {
            out.push({kind: 'math', mathml: formulaMap.get(node) || ''});
        } else if (tag === 'IMG') {
            out.push({kind: 'image', src: node.src, alt: node.alt || ''});
        } else if (tag === 'BR') {
            out.push({kind: 'break'});
        } else if (tag === 'SCRIPT' || tag === 'STYLE') {
            return;
        } else {
            for (const child of node.childNodes) walk(child, out);
        }
    }

    function extractCell(block) {
        // Основной контент задачи — в первой ячейке td с классом cell_0
        return block.querySelector('table td.cell_0') || block.querySelector('td.cell_0');
    }

    const blocks = [...document.querySelectorAll('.qblock')];
    return blocks.map(block => {
        const qid = (block.id || '').replace(/^q/, '');
        const guid = block.querySelector('input[name="guid"]')?.value || '';
        const cell = extractCell(block);

        const content = [];
        if (cell) {
            // Вытаскиваем содержимое всех ячеек .cell_0 (там может быть несколько — варианты)
            const cells = [...block.querySelectorAll('td.cell_0')];
            cells.forEach((c, idx) => {
                walk(c, content);
                if (idx < cells.length - 1) content.push({kind: 'break'});
            });
        }

        // Тип задачи
        const hasRadio = !!block.querySelector('input[type="radio"]');
        const blockText = block.innerText || '';
        const isExtended = /развёрн|подроб/i.test(blockText);
        let answerType = 'short';
        if (hasRadio) answerType = 'choice';
        else if (isExtended) answerType = 'extended';

        return {qid, guid, answerType, content};
    });
}
"""


async def fetch_image(page, src: str) -> tuple[bytes, str]:
    """Скачать картинку через тот же контекст браузера (cookies/auth)."""
    try:
        resp = await page.context.request.get(src)
        ext = "png"
        ct = resp.headers.get("content-type", "")
        if "jpeg" in ct or "jpg" in ct:
            ext = "jpg"
        elif "gif" in ct:
            ext = "gif"
        elif "svg" in ct:
            ext = "svg"
        return await resp.body(), ext
    except Exception as e:
        print(f"[extractor] image fetch failed {src}: {e}", file=sys.stderr)
        return b"", "png"


async def _load_page(page, url: str) -> list[dict]:
    """Перейти на url и вернуть raw-список qblock'ов."""
    await page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_selector(".qblock", timeout=30_000)
    await page.wait_for_load_state("networkidle", timeout=20_000)
    await asyncio.sleep(3)  # MathJax
    return await page.evaluate(EXTRACT_JS)


async def _click_next_if_any(page) -> bool:
    """Попробовать перейти на следующую страницу пагинации.
    Возвращает True если получилось.
    """
    candidates = [
        ".pagination a.next",
        "a.pagination-next",
        "img[onclick*='nextPage']",
        "[onclick*='nextPage']",
        "[onclick*='changePage']",
    ]
    for sel in candidates:
        loc = page.locator(sel).first
        if await loc.count():
            try:
                await loc.click(timeout=3_000)
                await page.wait_for_load_state("networkidle", timeout=15_000)
                await asyncio.sleep(2)
                return True
            except Exception:
                continue
    return False


async def extract_tasks(
    proj: str,
    host: str = "oge.fipi.ru",
    limit: int = 0,
    max_pages: int = 30,
) -> list[Task]:
    """Открыть банк, перебрать страницы пагинации, собрать Task'и до limit."""
    base_url = FIPI_URL_TEMPLATE.format(host=host, proj=proj)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            viewport=VIEWPORT,
            user_agent=USER_AGENT,
            locale="ru-RU",
            ignore_https_errors=True,
        )
        page = await ctx.new_page()
        try:
            tasks: list[Task] = []
            seen_guids: set[str] = set()
            page_num = 1

            while page_num <= max_pages:
                # 1-й проход — base_url; далее пробуем pagenum= параметр + клик по next
                if page_num == 1:
                    url = base_url
                else:
                    url = f"{base_url}&page={page_num}"
                print(f"[extract] page {page_num}: open {url}", file=sys.stderr)
                raw = await _load_page(page, url)

                new_on_page = 0
                for r in raw:
                    if r["guid"] in seen_guids:
                        continue
                    seen_guids.add(r["guid"])
                    new_on_page += 1
                    if limit and len(tasks) >= limit:
                        break
                    chunks: list[Chunk] = []
                    for c in r["content"]:
                        if c["kind"] == "text":
                            chunks.append(Chunk(kind="text", value=c["value"]))
                        elif c["kind"] == "math":
                            chunks.append(Chunk(kind="math", mathml=c["mathml"]))
                        elif c["kind"] == "break":
                            chunks.append(Chunk(kind="break"))
                        elif c["kind"] == "image":
                            data, ext = await fetch_image(page, c["src"])
                            if data:
                                chunks.append(Chunk(kind="image", image_bytes=data, image_ext=ext))
                    tasks.append(Task(
                        qid=r["qid"],
                        guid=r["guid"],
                        answer_type=r["answerType"],
                        content=chunks,
                    ))

                print(f"[extract] page {page_num}: новых {new_on_page}, всего {len(tasks)}",
                      file=sys.stderr)

                if limit and len(tasks) >= limit:
                    break
                if new_on_page == 0:
                    # ни одной новой — либо pagenum не работает, либо банк кончился
                    if page_num == 1:
                        page_num += 1
                        continue
                    break
                page_num += 1

            return tasks
        finally:
            await browser.close()


import re as _re


def clean_chunks(chunks: list[Chunk]) -> list[Chunk]:
    """Убрать шапку 'Впишите правильный ответ.' и подобный мусор.
    Схлопнуть internal whitespace, но СОХРАНИТЬ пробелы на краях text-чанков —
    иначе слипнутся со следующей формулой/чанком.
    """
    SKIP_PATTERNS = [
        "Впишите правильный ответ.",
        "Впишите правильный ответ",
        "Выберите правильный ответ.",
        "Выберите правильный ответ",
        "Дайте развёрнутый ответ.",
        "Дайте развёрнутый ответ",
        "Дайте развернутый ответ.",
        "Дайте развернутый ответ",
    ]
    cleaned: list[Chunk] = []
    for c in chunks:
        if c.kind == "text":
            v = c.value
            for pat in SKIP_PATTERNS:
                v = v.replace(pat, "")
            # Схлопнуть множественные whitespace, но не trim'ить края
            v = _re.sub(r"\s+", " ", v)
            if v and v != " ":
                cleaned.append(Chunk(kind="text", value=v))
        else:
            cleaned.append(c)

    # Убрать ведущие break'ы и чисто-пробельные тексты в начале
    while cleaned and (cleaned[0].kind == "break" or
                       (cleaned[0].kind == "text" and not cleaned[0].value.strip())):
        cleaned.pop(0)
    return cleaned
