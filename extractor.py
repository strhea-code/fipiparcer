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
    let captureSeq = 0;

    function isVisible(el) {
        if (!el || !el.isConnected) return false;
        const st = window.getComputedStyle(el);
        if (!st || st.display === 'none' || st.visibility === 'hidden' || st.opacity === '0') return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
    }

    function mathMLFromMjx(node) {
        // Лучший вариант: настоящий MathML, который Word можно конвертировать в OMath.
        const mml = node.querySelector('mjx-assistive-mml math');
        if (mml) return mml.outerHTML;
        return '';
    }

    function normalizeMathText(text) {
        if (!text) return '';
        return String(text)
            .normalize('NFKC')
            .replace(/[\u2061\u2062\u2063\u2064]/g, '')  // невидимые мат. операторы
            .replace(/[\u2009\u200A\u200B\u202F\u00A0]/g, ' ')
            .replace(/\s+/g, ' ')
            .replace(/\s+([,.;:])/g, '$1')
            .trim();
    }


    function polishFormulaText(text) {
        if (!text) return '';
        return normalizeMathText(text)
            .replace(/([A-Za-zА-Яа-я])2\b/g, '$1²')
            .replace(/\s*([=<>+−\-])\s*/g, ' $1 ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function toSuperscript(text) {
        const map = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','+':'⁺','-':'⁻','=':'⁼','(':'⁽',')':'⁾'};
        const t = normalizeMathText(text);
        if (!t) return '';
        if (/^[0-9+\-=()]+$/.test(t)) return [...t].map(ch => map[ch] || ch).join('');
        return '^(' + t + ')';
    }

    function mjxText(el) {
        if (!el) return '';
        if (el.nodeType === 3) return normalizeMathText(el.textContent || '');
        if (el.nodeType !== 1) return '';

        const tag = el.tagName;
        if (tag === 'MJX-ASSISTIVE-MML') return '';
        if (tag === 'MJX-C') {
            const raw = el.textContent || '';
            // MathJax использует эти символы как невидимое умножение/тонкие пробелы.
            if (/^[\u2061\u2062\u2063\u2064\u2009\u200A\u200B\u202F\u00A0\s]+$/.test(raw)) return '';
            return normalizeMathText(raw);
        }
        if (tag === 'MJX-MO') {
            const t = normalizeMathText(el.textContent || '');
            if (!t || t === '⁢') return '';
            return t;
        }
        if (tag === 'MJX-MTEXT') {
            const t = normalizeMathText(el.textContent || '');
            // Часто это технический тонкий пробел, а не часть формулы.
            if (!t) return '';
            return t;
        }
        if (tag === 'MJX-MSUP') {
            const children = [...el.children].filter(ch => ch.tagName !== 'MJX-ASSISTIVE-MML');
            const base = children.find(ch => ch.tagName !== 'MJX-SCRIPT') || children[0];
            const script = children.find(ch => ch.tagName === 'MJX-SCRIPT') || children[1];
            return mjxText(base) + toSuperscript(mjxText(script));
        }
        if (tag === 'MJX-MSUB') {
            const children = [...el.children].filter(ch => ch.tagName !== 'MJX-ASSISTIVE-MML');
            const base = children.find(ch => ch.tagName !== 'MJX-SCRIPT') || children[0];
            const script = children.find(ch => ch.tagName === 'MJX-SCRIPT') || children[1];
            const sub = normalizeMathText(mjxText(script));
            return mjxText(base) + (sub ? '_' + sub : '');
        }
        if (tag === 'MJX-MFRAC') {
            const kids = [...el.children];
            const num = mjxText(kids[0]);
            const den = mjxText(kids[1]);
            return den ? '(' + num + ')/(' + den + ')' : num;
        }
        if (tag === 'MJX-BREAK') return ' ';

        let parts = [];
        for (const ch of el.childNodes) {
            const t = mjxText(ch);
            if (t) parts.push(t);
        }
        let joined = parts.join(' ');
        joined = joined
            .replace(/\s*([=<>+−\-])\s*/g, ' $1 ')
            .replace(/\s+/g, ' ')
            .replace(/([A-Za-zА-Яа-я])\s+([A-Za-zА-Яа-я])/g, '$1$2')
            .replace(/\s+([²³¹⁰⁴⁵⁶⁷⁸⁹])/g, '$1')
            .trim();
        return joined;
    }

    function visibleMathText(node) {
        // На ФИПИ MathJax часто отдает CHTML без assistive MathML.
        // innerText копируется правильно, но в Word символы 𝑎, 𝑥, ⁢ могут
        // отображаться пустыми местами. Поэтому восстанавливаем формулу из
        // mjx-тегов и сразу приводим её к обычному редактируемому тексту.
        const structured = mjxText(node);
        if (structured) return polishFormulaText(structured);

        const txt = normalizeMathText(node.innerText || node.textContent || '');
        if (txt) return polishFormulaText(txt);

        const speech = node.getAttribute('data-semantic-speech-none') || '';
        return normalizeMathText(speech)
            .replace(/ equals /g, ' = ')
            .replace(/ plus /g, ' + ')
            .replace(/ less than /g, ' < ')
            .replace(/ greater than /g, ' > ')
            .replace(/ squared/g, '²')
            .replace(/ empty/g, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function backgroundImageUrl(el) {
        try {
            const bg = window.getComputedStyle(el).backgroundImage || '';
            const m = bg.match(/url\(["']?(.+?)["']?\)/);
            if (m && m[1] && m[1] !== 'none') return new URL(m[1], document.baseURI).href;
        } catch (e) {}
        return '';
    }

    function shouldCaptureElement(el) {
        const tag = el.tagName;
        if (tag === 'SVG' || tag === 'CANVAS') return true;
        if (tag === 'OBJECT' || tag === 'EMBED') return true;
        // Иногда графики/рисунки лежат в блочных контейнерах без IMG, но с SVG/CANVAS внутри.
        if (el.querySelector && (el.querySelector('svg') || el.querySelector('canvas'))) return true;
        return false;
    }

    function pushText(out, text) {
        if (!text) return;
        const value = normalizeMathText(text);
        if (!value) return;
        out.push({kind: 'text', value: value});
    }

    function walk(node, out) {
        if (node.nodeType === 3) {  // TEXT_NODE
            pushText(out, node.textContent);
            return;
        }
        if (node.nodeType !== 1) return;  // не ELEMENT_NODE

        const tag = node.tagName;
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') return;
        if (tag === 'INPUT' || tag === 'BUTTON' || tag === 'SELECT' || tag === 'TEXTAREA') return;
        // Не выкидываем .answer/.qanswer: в ФИПИ в этих блоках часто лежат
        // варианты 1), 2), 3), а не готовый ответ. Иначе теряются условия
        // задач вроде 45B944 и 426640. Но служебные блоки ответа/проверки
        // всё равно не должны попадать в карточку.
        if (node.classList && (
            node.classList.contains('q_footer') ||
            node.classList.contains('solution') ||
            node.classList.contains('submit-block') ||
            node.classList.contains('answer-table') ||
            node.classList.contains('check-answer') ||
            node.classList.contains('answer-table-wrapper')
        )) return;
        if (!isVisible(node)) return;

        if (tag === 'MJX-CONTAINER') {
            const mml = mathMLFromMjx(node);
            if (mml) out.push({kind: 'math', mathml: mml});
            else pushText(out, ' ' + visibleMathText(node) + ' ');
            return;
        }
        if (tag === 'IMG') {
            const rawSrc = node.currentSrc || node.src || node.getAttribute('data-src') || node.getAttribute('src') || '';
            const src = rawSrc ? new URL(rawSrc, document.baseURI).href : '';
            out.push({kind: 'image', src: src, alt: node.alt || ''});
            return;
        }
        if (shouldCaptureElement(node)) {
            const id = 'fipi_capture_' + (++captureSeq);
            node.setAttribute('data-fipi-capture-id', id);
            out.push({kind: 'capture', captureId: id});
            return;
        }
        const bgUrl = backgroundImageUrl(node);
        if (bgUrl) {
            out.push({kind: 'image', src: bgUrl, alt: ''});
        }
        if (tag === 'BR') {
            out.push({kind: 'break'});
            return;
        }

        // Блочные элементы разделяем переносами, чтобы варианты 1), 2), 3) не слипались.
        const blockTags = new Set(['P', 'DIV', 'TR', 'TABLE', 'UL', 'OL']);
        if (blockTags.has(tag) && out.length) out.push({kind: 'break'});

        for (const child of node.childNodes) walk(child, out);

        if (blockTags.has(tag)) out.push({kind: 'break'});
    }

    function extractContentNodes(block) {
        // Самое надёжное на текущем ФИПИ — брать основной form checkform<ID> целиком.
        // Внутри него лежат и условие, и варианты 1), 2), 3), и таблицы с графиками.
        // Старый вариант брал только td.cell_0, поэтому у 45B944/426640 терялись
        // варианты, а у FADB4B часть таблицы с коэффициентами/графиками.
        const mainForm = block.querySelector('form[id^="checkform"]');
        if (mainForm && isVisible(mainForm)) return [mainForm];

        let nodes = [...block.querySelectorAll('.qtext, .question, .task, .content')].filter(isVisible);
        if (!nodes.length) nodes = [block];
        return nodes;
    }

    const blocks = [...document.querySelectorAll('.qblock')];
    return blocks.map(block => {
        const qid = (block.id || '').replace(/^q/, '');
        const guid = block.querySelector('input[name="guid"]')?.value || '';
        const content = [];

        const nodes = extractContentNodes(block);
        nodes.forEach((node, idx) => {
            walk(node, content);
            if (idx < nodes.length - 1) content.push({kind: 'break'});
        });

        const hasRadio = !!block.querySelector('input[type="radio"]');
        const blockText = block.innerText || '';
        const isExtended = /развёрн|разверн|подроб/i.test(blockText);
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




async def capture_element_image(page, capture_id: str) -> tuple[bytes, str]:
    """Сделать PNG-скриншот inline SVG/CANVAS/OBJECT, которые не являются <img>."""
    if not capture_id:
        return b"", "png"
    try:
        loc = page.locator(f'[data-fipi-capture-id="{capture_id}"]').first
        if await loc.count() == 0:
            return b"", "png"
        data = await loc.screenshot(type="png", timeout=10_000)
        return data, "png"
    except Exception as e:
        print(f"[extractor] capture failed {capture_id}: {e}", file=sys.stderr)
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
                            data, ext = await fetch_image(page, c.get("src", ""))
                            if data:
                                chunks.append(Chunk(kind="image", image_bytes=data, image_ext=ext))
                        elif c["kind"] == "capture":
                            data, ext = await capture_element_image(page, c.get("captureId", ""))
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
            # Убираем служебную инструкцию ФИПИ в конце условия; поле ответа
            # мы рисуем сами в карточке.
            v = _re.sub(r"В ответ запишите[^.?!]*(?:[.?!]|$)", "", v, flags=_re.IGNORECASE)
            v = _re.sub(r"В таблице под каждой буквой укажите[^.?!]*(?:[.?!]|$)", "", v, flags=_re.IGNORECASE)
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
