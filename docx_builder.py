"""Генератор docx по образцу example_target.docx.

Альбомная A4, таблица 2 колонки, 1 карточка = 1 ячейка.
Структура карточки: «Номер: <qid>» жирным, текст с inline OMath-формулами,
«Ответ:» + PNG-плашка квадратиков.

Развёрнутые задачи (по 2 на лист, без квадратиков ответа) идут отдельным файлом.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree

from extractor import Task, Chunk, clean_chunks
from math_convert import fipi_mathml_to_omath

OMATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# Альбомная A4: 29.7 × 21.0 см.
# 2 колонки секции с gap → ~13.5 см на колонку → карточка 13 см.
A4_LANDSCAPE_WIDTH_CM = 29.7
A4_LANDSCAPE_HEIGHT_CM = 21.0
PAGE_MARGIN_CM = 1.0
CARD_WIDTH_CM = 13.0
SECTION_GAP_CM = 0.5

# Минимальная высота карточки. None = auto (по контенту), плотно.
# 9.0 = 4 на лист, 6.0 = 6 на лист.
CARD_HEIGHT_BY_PER_PAGE: dict[int, float | None] = {0: None, 4: 9.0, 6: 6.0}


ASSETS_DIR = Path(__file__).parent / "assets"


def _setup_landscape_a4(section, margin_cm: float = PAGE_MARGIN_CM) -> None:
    """Явно A4 (29.7 × 21.0 см) ландшафт. Не полагаемся на дефолт (там Letter)."""
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(A4_LANDSCAPE_WIDTH_CM)
    section.page_height = Cm(A4_LANDSCAPE_HEIGHT_CM)
    section.left_margin = Cm(margin_cm)
    section.right_margin = Cm(margin_cm)
    section.top_margin = Cm(margin_cm)
    section.bottom_margin = Cm(margin_cm)


def _set_section_columns(section, num: int = 2, gap_cm: float = SECTION_GAP_CM) -> None:
    """Переключить секцию в N-колоночный layout (карточки потекут в 2 колонки)."""
    sectPr = section._sectPr
    for old in sectPr.findall(qn("w:cols")):
        sectPr.remove(old)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), str(num))
    cols.set(qn("w:space"), str(int(gap_cm * 567)))  # twentieths of a point
    sectPr.append(cols)


def _set_table_fixed_width(table, width_cm: float) -> None:
    """Зафиксировать ширину таблицы — отключить autofit (чтобы карточки были одинаковые)."""
    width_dxa = int(width_cm * 567)
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(width_dxa))
    tblPr.append(tblW)
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    # И ячейки/колонки тоже
    for col in table.columns:
        col.width = Cm(width_cm)
    for row in table.rows:
        for cell in row.cells:
            cell.width = Cm(width_cm)


def _keep_table_together(table) -> None:
    """Запретить разрыв таблицы между страницами/колонками."""
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        for old in trPr.findall(qn("w:cantSplit")):
            trPr.remove(old)
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)


def _set_table_borders(table) -> None:
    """Явно прописать рамки со всех 4 сторон + внутренние (на всякий) —
    «Table Grid» style иногда не отрисовывается во всех ридерах."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblBorders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "8")        # 8 = 1 pt
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def _set_min_row_height(table, height_cm: float) -> None:
    """Задать минимальную высоту строки (карточка не меньше height_cm)."""
    for row in table.rows:
        row.height = Cm(height_cm)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def _add_text_or_math_to_paragraph(paragraph, ch: Chunk) -> None:
    """Добавить inline-чанк (текст или формула) в параграф."""
    if ch.kind == "text":
        if ch.value:
            paragraph.add_run(ch.value)
    elif ch.kind == "math":
        if not ch.mathml:
            return
        try:
            omath_xml = fipi_mathml_to_omath(ch.mathml)
            wrapped = f'<root xmlns:m="{OMATH_NS}">{omath_xml}</root>'
            root = etree.fromstring(wrapped.encode("utf-8"))
            for elem in root:
                paragraph._p.append(elem)
        except Exception as e:
            paragraph.add_run(" [formula?] ")
            print(f"[builder] math fail: {e}", file=sys.stderr)
    elif ch.kind == "break":
        paragraph.add_run().add_break()


def _add_chunks_with_images(cell, chunks: list[Chunk], max_img_cm: float) -> None:
    """Заполнить ячейку: inline-текст/формулы — в одном параграфе, картинки —
    в отдельных центрированных параграфах между ними. Картинки ограничены
    по ширине так, чтобы влезали в карточку."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    current = cell.add_paragraph()
    for ch in chunks:
        if ch.kind == "image":
            if not ch.image_bytes:
                continue
            # Закрыть текущий inline-параграф (даже если он пустой — пусть будет)
            p_img = cell.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_img.add_run()
            try:
                # Auto-ресайз: вписать в max_img_cm по ширине, сохраняя пропорции
                img_w_cm = _fit_image_width_cm(ch.image_bytes, max_img_cm)
                run.add_picture(io.BytesIO(ch.image_bytes), width=Cm(img_w_cm))
            except Exception as e:
                print(f"[builder] image fail: {e}", file=sys.stderr)
            # Открыть новый inline-параграф для последующего текста
            current = cell.add_paragraph()
        else:
            _add_text_or_math_to_paragraph(current, ch)


def _fit_image_width_cm(image_bytes: bytes, max_cm: float) -> float:
    """Вернуть ширину в см: min(нативная ширина / 100 DPI, max_cm)."""
    try:
        from PIL import Image as PILImage
        img = PILImage.open(io.BytesIO(image_bytes))
        # 96 DPI — стандарт docx для авторазмера
        native_cm = img.width / 96 * 2.54
        return min(native_cm, max_cm)
    except Exception:
        return max_cm


def _fill_card(cell, task: Task, with_answer_squares: bool = True,
               max_img_cm: float = 11.0) -> None:
    """Заполнить ячейку-карточку: Номер, тело задачи (текст+формулы+картинки),
    Ответ + квадратики."""
    # Удалить дефолтный пустой параграф (python-docx ставит его при создании ячейки)
    for p in list(cell.paragraphs):
        p._p.getparent().remove(p._p)

    # 1. Заголовок «Номер: <qid>»
    p_num = cell.add_paragraph()
    run = p_num.add_run(f"Номер: {task.qid}")
    run.bold = True
    run.font.size = Pt(10)

    # 2. Тело задачи — текст и формулы inline, картинки отдельными параграфами
    _add_chunks_with_images(cell, clean_chunks(task.content), max_img_cm)

    # 3. Ответ — слово на отдельной строке, квадратики — на следующей
    if with_answer_squares:
        p_lbl = cell.add_paragraph()
        run_lbl = p_lbl.add_run("Ответ:")
        run_lbl.bold = True
        run_lbl.font.size = Pt(11)

        p_sq = cell.add_paragraph()
        run_img = p_sq.add_run()
        try:
            run_img.add_picture(str(ASSETS_DIR / "answer_squares.png"), width=Cm(8))
        except Exception as e:
            print(f"[builder] answer-squares fail: {e}", file=sys.stderr)


def build_docx(
    tasks: list[Task],
    out_path: Path,
    with_answer_squares: bool = True,
    card_width_cm: float = CARD_WIDTH_CM,
    per_page: int = 0,
) -> None:
    """Собрать docx: каждая карточка = отдельная таблица 1×1.

    per_page: ориентир сколько карточек на лист — 4 или 6.
    Карточки имеют минимальную высоту (AT_LEAST) — могут вырасти под крупный
    контент, но не меньше базовой. Документ — альбомная A4 с 2-колоночной секцией.
    """
    doc = Document()
    section = doc.sections[0]
    _setup_landscape_a4(section)
    _set_section_columns(section, num=2)

    if not tasks:
        doc.add_paragraph("Пусто.")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_path))
        return

    card_height_cm = CARD_HEIGHT_BY_PER_PAGE.get(per_page, None)

    for task in tasks:
        table = doc.add_table(rows=1, cols=1)
        table.style = "Table Grid"
        table.autofit = False
        _set_table_fixed_width(table, card_width_cm)
        _set_table_borders(table)
        if card_height_cm:
            _set_min_row_height(table, card_height_cm)
        _keep_table_together(table)
        _fill_card(table.cell(0, 0), task,
                   with_answer_squares=with_answer_squares,
                   max_img_cm=card_width_cm - 2.0)
        # Разделитель — иначе python-docx может склеить соседние таблицы в одну
        doc.add_paragraph()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(
        f"[builder] {out_path} — {len(tasks)} карточек, "
        f"ширина {card_width_cm} см, высота {'авто' if not card_height_cm else f'{card_height_cm} см'}",
        file=sys.stderr,
    )


def group_tasks(tasks: list[Task]) -> list[Task]:
    """Группировка: сортировка по первым словам, развёрнутые в конец."""
    short_and_choice = [t for t in tasks if t.answer_type != "extended"]
    extended = [t for t in tasks if t.answer_type == "extended"]
    short_and_choice.sort(key=lambda t: t.first_words)
    extended.sort(key=lambda t: t.first_words)
    return short_and_choice + extended


def split_by_answer_type(tasks: list[Task]) -> tuple[list[Task], list[Task]]:
    """Вернуть (короткие+выбор, развёрнутые)."""
    short = [t for t in tasks if t.answer_type != "extended"]
    extended = [t for t in tasks if t.answer_type == "extended"]
    return short, extended
