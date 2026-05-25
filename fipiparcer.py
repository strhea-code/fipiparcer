"""ФИПИ-парсер: банк заданий → редактируемый docx с карточками.

Карточка = ячейка таблицы 2 колонки на альбомной A4 (4 карточки на лист).
Формулы переводятся в OMath (Office Math) — редактируются в Word,
не картинки. Развёрнутые задачи кладутся в отдельный docx (по 2 на лист).

Команды:
    sample-docx  — спарсить N задач и собрать docx (короткие + развёрнутые отдельно)

Примеры:
    python fipiparcer.py sample-docx                # ОГЭ матем, 20 задач
    python fipiparcer.py sample-docx --n 50
    python fipiparcer.py sample-docx --host ege.fipi.ru --proj AC437B34557F88EA4115D2F374B0A07B --n 20

ВАЖНО: fipi.ru закрыт для не-российских IP. Запускать с RU-VPN.
"""
from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

from extractor import extract_tasks
from docx_builder import build_docx, group_tasks, split_by_answer_type, split_large_tasks


PROJ_OGE_MATH = "DE0E276E497AB3784C3FC4CC20248DC0"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sample-docx", help="спарсить и собрать docx")
    s.add_argument("--host", default="oge.fipi.ru",
                   help="ФИПИ-хост: oge.fipi.ru (по умолчанию) или ege.fipi.ru")
    s.add_argument("--proj", default=PROJ_OGE_MATH,
                   help="proj GUID банка (по умолчанию — ОГЭ математика)")
    s.add_argument("--n", type=int, default=20, help="Сколько задач взять")
    s.add_argument("--per-page", type=int, default=0, choices=[0, 4, 6],
                   help="0 = высота карточки по контенту (плотно, по умолчанию); "
                        "4 = минимум 9 см (4 на лист); 6 = минимум 6 см (6 на лист)")
    s.add_argument("--out-dir", type=Path, default=Path("data/output"))
    s.add_argument("--out-name", default="oge_math_sample",
                   help="префикс имён файлов (.docx и -extended.docx)")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.cmd != "sample-docx":
        return 1

    print(f"[main] host={args.host} proj={args.proj} n={args.n}", file=sys.stderr)
    tasks = asyncio.run(extract_tasks(args.proj, host=args.host, limit=args.n))
    if not tasks:
        print("[main] нет задач", file=sys.stderr)
        return 1

    tasks = group_tasks(tasks)
    short, extended = split_by_answer_type(tasks)
    regular, large = split_large_tasks(short)

    args.out_dir.mkdir(parents=True, exist_ok=True)
    short_path = args.out_dir / f"{args.out_name}.docx"
    large_path = args.out_dir / f"{args.out_name}-large.docx"
    ext_path = args.out_dir / f"{args.out_name}-extended.docx"

    if regular:
        build_docx(regular, short_path, with_answer_squares=True, per_page=args.per_page)
    if large:
        # Крупные карточки не кладём в общий файл на 6 карточек: они ломают
        # печатную раскладку. В отдельном файле даём им больше места.
        build_docx(large, large_path, with_answer_squares=True, per_page=4)
    if extended:
        build_docx(extended, ext_path, with_answer_squares=False, per_page=args.per_page)

    print(
        f"[main] готово: {len(regular)} коротких, {len(large)} крупных, "
        f"{len(extended)} развёрнутых",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
