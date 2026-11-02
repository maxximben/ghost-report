#!/usr/bin/env python3
"""Сборка отчёта (ЛР, ДЗ, РГР) в .docx по правилам оформления.

    python build_report.py report.md                  # -> report.docx рядом с report.md
    python build_report.py report.md -o out/ЛР3.docx --pdf --png out/pages

Что делает:
  1. разбирает report.md (формат: references/markdown-format.md) и sources.json (источники);
  2. собирает .docx со всеми стилями/полями/нумерацией по правилам;
  3. если на машине есть Microsoft Word – рассчитывает вёрстку (номера страниц в содержании,
     «Продолжение таблицы N») и пересобирает документ до устойчивого результата;
  4. прогоняет проверку оформления (check_report.py) и печатает замечания.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gostlib import layout as LY            # noqa: E402
from gostlib import mdparse                 # noqa: E402
from gostlib.docxbuild import Builder       # noqa: E402


def say(*a):
    print(*a, flush=True)


def load_sources(path: str | None) -> list[dict]:
    if not path or not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8-sig") as f:
        data = json.load(f)
    if isinstance(data, dict):
        data = data.get("sources", [])
    return data


def build_once(meta, blocks, sources, base_dir, out, layout, breaks):
    b = Builder(meta, blocks, sources, base_dir, layout=layout, table_breaks=breaks)
    return b.build(out)


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Markdown -> отчёт .docx по правилам оформления")
    ap.add_argument("md", help="исходный файл report.md")
    ap.add_argument("-o", "--out", help="куда сохранить .docx (по умолчанию рядом с md)")
    ap.add_argument("--sources", help="файл с источниками (по умолчанию sources.json рядом с md или поле sources: в front matter)")
    ap.add_argument("--pdf", nargs="?", const="auto", help="дополнительно сохранить PDF (через Word)")
    ap.add_argument("--png", metavar="DIR", help="сохранить страницы PDF как PNG в DIR (для визуальной проверки)")
    ap.add_argument("--no-word", action="store_true", help="не использовать Word для расчёта вёрстки")
    ap.add_argument("--no-check", action="store_true", help="не запускать проверку оформления после сборки")
    args = ap.parse_args(argv)

    md_path = os.path.abspath(args.md)
    base_dir = os.path.dirname(md_path)
    out = os.path.abspath(args.out) if args.out else os.path.splitext(md_path)[0] + ".docx"
    os.makedirs(os.path.dirname(out), exist_ok=True)

    try:
        with open(md_path, encoding="utf-8-sig") as f:
            meta, blocks = mdparse.parse(f.read())
    except mdparse.MdError as e:
        say(f"ОШИБКА РАЗБОРА: {e}")
        return 1

    src_path = args.sources or (os.path.join(base_dir, str(meta["sources"])) if meta.get("sources") else
                                os.path.join(base_dir, "sources.json"))
    try:
        sources = load_sources(src_path)
    except (ValueError, OSError) as e:
        say(f"ОШИБКА: не удалось прочитать источники {src_path}: {e}")
        return 1

    res = build_once(meta, blocks, sources, base_dir, out, None, {})
    if res.errors:
        say("ОШИБКИ СБОРКИ (документ не создан):")
        for e in res.errors:
            say("  ✗", e)
        for w in res.warnings:
            say("  !", w)
        return 1

    use_word = (not args.no_word) and LY.word_available()
    rounds = 0
    layout, breaks = None, {}
    if use_word:
        say("Расчёт вёрстки в Microsoft Word…")
        prev = None
        try:
            while rounds < 6:
                rounds += 1
                info = LY.probe(out)
                new_breaks = LY.table_breaks_from_probe(info)
                for tid, rows in new_breaks.items():
                    breaks.setdefault(tid, set()).update(rows)
                cur = (info["pages"], tuple(sorted(info["bookmarks"].items())),
                       tuple(sorted((k, tuple(sorted(v))) for k, v in breaks.items())))
                if cur == prev:
                    break
                prev = cur
                layout = {"pages": info["pages"], "bookmarks": info["bookmarks"]}
                res = build_once(meta, blocks, sources, base_dir, out, layout, {k: sorted(v) for k, v in breaks.items()})
                if res.errors:
                    break
            else:
                res.warnings.append("вёрстка не стабилизировалась за 6 проходов: проверьте номера страниц в содержании вручную (F9)")
        except RuntimeError as e:
            use_word = False
            res = build_once(meta, blocks, sources, base_dir, out, None, {})
            res.warnings.append(f"расчёт вёрстки через Word не удался ({e}); номера страниц в содержании обновятся при открытии")
    else:
        why = "отключён флагом --no-word" if args.no_word else "Microsoft Word не найден"
        res.warnings.append(f"расчёт вёрстки ({why}): номера страниц в содержании обновятся при открытии документа (F9), "
                            "переносы таблиц на новую страницу («Продолжение таблицы») не расставлены")

    say(f"\nГотово: {out}")
    s = res.stats
    pages = layout["pages"] if layout else "?"
    say(f"  страниц: {pages}; рисунков: {s['figures']}; таблиц: {s['tables']}; формул: {s['equations']}; "
        f"источников: {s['sources']}; приложений: {s['appendices']}")
    for n in res.notes:
        say("  ·", n)
    for w in res.warnings:
        say("  !", w)

    pdf_path = None
    if args.pdf or args.png:
        if not (use_word and LY.word_available()):
            say("  ! PDF не создан: требуется Microsoft Word")
        else:
            pdf_path = os.path.splitext(out)[0] + ".pdf" if args.pdf in (None, "auto") else os.path.abspath(args.pdf)
            LY.probe(out, pdf=pdf_path)
            say(f"  PDF: {pdf_path}")
            if args.png:
                try:
                    try:
                        import pymupdf as fitz
                    except ImportError:
                        import fitz  # PyMuPDF (старое имя модуля)
                    os.makedirs(args.png, exist_ok=True)
                    d = fitz.open(pdf_path)
                    for i, page in enumerate(d, 1):
                        page.get_pixmap(dpi=70).save(os.path.join(args.png, f"page-{i:02d}.png"))
                    say(f"  PNG страниц: {args.png} ({len(d)} шт.)")
                except ImportError:
                    say("  ! PNG не созданы: установите PyMuPDF (pip install pymupdf)")

    if not args.no_check:
        say("\nПроверка оформления:")
        from gostlib import checker
        rc = checker.run(out, work=str(meta.get("work", meta.get("type", "lr"))).lower(),
                         use_word=use_word, quiet_ok=False)
        return 0 if rc == 0 else 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
