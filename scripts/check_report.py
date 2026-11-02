#!/usr/bin/env python3
"""Проверка оформления готового .docx по правилам (шрифт, поля, отступы, заголовки, подписи, источники …).

    python check_report.py report.docx --work lr|dz|rgr|practice [--word] [--strict] [--json]

Код возврата: 0 – ошибок нет; 1 – есть ошибки (или предупреждения при --strict).
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gostlib import checker  # noqa: E402


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--work", default="lr", choices=["lr", "dz", "rgr", "practice"], help="вид работы (влияет на обязательные разделы и интервал)")
    ap.add_argument("--word", action="store_true", help="дополнительно проверить вёрстку через Microsoft Word (переносы таблиц)")
    ap.add_argument("--strict", action="store_true", help="считать предупреждения ошибками")
    ap.add_argument("--font", default="Times New Roman", help="основной шрифт (по умолчанию Times New Roman)")
    ap.add_argument("--line-spacing", type=float, help="ожидаемый междустрочный интервал (по умолчанию по виду работы)")
    ap.add_argument("--json", action="store_true", help="вывести замечания в JSON")
    a = ap.parse_args()
    sys.exit(checker.run(a.docx, work=a.work, use_word=a.word, strict=a.strict, font=a.font,
                         line_spacing=a.line_spacing, as_json=a.json))


if __name__ == "__main__":
    main()
