"""Титульный лист по образцу ФИТ (шаблон ИДЗ), перенесён из проекта md-to-docx.

Макет: колонтитул первой страницы (министерство, вуз, факультет, кафедра), блок «вид работы / тема / дисциплина /
вариант», блок подписей (студент, курс, группа, направление, проверяющий), внизу первой страницы – город и год.
Номер страницы на титульном листе не показывается, но страница учитывается в нумерации (реферат – стр. 2).
"""
from __future__ import annotations

import re
from datetime import date

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import config as C
from .inline import Typo, fix_dashes, fix_quotes

NBSP = " "

# Значения по умолчанию (редактируются здесь или переопределяются полями front matter)
DEFAULTS = {
    "ministry": "Министерство науки и высшего образования РФ",
    "legal_form": "федеральное государственное автономное образовательное учреждение высшего образования",
    "university": "Омский государственный технический университет",
    "faculty": "Факультет информационных технологий и компьютерных систем",
    "department": "Прикладная математика и фундаментальная информатика",
    "city": "Омск",
}
# Сокращения, которые можно писать в front matter вместо полных названий
DEPARTMENT_ALIASES = {"пмифи": "Прикладная математика и фундаментальная информатика"}
UNIVERSITY_ALIASES = {"омгту": ("Омский государственный технический университет", "Омск")}

# Без этих полей титульный лист не собирается: значения не угадываются
REQUIRED = {
    "title": "тема работы",
    "discipline": "дисциплина",
    "student": "ФИО студента",
    "course": "курс",
    "group": "группа",
    "teacher": "ФИО проверяющего",
    "teacher_title": "должность и степень проверяющего",
}


def enabled(meta: dict) -> bool:
    mode = str(meta.get("title_page", "auto")).strip().lower()
    if mode in ("none", "no", "false", "нет", "0"):
        return False
    if mode in ("fit", "true", "yes", "да", "1"):
        return True
    return bool(meta.get("student") or meta.get("teacher"))      # auto: есть данные студента


def resolve(meta: dict, work_title: str) -> tuple[dict, list[str]]:
    """Собирает значения полей; возвращает (поля, список недостающих обязательных)."""
    v = {k: meta.get(k) for k in list(REQUIRED) + ["program", "variant", "work_number", "work_type", "year",
                                                    "university", "faculty", "department", "city"]}
    if not v["student"] and meta.get("author"):
        v["student"] = meta["author"]
    missing = [f"{k} ({d})" for k, d in REQUIRED.items() if not str(v.get(k) or "").strip()]
    uni = str(v.get("university") or "").strip()
    city = v.get("city")
    if uni.casefold() in UNIVERSITY_ALIASES:
        uni, alias_city = UNIVERSITY_ALIASES[uni.casefold()]
        city = city or alias_city
    dep = str(v.get("department") or "").strip()
    dep = DEPARTMENT_ALIASES.get(dep.casefold(), dep)
    out = {
        "ministry": DEFAULTS["ministry"],
        "legal_form": DEFAULTS["legal_form"],
        "university": uni or DEFAULTS["university"],
        "faculty": str(v.get("faculty") or DEFAULTS["faculty"]).strip(),
        "department": dep or DEFAULTS["department"],
        "city": str(city or DEFAULTS["city"]).strip(),
        "year": str(v.get("year") or date.today().year).strip(),
        "title": str(v.get("title") or "").strip(),
        "discipline": str(v.get("discipline") or "").strip(),
        "student": str(v.get("student") or "").strip(),
        "course": str(v.get("course") or "").strip(),
        "group": str(v.get("group") or "").strip(),
        "program": str(v.get("program") or "").strip(),
        "teacher": str(v.get("teacher") or "").strip(),
        "teacher_title": str(v.get("teacher_title") or "").strip(),
        "variant": str(v.get("variant") or "").strip(),
    }
    wt = str(v.get("work_type") or work_title).strip()
    num = str(v.get("work_number") or "").strip()
    out["work_type"] = f"{wt} № {num}" if num else wt
    return out, missing


def _fill(value: str, cm: float) -> str:
    """Значение для подчёркивания: добивается неразрывными пробелами до ширины cm, чтобы линия была сплошной.
    Оценка: средний символ TNR 14 пт ~ 0,26 см, неразрывный пробел ~ 0,123 см."""
    return value + NBSP * max(2, int((cm - len(value) * 0.26) / 0.123))


def build(b, data: dict) -> None:
    """Добавляет титульный лист в начало документа builder-а b (до всех остальных блоков)."""
    doc = b.doc
    typo = Typo()

    def tx(s: str) -> str:
        return fix_dashes(fix_quotes(s, " ", typo), typo)

    style = b._mk_style("Титульный лист", "GostCover", "Normal", C.FONT_SIZE, False, False,
                        WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 1.5, False, None, None, C.FONT)

    def par(runs, align="center", size=14, line=1.5, left=0.0, first=0.0, before=0.0, tabs=(), keep=False):
        p = doc.add_paragraph(style=style)
        pf = p.paragraph_format
        pf.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
        pf.left_indent, pf.first_line_indent = Cm(left), Cm(first)
        pf.space_before, pf.space_after = Pt(before), Pt(0)
        pf.line_spacing = line
        for pos in tabs:
            pf.tab_stops.add_tab_stop(Cm(pos), WD_TAB_ALIGNMENT.LEFT)
        if isinstance(runs, str):
            runs = [(runs, {})]
        for text, fmt in runs:
            r = p.add_run(text)
            r.font.size = Pt(fmt.get("size", size))
            if fmt.get("bold"):
                r.bold = True
            if fmt.get("underline"):
                r.underline = True
        return p

    U = {"underline": True}
    LABEL, VALUE = 5.5, 9.0

    # ---- первая страница: вид работы, тема, дисциплина -------------------------------------
    par([(tx(data["work_type"]), {"bold": True})], before=96)
    par([(tx(f"«{data['title'].strip('«»')}»"), {"bold": True})])
    par("")
    par(tx(f"по дисциплине «{data['discipline'].strip('«»')}»"))
    if data["variant"]:
        par(f"Вариант {data['variant']}")
    par("", before=12)
    par("")

    # ---- блок исполнителя и проверяющего ----------------------------------------------------
    par([("Студента\t", {}), (_fill(data["student"], 7.3), U)], align="left", left=LABEL, tabs=[VALUE], line=1.0, before=12)
    par([("Курс\t", {}), (_fill(data["course"], 1.3), U), ("\tГруппа\t", {}), (_fill(data["group"], 3.4), U)],
        align="left", left=LABEL, tabs=[VALUE, 10.6, 13.0], before=6)
    if data["program"]:
        par([("Направление\t", {}), (data["program"], U)], align="left", left=VALUE, first=-(VALUE - LABEL),
            tabs=[VALUE], before=6)
    par([("Проверил\t", {}), (_fill(data["teacher_title"], 7.3), U)], align="left", left=LABEL, tabs=[VALUE], line=1.0, before=12)
    par([("должность", {"size": 9})], left=VALUE, line=1.0)
    par([(_fill(data["teacher"], 7.3), U)], align="left", left=VALUE, line=1.0, before=6)
    par([("Выполнил\t", {}), (_fill("", 7.3), U)], align="left", left=LABEL, tabs=[VALUE], line=1.0, before=24)
    par([("дата, подпись студента", {"size": 9})], left=VALUE, line=1.0)
    par([(_fill("", 7.3), U)], align="left", left=VALUE, line=1.0, before=18)
    par([("дата, подпись преподавателя", {"size": 9})], left=VALUE, line=1.0)

    # ---- колонтитулы первой страницы -------------------------------------------------------
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    hdr = sec.first_page_header
    hdr.is_linked_to_previous = False
    head = [
        (data["ministry"], 1, 0),
        (data["legal_form"], 3, 12),
        (f"«{data['university']}»", 0, 0),
        (data["faculty"], 12, 0),
        (f"Кафедра «{data['department']}»", 12, 0),
    ]
    for i, (text, before, after) in enumerate(head):
        p = hdr.paragraphs[0] if i == 0 else hdr.add_paragraph()
        p.style = style
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
        pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), 1.0
        r = p.add_run(tx(text))
        r.font.size = Pt(12)

    ftr = sec.first_page_footer
    ftr.is_linked_to_previous = False
    for i, text in enumerate((f"г. {data['city']}", f"{data['year']} г.")):
        p = ftr.paragraphs[0] if i == 0 else ftr.add_paragraph()
        p.style = style
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        r.font.size = Pt(14)
    for k, n in typo.fixes.items():
        b.res.notes.append(f"титульный лист: {k} ({n})")
