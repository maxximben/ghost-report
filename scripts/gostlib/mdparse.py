"""Разбор исходного Markdown отчёта в список блоков.

Формат описан в references/markdown-format.md.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from . import config as C


class MdError(Exception):
    pass


@dataclass
class Block:
    kind: str                      # heading|para|list|code|figure|table|math|pagebreak
    line: int = 0                  # номер строки в исходнике (для сообщений)
    data: dict[str, Any] = field(default_factory=dict)

    def __getitem__(self, k):
        return self.data[k]

    def get(self, k, default=None):
        return self.data.get(k, default)


# ---------------------------------------------------------------- front matter

def _scalar(v: str):
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        return v[1:-1]
    low = v.lower()
    if low in ("true", "yes", "да"):
        return True
    if low in ("false", "no", "нет"):
        return False
    if re.fullmatch(r"-?\d+", v):
        return int(v)
    if re.fullmatch(r"-?\d+\.\d+", v):
        return float(v)
    return v


def parse_front_matter(text: str) -> tuple[dict, str, int]:
    """Возвращает (meta, остаток текста, число строк, занятых front matter)."""
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return {}, text, 0
    meta: dict[str, Any] = {}
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
        ln = lines[i]
        if not ln.strip() or ln.lstrip().startswith("#"):
            continue
        m = re.match(r"^([A-Za-z_][\w-]*)\s*:\s*(.*)$", ln)
        if not m:
            raise MdError(f"front matter, строка {i + 1}: ожидается «ключ: значение», получено «{ln}»")
        meta[m.group(1).replace("-", "_")] = _scalar(m.group(2))
    if end is None:
        raise MdError("front matter не закрыт строкой «---»")
    return meta, "\n".join(lines[end + 1:]), end + 1


# ------------------------------------------------------------------- регулярки

RE_HEADING = re.compile(r"^(#{1,4})\s+(.*?)\s*(?:\{#([\w:.-]+)\})?\s*$")
RE_APPENDIX = re.compile(r"^приложение(?:\s*\[\s*(обязательное|справочное)\s*\])?\s*[:\-–—]\s*(.+)$", re.I)
RE_FIGURE = re.compile(r"^!\[(.*?)\]\((.*?)\)\s*(?:\{(.*?)\})?\s*$")
RE_TABLE_CAP = re.compile(r"^Таблица\s*:\s*(.+?)\s*(?:\{#([\w:.-]+)\})?\s*$")
RE_LIST = re.compile(r"^(\s*)([-*•–]|\d+[.)]|[а-яё][.)])\s+(.*)$")
RE_FENCE = re.compile(r"^```\s*([\w+#.-]*)\s*$")


def _attrs(s: str | None) -> dict[str, str]:
    out: dict[str, str] = {}
    if not s:
        return out
    for tok in s.split():
        if tok.startswith("#"):
            out["id"] = tok[1:]
        elif "=" in tok:
            k, v = tok.split("=", 1)
            out[k] = v
    return out


def _split_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    cells, cur, i = [], "", 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur += "|"
            i += 2
            continue
        if s[i] == "|":
            cells.append(cur.strip())
            cur = ""
        else:
            cur += s[i]
        i += 1
    cells.append(cur.strip())
    return cells


def _is_sep_row(line: str) -> bool:
    cells = _split_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{1,}:?", c.strip()) for c in cells)


def classify_heading(level: int, text: str, hid: str | None, line: int) -> Block:
    m = RE_APPENDIX.match(text)
    if m and level == 1:
        return Block("heading", line, dict(cls="appendix", level=1, text=m.group(2).strip(), id=hid))
    key = re.sub(r"\s+", " ", text.strip().rstrip(".").lower())
    if key in C.STRUCT_ALIASES:
        if level != 1:
            raise MdError(f"строка {line}: структурный элемент «{text}» должен быть заголовком первого уровня (#)")
        return Block("heading", line, dict(cls="struct", level=1, name=C.STRUCT_ALIASES[key], text=text.strip(), id=hid))
    if level > 3:
        # внутри приложения допускаются подразделы; пункт глубже 3 уровня не поддерживается
        raise MdError(f"строка {line}: заголовки глубже третьего уровня не поддерживаются («{text}»)")
    return Block("heading", line, dict(cls="section", level=level, text=text.strip(), id=hid))


def parse_blocks(text: str, first_line: int = 1) -> list[Block]:
    lines = text.split("\n")
    blocks: list[Block] = []
    i, n = 0, len(lines)
    para: list[str] = []
    para_start = 0

    def flush_para():
        nonlocal para
        if para:
            blocks.append(Block("para", para_start, dict(text=" ".join(s.strip() for s in para))))
            para = []

    def ln(idx):  # человекочитаемый номер строки
        return first_line + idx

    while i < n:
        raw = lines[i].rstrip()
        s = raw.strip()

        # комментарии
        if s.startswith("<!--"):
            low = s.lower().replace(" ", "")
            if low in ("<!--pagebreak-->", "<!--newpage-->"):
                flush_para()
                blocks.append(Block("pagebreak", ln(i)))
                i += 1
                continue
            while i < n and "-->" not in lines[i]:
                i += 1
            i += 1
            continue
        if s == "\\newpage":
            flush_para()
            blocks.append(Block("pagebreak", ln(i)))
            i += 1
            continue

        if not s:
            flush_para()
            i += 1
            continue

        # заголовок
        m = RE_HEADING.match(raw)
        if m and not raw.startswith("    "):
            flush_para()
            blocks.append(classify_heading(len(m.group(1)), m.group(2), m.group(3), ln(i)))
            i += 1
            continue

        # блок кода
        m = RE_FENCE.match(s)
        if m:
            flush_para()
            start = i
            i += 1
            code: list[str] = []
            while i < n and not lines[i].strip().startswith("```"):
                code.append(lines[i].rstrip("\n").replace("\t", "    "))
                i += 1
            if i >= n:
                raise MdError(f"строка {ln(start)}: блок кода не закрыт")
            i += 1
            blocks.append(Block("code", ln(start), dict(lang=m.group(1), lines=code)))
            continue

        # рисунок
        m = RE_FIGURE.match(s)
        if m:
            flush_para()
            a = _attrs(m.group(3))
            blocks.append(Block("figure", ln(i), dict(caption=m.group(1).strip(), path=m.group(2).strip(),
                                                      id=a.get("id"), width=a.get("width"))))
            i += 1
            continue

        # формула
        if s.startswith("$$"):
            flush_para()
            start = i
            body = s[2:]
            tail = None
            if "$$" in body:
                expr, tail = body.split("$$", 1)
                exprs = [expr]
            else:
                exprs = [body]
                i += 1
                while i < n:
                    seg = lines[i]
                    if "$$" in seg:
                        before, tail = seg.split("$$", 1)
                        exprs.append(before)
                        break
                    exprs.append(seg)
                    i += 1
                else:
                    raise MdError(f"строка {ln(start)}: формула $$ не закрыта")
            i += 1
            a = _attrs((tail or "").strip().strip("{}"))
            blocks.append(Block("math", ln(start), dict(latex=" ".join(e.strip() for e in exprs).strip(), id=a.get("id"))))
            continue

        # таблица (с подписью «Таблица: …» непосредственно выше, пустые строки допустимы)
        cap = RE_TABLE_CAP.match(s)
        if cap:
            flush_para()
            j = i + 1
            while j < n and not lines[j].strip():
                j += 1
            if j < n and lines[j].strip().startswith("|"):
                tb, i = _parse_table(lines, j, ln)
                tb.data.update(caption=cap.group(1).strip(), id=cap.group(2))
                tb.line = ln(i)
                blocks.append(tb)
                continue
            raise MdError(f"строка {ln(i)}: после «Таблица: …» ожидается таблица в формате |…|")
        if s.startswith("|"):
            raise MdError(f"строка {ln(i)}: таблица без подписи. Добавьте строку «Таблица: Название {{#tab-id}}» перед таблицей")

        # список
        m = RE_LIST.match(raw)
        if m and not para:
            flush_para()
            items, kind = [], None
            start = i
            while i < n:
                mm = RE_LIST.match(lines[i].rstrip())
                if mm and len(mm.group(1)) < 2:
                    mk = mm.group(2)
                    k = "dash" if mk in "-*•–" else ("digit" if mk[0].isdigit() else "letter")
                    if kind is None:
                        kind = k
                    elif k != kind:
                        raise MdError(f"строка {ln(i)}: в одном списке смешаны разные маркеры")
                    items.append(mm.group(3).strip())
                    i += 1
                elif mm and len(mm.group(1)) >= 2:
                    raise MdError(f"строка {ln(i)}: вложенные списки не поддерживаются. Сделайте список плоским")
                elif lines[i].strip() and lines[i].startswith((" ", "\t")) and items:
                    items[-1] += " " + lines[i].strip()
                    i += 1
                else:
                    break
            blocks.append(Block("list", ln(start), dict(style=kind, items=items)))
            continue

        # обычный абзац
        if not para:
            para_start = ln(i)
        para.append(raw)
        i += 1

    flush_para()
    return blocks


def _parse_table(lines: list[str], j: int, ln) -> tuple[Block, int]:
    rows: list[list[str]] = []
    start = j
    while j < len(lines) and lines[j].strip().startswith("|"):
        rows.append(_split_row(lines[j]))
        j += 1
    if len(rows) < 2 or not _is_sep_row(lines[start + 1]):
        raise MdError(f"строка {ln(start)}: вторая строка таблицы должна быть разделителем |---|---|")
    ncols = len(rows[0])
    aligns = []
    for c in rows[1]:
        c = c.strip()
        aligns.append("center" if c.startswith(":") and c.endswith(":") else ("right" if c.endswith(":") else "left"))
    body = rows[2:]
    for r in [rows[0]] + body:
        if len(r) != ncols:
            raise MdError(f"строка {ln(start)}: в таблице разное число столбцов ({len(r)} вместо {ncols})")
    return Block("table", ln(start), dict(header=rows[0], rows=body, aligns=aligns)), j


def parse(text: str) -> tuple[dict, list[Block]]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.startswith("﻿"):
        text = text[1:]
    meta, rest, used = parse_front_matter(text)
    return meta, parse_blocks(rest, first_line=used + 1)
