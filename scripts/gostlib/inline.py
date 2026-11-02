"""Разбор строчной разметки и типографика.

Поддерживается: **жирный**, *курсив*, `код`, $формула$, [@источник], [[fig:id]] (ссылка), <br>.
Типографика: «ёлочки», короткое тире, автокурсив для латиницы.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import config as C


@dataclass
class Span:
    kind: str = "text"      # text|code|math|cite|ref|br
    text: str = ""
    bold: bool = False
    italic: bool = False
    ids: list = field(default_factory=list)   # cite: [(id, extra)]
    ref: tuple | None = None                  # ref: (kind, id)


_TOKEN = re.compile(
    r"(?P<code>`[^`]+`)"
    r"|(?P<math>(?<!\$)\$(?!\$)(?P<mbody>[^$\n]+?)(?<!\\)\$(?!\$))"
    r"|(?P<cite>\[@[^\]]+\])"
    r"|(?P<ref>\[\[(?:fig|tab|eq|app|sec)-[\w.:-]+\]\])"
    r"|(?P<br><br\s*/?>)"
    r"|(?P<bold>\*\*(?=\S)(?P<bbody>.+?)(?<=\S)\*\*)"
    r"|(?P<ital>(?<![\w*])\*(?=[^\s*])(?P<ibody>[^*]+?)(?<=[^\s*])\*(?![\w*]))"
)


def parse_inline(s: str, bold: bool = False, italic: bool = False) -> list[Span]:
    spans: list[Span] = []
    pos = 0
    for m in _TOKEN.finditer(s):
        if m.start() > pos:
            spans.append(Span("text", s[pos:m.start()], bold, italic))
        if m.group("code"):
            spans.append(Span("code", m.group("code")[1:-1], bold, italic))
        elif m.group("math"):
            spans.append(Span("math", m.group("mbody").strip(), bold, italic))
        elif m.group("cite"):
            body = m.group("cite")[1:-1]
            ids = []
            for part in body.split(";"):
                part = part.strip()
                mm = re.match(r"@([\w.:-]+)\s*(?:,\s*(.+))?$", part)
                if not mm:
                    raise ValueError(f"не удалось разобрать ссылку на источник «[{body}]»")
                ids.append((mm.group(1), (mm.group(2) or "").strip()))
            spans.append(Span("cite", "", bold, italic, ids=ids))
        elif m.group("ref"):
            ident = m.group("ref")[2:-2]
            spans.append(Span("ref", "", bold, italic, ref=(ident.split("-", 1)[0], ident)))
        elif m.group("br"):
            spans.append(Span("br"))
        elif m.group("bold"):
            spans.extend(parse_inline(m.group("bbody"), True, italic))
        elif m.group("ital"):
            spans.extend(parse_inline(m.group("ibody"), bold, True))
        pos = m.end()
    if pos < len(s):
        spans.append(Span("text", s[pos:], bold, italic))
    return spans


# ------------------------------------------------------------------ типографика

class Typo:
    """Накапливает информацию об автоисправлениях, чтобы сообщить автору."""

    def __init__(self):
        self.fixes: dict[str, int] = {}

    def note(self, what: str, n: int = 1):
        if n:
            self.fixes[what] = self.fixes.get(what, 0) + n


_OPEN_PREV = set(" \t\n([{\u00ab\u2013\u2014/\u00a0")


def fix_quotes(t: str, prev: str = " ", typo: Typo | None = None) -> str:
    """Любые кавычки -> «ёлочки». prev — символ перед фрагментом (для парных кавычек между спанами)."""
    out, changed = [], 0
    for ch in t:
        if ch in '"“”„‟':
            if ch in "“„":
                is_open = True
            elif ch in "”‟":
                is_open = False
            else:
                is_open = prev in _OPEN_PREV
            out.append("«" if is_open else "»")
            changed += 1
            prev = out[-1]
        else:
            out.append(ch)
            prev = ch
    if typo:
        typo.note("кавычки приведены к «ёлочкам»", changed)
    return "".join(out)


def fix_dashes(t: str, typo: Typo | None = None) -> str:
    n = 0

    def sub(pattern, repl, s):
        nonlocal n
        s2, k = re.subn(pattern, repl, s)
        n += k
        return s2

    t = sub(r"\s*—\s*", " – ", t)        # длинное тире -> короткое
    t = sub(r"\s+---\s+|\s+--\s+", " – ", t)
    t = sub(r"(?<=\s)-(?=\s)", "–", t)        # дефис между пробелами -> тире
    if typo:
        typo.note("тире приведены к короткому «–»", n)
    return t


_LATIN = re.compile(r"(?<![\w/@.:#-])[A-Za-z][A-Za-z0-9+#_.\-]*(?:[A-Za-z0-9+#])?")
_ROMAN = re.compile(r"^(?=[IVXLCDM])M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")
_URL = re.compile(r"https?://[^\s)]+|www\.[^\s)]+|\b[\w.-]+@[\w.-]+\.\w+\b")


def split_latin(t: str) -> list[tuple[str, bool]]:
    """Делит текст на (кусок, это_латиница). URL и римские цифры латиницей не считаются."""
    protected = [(m.start(), m.end()) for m in _URL.finditer(t)]
    parts: list[tuple[str, bool]] = []
    pos = 0
    for m in _LATIN.finditer(t):
        a, b = m.start(), m.end()
        word = m.group(0).rstrip(".-_")
        b = a + len(word)
        if any(ps <= a < pe for ps, pe in protected):
            continue
        if len(word) < 2 or _ROMAN.match(word):
            continue
        if a > pos:
            parts.append((t[pos:a], False))
        parts.append((t[a:b], True))
        pos = b
    if pos < len(t):
        parts.append((t[pos:], False))
    return parts or [(t, False)]


def normalize_spans(spans: list[Span], typo: Typo | None = None, auto_italic: bool = True) -> list[Span]:
    """Применяет типографику к текстовым спанам; латиницу делает курсивом."""
    out: list[Span] = []
    prev = " "
    for sp in spans:
        if sp.kind != "text":
            out.append(sp)
            prev = "a" if sp.kind != "br" else " "
            continue
        t = fix_dashes(fix_quotes(sp.text, prev, typo), typo)
        t = re.sub(r"(?<=\S) {2,}(?=\S)", " ", t)
        if t:
            prev = t[-1]
        if auto_italic and not sp.italic:
            for piece, latin in split_latin(t):
                if piece:
                    out.append(Span("text", piece, sp.bold, True if latin else sp.italic))
        else:
            out.append(Span("text", t, sp.bold, sp.italic))
    return out
