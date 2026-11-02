"""Библиографические записи по ГОСТ Р 7.0.100-2018 из структурированных данных (sources.json)."""
from __future__ import annotations

import re
from datetime import date

from . import config as C

D = f" {C.DASH} "         # разделитель областей « – »


class BibError(ValueError):
    pass


REQUIRED = {
    "book": ["title", "city", "year"],
    "article": ["title", "journal", "year"],
    "chapter": ["title", "container", "year"],
    "web": ["title", "url"],
    "standard": ["number", "title", "year"],
    "thesis": ["authors", "title", "year", "degree"],
    "raw": ["raw"],
}
TYPE_ALIASES = {"proceedings": "chapter", "conference": "chapter", "paper": "chapter", "site": "web", "repo": "web",
                "software": "web", "doc": "web", "gost": "standard", "std": "standard", "dissertation": "thesis",
                "journal": "article"}


# ------------------------------------------------------------------ авторы

_INIT = re.compile(r"^(?:[A-ZА-ЯЁ]\.?){1,3}$")


def parse_author(s: str) -> tuple[str, str]:
    """'Иванов И. И.' | 'Иванов, И.И.' | 'Иванов Иван Иванович' -> ('Иванов', 'И. И.')."""
    s = s.strip()
    toks = [t for t in re.split(r"[\s,]+", s) if t]
    if not toks:
        raise BibError("пустое имя автора")
    sur, inits = None, []
    for t in toks:
        t0 = t.rstrip(".")
        if sur is None and not (_INIT.match(t) and len(t0) <= 3 and "." in t) and not (len(t0) == 1):
            sur = t
            continue
        if _INIT.match(t) or len(t0) == 1:
            inits += [c for c in t if c.isalpha()]
        else:
            inits.append(t[0].upper())     # полное имя/отчество -> инициал
    if sur is None:
        sur = toks[-1]
        inits = [c for t in toks[:-1] for c in t if c.isalpha()]
    return sur, " ".join(f"{c}." for c in inits)


def heading_name(a: str) -> str:
    sur, ini = parse_author(a)
    return f"{sur}, {ini}" if ini else sur


def resp_name(a: str) -> str:
    sur, ini = parse_author(a)
    return f"{ini} {sur}" if ini else sur


def _responsibility(authors: list[str], extra: str | None = None) -> str:
    names = [resp_name(a) for a in authors[:3]]
    txt = ", ".join(names)
    if len(authors) > 3:
        txt += " [и др.]"
    if extra:
        txt = f"{txt} ; {extra}" if txt else extra
    return txt


def _pages(v) -> str | None:
    if v in (None, ""):
        return None
    s = str(v).strip()
    return s if re.search(r"[сC]\.|стр", s) else f"{s} с."


def _range(v: str) -> str:
    return re.sub(r"\s*[-‒–—]+\s*", C.DASH, str(v).strip())


def _url_part(d: dict, warn: list[str]) -> str | None:
    if not d.get("url"):
        return None
    acc = d.get("accessed")
    if not acc:
        warn.append(f"источник «{d.get('id', d.get('title', ''))[:40]}»: не указана дата обращения (accessed) к URL; "
                    f"поставлен шаблон «дд.мм.гггг» – замените реальной датой")
        acc = "дд.мм.гггг"
    return f"URL: {d['url']} (дата обращения: {acc})"


def _join(parts: list[str | None]) -> str:
    """Области записи разделяются «. – »; запись заканчивается точкой."""
    out = []
    for p in parts:
        if not p or not p.strip():
            continue
        p = p.strip()
        out.append(p if p.endswith((".", "?", "!")) else p + ".")
    return D.join(out)


def _place(d: dict) -> str | None:
    city, pub, year = d.get("city"), d.get("publisher"), d.get("year")
    if city and pub:
        return f"{city} : {pub}, {year}" if year else f"{city} : {pub}"
    if city:
        return f"{city}, {year}" if year else city
    if pub:
        return f"{pub}, {year}" if year else pub
    return str(year) if year else None


def _title_block(d: dict, authors: list[str]) -> tuple[str, str]:
    """(заголовок описания или '', основное заглавие со сведениями и ответственностью)"""
    title = d["title"].strip().rstrip(".")
    if d.get("subtitle"):
        title += f" : {d['subtitle'].strip()}"
    head = heading_name(authors[0]) if 1 <= len(authors) <= 3 else ""
    resp_extra = d.get("responsibility")
    if authors or resp_extra:
        title += " / " + _responsibility(authors, resp_extra)
    return head, title


def _first(head: str, title: str) -> str:
    return f"{head} {title}" if head else title


# ------------------------------------------------------------------ типы

def fmt_book(d, warn):
    authors = d.get("authors") or []
    head, title = _title_block(d, authors)
    isbn = f"ISBN {d['isbn']}" if d.get("isbn") else None
    return _join([_first(head, title), d.get("edition"), _place(d), _pages(d.get("pages")), isbn, _url_part(d, warn)])


def fmt_article(d, warn):
    authors = d.get("authors") or []
    head, title = _title_block(d, authors)
    vol = []
    if d.get("volume"):
        vol.append(f"Т. {d['volume']}")
    if d.get("issue"):
        vol.append(f"№ {d['issue']}")
    if d.get("number"):
        vol.append(str(d["number"]))
    pages = f"С. {_range(d['pages'])}" if d.get("pages") else None
    first = f"{_first(head, title)} // {d['journal'].strip().rstrip('.')}"
    doi = f"DOI: {d['doi']}" if d.get("doi") else None
    return _join([first, str(d["year"]), ", ".join(vol) or None, pages, doi, _url_part(d, warn)])


def fmt_chapter(d, warn):
    authors = d.get("authors") or []
    head, title = _title_block(d, authors)
    cont = d["container"].strip().rstrip(".")
    if d.get("conference"):
        cont += f" : {d['conference'].strip()}"
    if d.get("editors"):
        cont += f" / {d['editors']}"
    first = f"{_first(head, title)} // {cont}"
    pages = f"С. {_range(d['pages'])}" if d.get("pages") else None
    place = _place(d)
    return _join([first, place, pages, _url_part(d, warn)])


def fmt_web(d, warn):
    authors = d.get("authors") or []
    head, title = _title_block(d, authors)
    first = _first(head, title)
    if d.get("kind"):
        first += f" : {d['kind']}"
    if d.get("site"):
        first += f" // {d['site'].strip().rstrip('.')}"
    place = _place({"city": d.get("city"), "year": d.get("year")})
    return _join([first, place, _url_part(d, warn)])


def fmt_standard(d, warn):
    first = f"{d['number'].strip()}. {d['title'].strip().rstrip('.')}"
    if d.get("kind"):
        first += f" : {d['kind']}"
    if d.get("official", True):
        first += " : изд. офиц."
    if d.get("approved"):
        first += f" : {d['approved'].strip().rstrip('.')}"
    return _join([first, _place(d), _pages(d.get("pages")), _url_part(d, warn)])


def fmt_thesis(d, warn):
    authors = d.get("authors") or []
    head = heading_name(authors[0])
    title = f"{d['title'].strip().rstrip('.')} : {d['degree']}"
    if d.get("specialty"):
        title += f" : {d['specialty']}"
    title += " / " + _responsibility(authors)
    if d.get("organization"):
        title += f" ; {d['organization']}"
    place = f"{d['city']}, {d['year']}" if d.get("city") else str(d["year"])
    return _join([f"{head} {title}", place, _pages(d.get("pages")), _url_part(d, warn)])


def fmt_raw(d, warn):
    t = d["raw"].strip()
    return t if t.endswith(".") else t + "."


FORMATTERS = {"book": fmt_book, "article": fmt_article, "chapter": fmt_chapter, "web": fmt_web,
              "standard": fmt_standard, "thesis": fmt_thesis, "raw": fmt_raw}


def normalize_entry(d: dict) -> dict:
    d = dict(d)
    t = (d.get("type") or "book").lower()
    d["type"] = TYPE_ALIASES.get(t, t)
    if isinstance(d.get("authors"), str):
        d["authors"] = [a.strip() for a in re.split(r";|\n", d["authors"]) if a.strip()]
    return d


def format_entry(d: dict, warn: list[str]) -> str:
    d = normalize_entry(d)
    t = d["type"]
    if t not in FORMATTERS:
        raise BibError(f"источник «{d.get('id', '?')}»: неизвестный type «{t}». Допустимо: {', '.join(FORMATTERS)}")
    missing = [k for k in REQUIRED[t] if not d.get(k)]
    if missing:
        raise BibError(f"источник «{d.get('id', '?')}» (type={t}): не заполнены поля: {', '.join(missing)}")
    return FORMATTERS[t](d, warn)


# ------------------------------------------------------------- сортировка

def sort_key(d: dict) -> tuple:
    d = normalize_entry(d)
    authors = d.get("authors") or []
    if d["type"] == "standard":
        key = d["number"]
    elif d["type"] == "raw":
        key = d["raw"]
    elif authors and len(authors) <= 3:
        key = heading_name(authors[0])
    else:
        key = d["title"]
    key = re.sub(r"^[«\"'\s]+", "", key).lower().replace("ё", "е")
    first = key[:1]
    group = 0 if re.match(r"[а-я]", first) else (1 if re.match(r"[a-z]", first) else 2)
    return (group, key)


def sort_sources(items: list[dict]) -> list[dict]:
    return sorted(items, key=sort_key)


def is_electronic(d: dict) -> bool:
    return bool(d.get("url")) or normalize_entry(d)["type"] == "web"


def today_ru() -> str:
    return date.today().strftime("%d.%m.%Y")
