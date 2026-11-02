"""Проверка оформления .docx отчёта по «Правилам оформления отчётов и ПЗ» и ГОСТ Р 7.0.100-2018."""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from . import config as C

ERR, WARN = "ERROR", "WARN"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

CODE_FONTS = {"dejavu sans", "dejavu sans mono", "consolas", "courier new", "liberation mono", "cascadia code", "lucida console"}
STRUCT_NAMES = {C.ABSTRACT, C.TOC, C.INTRO, C.CONCL, C.BIB}
RE_APPX = re.compile(r"^ПРИЛОЖЕНИЕ\s+([А-ЯЁ])\b")
RE_SECTION = re.compile(r"^((?:[А-ЯЁ]\.)?\d+(?:\.\d+){0,2})(\.?)\s+(\S.*)$")   # «1.2» или «А.1» (подраздел приложения)
RE_FIG = re.compile(r"^Рисунок\s+((?:[А-Я]\.)?\d+(?:\.\d+)?)\s*([–—-])\s*(.*)$", re.S)
RE_TAB = re.compile(r"^Таблица\s+((?:[А-Я]\.)?\d+(?:\.\d+)?)\s*([–—-])\s*(.*)$", re.S)
RE_CONT = re.compile(r"^Продолжение таблицы\s+((?:[А-Я]\.)?\d+(?:\.\d+)?)\s*$")
RE_LATIN = re.compile(r"(?<![\w/@.:#-])[A-Za-z][A-Za-z0-9+#_\-]*[A-Za-z0-9+#]")
RE_ROMAN = re.compile(r"^(?=[IVXLCDM])M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")
RE_URL = re.compile(r"https?://\S+|www\.\S+|\b[\w.-]+@[\w.-]+\.\w+\b")
PLACEHOLDERS = ("число, месяц, год", "дд.мм.гггг", "TODO", "???", "Lorem ipsum", "XXX")


@dataclass
class Finding:
    sev: str
    code: str
    msg: str
    loc: str = ""
    snippet: str = ""


@dataclass
class Item:
    idx: int
    kind: str = "body"          # body|struct|appendix|section|figcap|tabcap|cont|code|list|empty|image|toc|eq|bib|other
    text: str = ""
    par: Paragraph | None = None
    in_table: bool = False
    meta: dict = field(default_factory=dict)


def clean(t: str) -> str:
    return t.replace("\x0b", " ").replace("\n", " ").replace(" ", " ").strip()


# ============================================================== свойства
class Props:
    def __init__(self, doc):
        self.doc = doc
        self.def_font = None
        self.def_size = None
        dd = doc.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rf = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:rFonts"))
            if rf is not None:
                self.def_font = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
            sz = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:sz"))
            if sz is not None:
                self.def_size = int(sz.get(qn("w:val"))) / 2

    def par(self, par: Paragraph, attr: str):
        v = getattr(par.paragraph_format, attr)
        if v is not None:
            return v
        st = par.style
        while st is not None:
            v = getattr(st.paragraph_format, attr)
            if v is not None:
                return v
            st = st.base_style
        return None

    def run(self, run, par: Paragraph, attr: str):
        v = getattr(run.font, attr)
        if v is not None:
            return v
        st = run.style
        while st is not None:
            v = getattr(st.font, attr)
            if v is not None:
                return v
            st = st.base_style
        st = par.style
        while st is not None:
            v = getattr(st.font, attr)
            if v is not None:
                return v
            st = st.base_style
        if attr == "name":
            return self.def_font
        if attr == "size":
            from docx.shared import Pt
            return Pt(self.def_size) if self.def_size else None
        return None

    def run_pt(self, run, par):
        sz = self.run(run, par, "size")
        return sz.pt if sz is not None else None

    def line_mult(self, par: Paragraph, base_pt: float = 14.0):
        v = self.par(par, "line_spacing")
        if v is None:
            return 1.0
        if isinstance(v, float):
            return v
        return float(v.pt) / (base_pt * 1.15)    # exact/atLeast приблизительно


# ============================================================== проверка
class Checker:
    def __init__(self, path: str, work: str = "lr", font: str = C.FONT, line_spacing: float | None = None):
        self.path = path
        self.doc = Document(path)
        self.pr = Props(self.doc)
        self.work = work if work in C.PROFILES else "lr"
        self.profile = C.PROFILES[self.work]
        self.font = font
        self.ls = line_spacing or self.profile["line_spacing"]
        self.f: list[Finding] = []
        self.items: list[Item] = []
        self.numbering = self._load_numbering()

    # ------------------------------------------------------------ утилиты
    def add(self, sev, code, msg, it: Item | None = None):
        loc, snip = "", ""
        if it is not None:
            loc = f"абзац {it.idx}" + (" (таблица)" if it.in_table else "")
            snip = clean(it.text)[:70]
        self.f.append(Finding(sev, code, msg, loc, snip))

    def _load_numbering(self):
        """numId -> (numFmt, lvlText) для уровня 0."""
        out = {}
        try:
            part = self.doc.part.numbering_part
        except Exception:
            return out
        el = part.element
        abstracts = {}
        for a in el.findall(qn("w:abstractNum")):
            aid = a.get(qn("w:abstractNumId"))
            lv = a.find(qn("w:lvl"))
            if lv is not None:
                fmt = lv.find(qn("w:numFmt"))
                txt = lv.find(qn("w:lvlText"))
                abstracts[aid] = (fmt.get(qn("w:val")) if fmt is not None else None,
                                  txt.get(qn("w:val")) if txt is not None else None)
        for n in el.findall(qn("w:num")):
            aid = n.find(qn("w:abstractNumId"))
            if aid is not None:
                out[n.get(qn("w:numId"))] = abstracts.get(aid.get(qn("w:val")), (None, None))
        return out

    def num_of(self, par: Paragraph):
        ppr = par._p.pPr
        if ppr is not None and ppr.numPr is not None and ppr.numPr.numId is not None:
            return str(ppr.numPr.numId.val)
        st = par.style
        while st is not None:
            sppr = st.element.pPr
            if sppr is not None and sppr.numPr is not None and sppr.numPr.numId is not None:
                return str(sppr.numPr.numId.val)
            st = st.base_style
        return None

    # --------------------------------------------------------- обход документа
    def collect(self):
        body = self.doc.element.body
        idx = 0

        def add_par(p_el, in_table=False):
            nonlocal idx
            par = Paragraph(p_el, self.doc)
            idx += 1
            self.items.append(Item(idx, text=par.text, par=par, in_table=in_table))

        for ch in body.iterchildren():
            if ch.tag == qn("w:p"):
                add_par(ch)
            elif ch.tag == qn("w:tbl"):
                tbl = Table(ch, self.doc)
                seen = set()
                tstart = idx + 1
                for row in tbl.rows:
                    for cell in row.cells:
                        if id(cell._tc) in seen:
                            continue
                        seen.add(id(cell._tc))
                        for p_el in cell._tc.findall(qn("w:p")):
                            add_par(p_el, True)
                for it in self.items[tstart - 1:]:
                    it.meta["table"] = tbl
                    it.meta["table_el"] = ch
        self.classify()
        self.mark_cover()

    def classify(self):
        in_bib = False
        for it in self.items:
            par = it.par
            t = clean(it.text)
            it.text = it.text
            el = par._p
            has_draw = bool(el.findall(".//" + qn("w:drawing"))) or bool(el.findall(".//" + qn("w:pict")))
            has_math = bool(el.findall(".//{%s}oMath" % M_NS)) or bool(el.findall(".//{%s}oMathPara" % M_NS))
            style = (par.style.name if par.style is not None else "") or ""
            in_toc = style.lower().startswith("toc") or bool(el.findall(".//" + qn("w:hyperlink"))) and \
                any((h.get(qn("w:anchor")) or "").startswith("_Toc") for h in el.findall(".//" + qn("w:hyperlink")))
            it.meta["has_draw"], it.meta["has_math"] = has_draw, has_math
            if in_toc and not it.in_table:
                it.kind = "toc"
                continue
            if it.in_table:
                it.kind = "eq" if has_math else "cell"
                continue
            if has_draw and not t:
                it.kind = "image"
                continue
            if not t:
                it.kind = "empty" if not has_math else "eq"
                continue
            if t.upper() == t and t in STRUCT_NAMES or (t.upper() == t and re.sub(r"\s+", " ", t) in STRUCT_NAMES):
                it.kind, it.meta["name"] = "struct", re.sub(r"\s+", " ", t)
                in_bib = it.meta["name"] == C.BIB
                continue
            m = RE_APPX.match(t)
            if m and len(t) < 160:
                it.kind, it.meta["letter"] = "appendix", m.group(1)
                in_bib = False
                continue
            if RE_CONT.match(t):
                it.kind = "cont"
                continue
            if RE_FIG.match(t):
                it.kind = "figcap"
                continue
            if RE_TAB.match(t):
                it.kind = "tabcap"
                continue
            if style.lower() in ("код", "code", "source code") or self._is_code(par):
                it.kind = "code"
                continue
            mnum = RE_SECTION.match(t)
            if mnum and not in_bib and len(t) < 160 and (style.lower().startswith(("heading", "заголовок")) or self._all_bold(par)):
                it.kind = "section"
                it.meta["num"], it.meta["dot"] = mnum.group(1), mnum.group(2)
                continue
            if in_bib:
                it.kind = "bib"
                continue
            if self.num_of(par) is not None or re.match(r"^([–—\-•·*]|\d+[.)]|[а-яё]\))\s", t):
                it.kind = "list"
                continue
            it.kind = "body"

    COVER_HINT = re.compile(r"Министерство|по дисциплине|Студента|Выполнил|Проверил|Кафедра|Факультет|^г\. [А-ЯЁ]")

    def mark_cover(self):
        """Всё до первого заголовка/оглавления – титульный лист (если похоже на титульник)."""
        self.has_cover = False
        first = next((it for it in self.items if it.kind in ("struct", "appendix", "section", "toc")), None)
        if first is None or first.idx == 1:
            return
        head = [it for it in self.items[:first.idx - 1] if not it.in_table]
        if not any(self.COVER_HINT.search(clean(it.text)) for it in head):
            return
        self.has_cover = True
        for it in head:
            if it.kind not in ("empty", "image"):
                it.kind = "cover"

    def _runs(self, par):
        return [r for r in par.runs if r.text.strip()]

    def _is_code(self, par) -> bool:
        runs = self._runs(par)
        if not runs:
            return False
        return all((self.pr.run(r, par, "name") or "").lower() in CODE_FONTS for r in runs)

    def _all_bold(self, par) -> bool:
        runs = self._runs(par)
        return bool(runs) and all(bool(self.pr.run(r, par, "bold")) for r in runs)

    # ------------------------------------------------------------------ запуск
    def run_all(self, probe: dict | None = None):
        self.collect()
        self.check_page()
        self.check_text()
        self.check_fonts_and_paragraphs()
        self.check_structure()
        self.check_headings()
        self.check_captions()
        self.check_lists()
        self.check_code()
        self.check_formulas()
        self.check_bibliography()
        self.check_toc()
        if probe:
            self.check_layout(probe)
        return self.f

    # ---------------------------------------------------------------- страница
    def check_page(self):
        s = self.doc.sections[0]
        want = dict(top=C.MARGIN_TOP, bottom=C.MARGIN_BOTTOM, left=C.MARGIN_LEFT, right=C.MARGIN_RIGHT)
        got = dict(top=s.top_margin, bottom=s.bottom_margin, left=s.left_margin, right=s.right_margin)
        ru = dict(top="верхнее", bottom="нижнее", left="левое", right="правое")
        for k, v in want.items():
            if got[k] is None or abs(got[k].cm - v) > 0.05:
                self.add(ERR, "P01", f"{ru[k]} поле {got[k].cm:.2f} см, нужно {v} см" if got[k] is not None else f"не задано {ru[k]} поле")
        if s.page_width is not None and (abs(s.page_width.cm - C.PAGE_W_CM) > 0.1 or abs(s.page_height.cm - C.PAGE_H_CM) > 0.1):
            self.add(ERR, "P02", f"формат страницы {s.page_width.cm:.1f}×{s.page_height.cm:.1f} см, нужен A4 (21×29,7)")
        footers = "".join(p._p.xml for p in s.footer.paragraphs)
        if "PAGE" not in footers:
            self.add(WARN, "P03", "в нижнем колонтитуле нет номера страницы (поле PAGE)")

    # ------------------------------------------------------------------ текст
    def check_text(self):
        for it in self.items:
            if it.kind in ("code", "empty", "image", "toc"):
                continue
            t = it.text
            tt = clean(t)
            if it.kind == "cover":
                if "\u2014" in t:
                    self.add(ERR, "T01", "длинное тире «—» не используется: замените на короткое «–»", it)
                if re.search(r"\{[A-Za-z_]+\}", t):
                    self.add(ERR, "T06", "на титульном листе не заполнено поле в фигурных скобках", it)
                continue
            if "—" in t:
                self.add(ERR, "T01", "длинное тире «—» не используется: замените на короткое «–»", it)
            if re.search(r'["“”„‟]', t):
                self.add(ERR, "T02", "кавычки должны быть только «ёлочками»", it)
            if it.kind not in ("bib", "cell") and re.search(r"(?<=\S) - (?=\S)", t):
                self.add(WARN, "T03", "дефис с пробелами вместо тире: используйте «–»", it)
            if "\t" in t:
                self.add(ERR, "T04", "в тексте символ табуляции: никакой табуляции, отступ задаётся форматом абзаца", it)
            if any(re.search(r"(?<=\S)  +(?=\S)", r.text) for r in it.par.runs):
                self.add(WARN, "T05", "двойные пробелы", it)
            for ph in PLACEHOLDERS:
                if ph.lower() in tt.lower():
                    self.add(ERR, "T06", f"незаполненный шаблон «{ph}»", it)
                    break

    # -------------------------------------------------- шрифты и форматы абзацев
    def check_fonts_and_paragraphs(self):
        allowed = {self.font.lower(), C.MATH_FONT.lower()}
        bad_fonts: dict[str, int] = {}
        for it in self.items:
            par = it.par
            if it.kind in ("empty", "toc"):
                continue
            for r in self._runs(par):
                name = (self.pr.run(r, par, "name") or "").lower()
                size = self.pr.run_pt(r, par)
                if it.kind == "code" or name in CODE_FONTS and it.kind != "body":
                    if name != C.CODE_FONT.lower() and it.kind == "code":
                        self.add(WARN, "K02", f"код набран шрифтом «{name}», по правилам нужен {C.CODE_FONT}", it)
                    if size and it.kind == "code" and abs(size - C.CODE_SIZE) > 0.1:
                        self.add(ERR, "K03", f"размер шрифта кода {size:g} пт, нужно {C.CODE_SIZE}", it)
                    continue
                if name in CODE_FONTS:      # встроенный в текст код
                    continue
                if name not in allowed:
                    bad_fonts[name or "(не задан)"] = bad_fonts.get(name or "(не задан)", 0) + 1
                if it.kind == "cover":
                    continue
                if size is not None:
                    if it.in_table:
                        if size < C.TABLE_MIN_SIZE - 0.01 or size > C.FONT_SIZE + 0.01:
                            self.add(ERR, "F02", f"размер шрифта в таблице {size:g} пт (допустимо 10–14)", it)
                    elif abs(size - C.FONT_SIZE) > 0.01:
                        self.add(ERR, "F02", f"размер шрифта {size:g} пт, нужно {C.FONT_SIZE} (уменьшать можно только в таблицах)", it)
                        break
            if it.in_table or it.kind in ("code", "image", "eq", "cover"):
                continue
            self._check_par_format(it)
        for name, n in bad_fonts.items():
            self.f.append(Finding(ERR, "F01", f"шрифт «{name}» в {n} фрагментах; основной шрифт – {self.font}", "документ", ""))

    def _check_par_format(self, it: Item):
        par, pr = it.par, self.pr
        kind = it.kind
        al = pr.par(par, "alignment")
        al_name = str(al).split(".")[-1].split(" ")[0] if al is not None else "LEFT"
        fi = pr.par(par, "first_line_indent")
        fi_cm = fi.cm if fi is not None else 0.0
        li = pr.par(par, "left_indent")
        has_num = self.num_of(par) is not None
        ls = pr.line_mult(par)
        if kind in ("body", "list", "bib"):
            if al_name != "JUSTIFY":
                self.add(ERR, "L01", "выравнивание должно быть «по ширине»", it)
            if not has_num or kind == "body":
                if abs(fi_cm - C.INDENT_CM) > 0.06 and not (kind == "list" and fi_cm == 0 and re.match(r"^([–\-•]|\d+[.)])", clean(it.text))):
                    self.add(ERR, "L02", f"абзацный отступ {fi_cm:.2f} см, нужен {C.INDENT_CM} см", it)
                if li is not None and li.cm > 0.05 and kind == "body":
                    self.add(WARN, "L02", f"абзац имеет левый отступ {li.cm:.2f} см", it)
            if abs(ls - self.ls) > 0.06:
                self.add(ERR, "L03", f"междустрочный интервал {ls:g}, нужен {self.ls:g}", it)
            if re.match(r"^\s{2,}\S", it.text):
                self.add(ERR, "L04", "отступ сделан пробелами: используйте абзацный отступ", it)
        elif kind in ("struct", "appendix"):
            if al_name != "CENTER":
                self.add(ERR, "H01", "заголовок структурного элемента – по центру", it)
            if fi_cm > 0.05:
                self.add(ERR, "H02", "у заголовка структурного элемента нет абзацного отступа (должен быть 0)", it)
            if self._any_bold(par):
                self.add(ERR, "H03", "заголовок структурного элемента – без жирного", it)
        elif kind == "figcap":
            if al_name != "CENTER":
                self.add(ERR, "C01", "подпись рисунка – по центру", it)
            if fi_cm > 0.05:
                self.add(ERR, "C02", "у подписи рисунка нет абзацного отступа", it)
        elif kind in ("tabcap", "cont"):
            if al_name not in ("LEFT", "JUSTIFY") or fi_cm > 0.05:
                self.add(ERR, "C03", "название таблицы – слева, без абзацного отступа", it)

    def _any_bold(self, par) -> bool:
        return any(self.pr.run(r, par, "bold") for r in self._runs(par))

    # ----------------------------------------------------------- структура
    def page_break_before(self, it: Item) -> bool:
        par = it.par
        if it.idx == 1:
            return True
        if self.pr.par(par, "page_break_before"):
            return True
        # предыдущий непустой абзац заканчивается разрывом страницы, либо идёт пустой с разрывом
        k = it.idx - 2
        while k >= 0:
            prev = self.items[k]
            xml = prev.par._p.xml
            if 'w:type="page"' in xml or "<w:sectPr" in xml:
                return True
            if prev.kind != "empty":
                return False
            k -= 1
        return False

    def check_structure(self):
        structs = [it for it in self.items if it.kind in ("struct", "appendix")]
        names = [it.meta.get("name") or C.APPENDIX for it in structs]
        for req in self.profile["required"]:
            if req not in names:
                self.f.append(Finding(ERR, "S01", f"нет обязательного элемента «{req}» для вида работы «{self.profile['title']}»", "документ", ""))
        for rec in self.profile.get("recommended", []):
            if rec not in names:
                self.f.append(Finding(WARN, "S01", f"рекомендуется элемент «{rec}»", "документ", ""))
        has_sections = any(it.kind == "section" for it in self.items)
        if not has_sections:
            self.f.append(Finding(ERR, "S02", "не найдены заголовки разделов основной части (формат «1 Название»)", "документ", ""))
        # порядок
        order = {n: i for i, n in enumerate(C.STRUCT_ORDER)}
        seq = []
        first_sec_done = False
        for it in self.items:
            if it.kind == "struct":
                seq.append((order[it.meta["name"]], it.meta["name"], it))
            elif it.kind == "section" and not first_sec_done:
                first_sec_done = True
                seq.append((order["ОСНОВНАЯ ЧАСТЬ"], "основная часть", it))
            elif it.kind == "appendix":
                seq.append((order[C.APPENDIX], C.APPENDIX, it))
        for a, b in zip(seq, seq[1:]):
            if b[0] < a[0]:
                self.add(ERR, "S03", f"нарушен порядок элементов: «{b[1]}» после «{a[1]}»", b[2])
                break
        dupl = {n for n in names if n != C.APPENDIX and names.count(n) > 1}
        for n in dupl:
            self.f.append(Finding(ERR, "S04", f"элемент «{n}» встречается более одного раза", "документ", ""))
        # каждая структурная единица – с новой страницы
        for it in structs:
            if not self.page_break_before(it):
                self.add(ERR, "S05", "каждая структурная единица должна начинаться с новой страницы", it)
        # приложения
        letters = [it.meta["letter"] for it in structs if it.kind == "appendix"]
        want = list(C.APPENDIX_LETTERS[:len(letters)])
        if letters and letters != want:
            self.f.append(Finding(ERR, "S06", f"буквы приложений {letters}, ожидается {want} (без Ё, З, Й, О, Ч, Ь, Ы, Ъ)", "документ", ""))
        self._check_special_sections()

    def _unit_texts(self, name):
        out, on = [], False
        for it in self.items:
            if it.kind in ("struct", "appendix"):
                on = it.kind == "struct" and it.meta["name"] == name
                continue
            if it.kind == "section" and name != "ОСНОВНАЯ ЧАСТЬ":
                on = False
            if on and it.kind in ("body", "list"):
                out.append(it)
        return out

    def _check_special_sections(self):
        # реферат
        ab = self._unit_texts(C.ABSTRACT)
        if ab:
            txt = " ".join(clean(i.text) for i in ab)
            if not re.search(r"^Отчет\s+\d+\s*с\.", clean(ab[0].text)) and not re.search(r"Отчет\s+\d+\s*с\.", txt):
                self.add(ERR, "R01", "в реферате нет строки вида «Отчет 52 с., 36 рис., 5 табл., 9 источн.»", ab[0])
            kw = [i for i in ab if clean(i.text).upper() == clean(i.text) and "," in i.text]
            if not kw:
                self.f.append(Finding(ERR, "R02", "в реферате нет строки ключевых слов (ПРОПИСНЫМИ через запятую)", "реферат", ""))
            for pat, what in ((r"Объект", "Объект"), (r"Предмет", "Предмет"), (r"Цел[ьи]", "Цель"),
                              (r"В процессе", "«В процессе выполнения…»"), (r"В результате", "«В результате выполнения…»")):
                if not re.search(pat, txt):
                    self.f.append(Finding(WARN, "R03", f"в реферате не найден элемент: {what}", "реферат", ""))
        # введение и заключение (оценка объёма по словам: ~330 слов на страницу)
        intro = self._unit_texts(C.INTRO)
        if intro:
            words = sum(len(clean(i.text).split()) for i in intro)
            txt = " ".join(clean(i.text) for i in intro).lower()
            if words < 250:
                self.f.append(Finding(WARN, "V01", f"введение короткое (~{words} слов ≈ {words / 330:.1f} стр.): по правилам ~2 стр., не менее 1", "введение", ""))
            if "цел" not in txt:
                self.f.append(Finding(WARN, "V02", "во введении не найдена цель работы", "введение", ""))
            if "задач" not in txt:
                self.f.append(Finding(WARN, "V02", "во введении нет перечня задач (они задают структуру основной части)", "введение", ""))
        concl = self._unit_texts(C.CONCL)
        if concl:
            words = sum(len(clean(i.text).split()) for i in concl)
            if words < 100:
                self.f.append(Finding(WARN, "V03", f"заключение короткое (~{words} слов): по правилам 0,5–1 стр.", "заключение", ""))
            elif words > 420:
                self.f.append(Finding(WARN, "V03", f"заключение длинное (~{words} слов ≈ {words / 330:.1f} стр.): по правилам 0,5–1 стр.", "заключение", ""))

    # --------------------------------------------------------------- заголовки
    def check_headings(self):
        prev = [0, 0, 0]
        app_prev: dict = {}
        for it in self.items:
            if it.kind != "section":
                continue
            par = it.par
            num, dot = it.meta["num"], it.meta["dot"]
            letter = num[0] if num[0].isalpha() else None
            parts = [int(x) for x in (num[2:] if letter else num).split(".")]
            lvl = len(parts) + (1 if letter else 0)
            if dot:
                self.add(ERR, "H10", "после номера раздела точка не ставится", it)
            title = clean(it.text)[len(num) + len(dot):].strip()
            if title.endswith("."):
                self.add(ERR, "H11", "в конце заголовка точка не ставится", it)
            bold = self._any_bold(par)
            if lvl in (1, 2) and not bold:
                self.add(ERR, "H12", "заголовки разделов и подразделов – жирным", it)
            if lvl == 3 and bold:
                self.add(ERR, "H12", "заголовок пункта – без жирного", it)
            fi = self.pr.par(par, "first_line_indent")
            if fi is None or abs(fi.cm - C.INDENT_CM) > 0.06:
                self.add(ERR, "H13", "у заголовка раздела/подраздела/пункта должен быть абзацный отступ", it)
            al = self.pr.par(par, "alignment")
            if al is not None and "CENTER" in str(al):
                self.add(ERR, "H14", "заголовок раздела – по ширине или по левому краю, не по центру", it)
            # нумерация
            exp = None
            if letter:                                   # подразделы приложения: А.1, А.1.1
                pa = app_prev.setdefault(letter, [0, 0])
                exp = [pa[0] + 1] if lvl == 2 else [pa[0], pa[1] + 1]
                if parts != exp:
                    self.add(ERR, "H15", f"нарушена нумерация: «{num}», ожидалось «{letter}.{'.'.join(map(str, exp))}»", it)
                app_prev[letter] = parts + [0] * (2 - len(parts))
                self._check_gap(it, "заголовком")
                continue
            if lvl == 1:
                exp = [prev[0] + 1]
            elif lvl == 2:
                exp = [prev[0], prev[1] + 1]
            else:
                exp = [prev[0], prev[1], prev[2] + 1]
            if parts != exp:
                self.add(ERR, "H15", f"нарушена нумерация: «{num}», ожидалось «{'.'.join(map(str, exp))}»", it)
            prev = (parts + [0, 0, 0])[:3] if lvl == 3 else ([parts[0], 0, 0] if lvl == 1 else [parts[0], parts[1], 0])
            self._check_gap(it, "заголовком")

    def _gap_before(self, it: Item) -> bool:
        sb = self.pr.par(it.par, "space_before")
        if sb is not None and sb.pt >= 8:
            return True
        k = it.idx - 2
        return k >= 0 and self.items[k].kind == "empty" or self.page_break_before(it)

    def _gap_after(self, it: Item) -> bool:
        k = it.idx
        if k >= len(self.items):
            return True
        nxt = self.items[k]
        sb = self.pr.par(nxt.par, "space_before")
        sa = self.pr.par(it.par, "space_after")
        return (nxt.kind == "empty" or (sb is not None and sb.pt >= 8) or (sa is not None and sa.pt >= 8)
                or (not nxt.in_table and self.page_break_before(nxt)))

    def _check_gap(self, it: Item, what: str):
        if not self._gap_before(it):
            self.add(WARN, "H16", "заголовок должен отделяться от текста пустой строкой сверху", it)
        if not self._gap_after(it):
            self.add(WARN, "H16", "заголовок должен отделяться от текста пустой строкой снизу", it)

    # ---------------------------------------------------------- подписи
    def check_captions(self):
        body_idx_text = [(i.idx, clean(i.text).lower()) for i in self.items if i.kind in ("body", "list", "cell", "bib")]
        figs, tabs = [], []
        for it in self.items:
            t = clean(it.text)
            if it.kind == "figcap":
                m = RE_FIG.match(t)
                label, dash, title = m.group(1), m.group(2), m.group(3)
                figs.append(label)
                if dash != "–":
                    self.add(ERR, "C10", "в подписи рисунка должно быть короткое тире: «Рисунок N – Название»", it)
                if title.endswith("."):
                    self.add(ERR, "C11", "в конце подписи рисунка точка не ставится", it)
                if title and title[0].islower():
                    self.add(WARN, "C12", "название рисунка начинается с большой буквы", it)
                prev = self.items[it.idx - 2] if it.idx >= 2 else None
                if prev is None or prev.kind != "image":
                    self.add(WARN, "C13", "над подписью рисунка не найден сам рисунок", it)
                self._need_ref_before(it, label, ("рисун",), body_idx_text, "рисунок")
                self._check_gap_after_only(it, "подписью рисунка")
            elif it.kind == "tabcap":
                m = RE_TAB.match(t)
                label, dash, title = m.group(1), m.group(2), m.group(3)
                tabs.append(label)
                if dash != "–":
                    self.add(ERR, "C14", "в названии таблицы должно быть короткое тире: «Таблица N – Название»", it)
                if title.endswith("."):
                    self.add(ERR, "C15", "в конце названия таблицы точка не ставится", it)
                nxt = self.items[it.idx] if it.idx < len(self.items) else None
                if nxt is None or not nxt.in_table:
                    self.add(WARN, "C16", "после названия таблицы не найдена сама таблица", it)
                self._need_ref_before(it, label, ("таблиц",), body_idx_text, "таблицу")
        for kind, labels in (("Рисунок", figs), ("Таблица", tabs)):
            main = [l for l in labels if "." not in l or not l[0].isalpha()]
            for k, l in enumerate([x for x in labels if x[0].isdigit() and "." not in x], 1):
                if l != str(k):
                    self.f.append(Finding(ERR, "C17", f"нарушена сквозная нумерация: {kind} {l} вместо {kind} {k}", "документ", ""))
                    break

    def _check_gap_after_only(self, it: Item, what: str):
        if not self._gap_after(it):
            self.add(WARN, "C18", "после подписи рисунка нужна пустая строка", it)

    def _need_ref_before(self, it: Item, label: str, stems, body_idx_text, what: str):
        pat = re.compile(r"(%s)\w*\s+%s(?![\d.]*\d)" % ("|".join(stems), re.escape(label.lower())))
        # допускаем «рисунках 1 и 2», «таблицы 1–3»
        found_before = any(i < it.idx and pat.search(t) for i, t in body_idx_text)
        if not found_before:
            later = any(i > it.idx and pat.search(t) for i, t in body_idx_text)
            msg = f"на {what} {label} нет ссылки в тексте ДО неё" + (" (ссылка есть, но после)" if later else "")
            self.add(ERR, "C19", msg, it)

    # ------------------------------------------------------------------ списки
    def check_lists(self):
        for it in self.items:
            if it.kind != "list":
                continue
            t = clean(it.text)
            nid = self.num_of(it.par)
            if nid is not None:
                fmt, txt = self.numbering.get(nid, (None, None))
                if fmt == "bullet" and txt not in ("–", C.DASH):
                    self.add(ERR, "LS1", f"маркер списка «{txt}»: допустимы тире «–», цифра со скобкой, русская буква со скобкой", it)
                if fmt in ("decimal", "russianLower", "lowerLetter", "upperLetter") and txt and txt.endswith(".") :
                    self.add(ERR, "LS2", "после номера в перечислении ставится скобка, не точка", it)
                if fmt in ("lowerLetter", "upperLetter", "upperRussian"):
                    self.add(ERR, "LS3", "буквы в маркерах – только маленькие русские", it)
            else:
                m = re.match(r"^(\S+)\s", t)
                mk = m.group(1) if m else ""
                if mk in ("-", "•", "·", "*", "—", "•"):
                    self.add(ERR, "LS1", f"маркер «{mk}» недопустим: используйте тире «–»", it)
                elif re.fullmatch(r"\d+\.", mk):
                    self.add(ERR, "LS2", "после номера в перечислении ставится скобка, не точка", it)
                elif re.fullmatch(r"[a-z]\)", mk):
                    self.add(ERR, "LS3", "буквы в маркерах – только маленькие русские", it)

    # --------------------------------------------------------------------- код
    def check_code(self):
        for it in self.items:
            if it.kind != "code":
                continue
            ls = self.pr.line_mult(it.par, C.CODE_SIZE)
            if abs(ls - 1.0) > 0.06:
                self.add(ERR, "K01", f"междустрочный интервал кода {ls:g}, нужен одинарный", it)
        # пустая строка до/после блока кода
        i = 0
        while i < len(self.items):
            if self.items[i].kind == "code":
                j = i
                while j + 1 < len(self.items) and self.items[j + 1].kind == "code":
                    j += 1
                first, last = self.items[i], self.items[j]
                if not self._gap_before(first):
                    self.add(WARN, "K04", "перед блоком кода нужна пустая строка", first)
                if not self._gap_after(last):
                    self.add(WARN, "K04", "после блока кода нужна пустая строка", last)
                i = j + 1
            else:
                i += 1

    # ---------------------------------------------------------------- формулы
    def check_formulas(self):
        nums = []
        for it in self.items:
            if it.kind != "eq":
                continue
            t = clean(it.text)
            m = re.search(r"\(((?:[А-Я]\.)?\d+(?:\.\d+)?)\)\s*$", t)
            if it.in_table:
                if m and not it.meta.get("has_math"):
                    nums.append(m.group(1))
                continue
            # формула отдельным абзацем (не в таблице)
            al = self.pr.par(it.par, "alignment")
            if "CENTER" not in str(al) and not it.par._p.findall(".//{%s}oMathPara" % M_NS):
                self.add(ERR, "M01", "формула в отдельной строке должна быть выровнена по центру", it)
            if m:
                nums.append(m.group(1))
            else:
                self.add(WARN, "M02", "формула в отдельной строке должна быть пронумерована", it)
        main = [n for n in nums if n.isdigit()]
        for k, n in enumerate(main, 1):
            if int(n) != k:
                self.f.append(Finding(ERR, "M03", f"нарушена нумерация формул: ({n}) вместо ({k})", "документ", ""))
                break

    # -------------------------------------------------------- список источников
    def check_bibliography(self):
        bib = [it for it in self.items if it.kind == "bib"]
        body = " ".join(clean(i.text) for i in self.items if i.kind in ("body", "list", "cell", "section"))
        cites = set()
        for m in re.finditer(r"\[(\d+(?:\s*[,–\-]\s*\d+|\s*,\s*[сc]\.[^\]]*)*)\]", body):
            inner = m.group(1)
            for part in re.split(r",", inner):
                part = part.strip()
                if re.fullmatch(r"\d+", part):
                    cites.add(int(part))
                elif re.fullmatch(r"\d+\s*[–\-]\s*\d+", part):
                    a, b = re.split(r"[–\-]", part)
                    cites.update(range(int(a), int(b) + 1))
        if not bib:
            if cites:
                self.f.append(Finding(ERR, "B01", "в тексте есть ссылки [n], но список источников пуст или не найден", "документ", ""))
            return
        n = len(bib)
        urls = 0
        keys = []
        for k, it in enumerate(bib, 1):
            t = clean(it.text)
            if not t.endswith("."):
                self.add(ERR, "B02", "библиографическое описание должно заканчиваться точкой", it)
            nid = self.num_of(it.par)
            if nid is None:
                m = re.match(r"^(\d+)([.)]?)\s+", t)
                if not m:
                    self.add(ERR, "B03", "источник не пронумерован (арабская цифра без точки и скобки)", it)
                elif m.group(2):
                    self.add(ERR, "B03", "после номера источника точка и скобка не ставятся", it)
                else:
                    t = t[m.end():]
            else:
                fmt, txt = self.numbering.get(nid, (None, None))
                if fmt != "decimal" or (txt and txt != "%1"):
                    self.add(ERR, "B03", f"нумерация источников – арабская цифра без точки и скобки (сейчас «{txt}»)", it)
            if "URL" in t or re.search(r"https?://", t):
                urls += 1
                if not re.search(r"\(дата обращения:\s*\d{2}\.\d{2}\.\d{4}\)", t):
                    self.add(ERR, "B04", "для электронного ресурса нужна «(дата обращения: дд.мм.гггг)»", it)
            if re.search(r"\[Электронный ресурс\]", t):
                self.add(WARN, "B05", "по ГОСТ Р 7.0.100-2018 пометка «[Электронный ресурс]» не обязательна", it)
            if not re.search(r"\b(19|20)\d{2}\b", t) and "URL" not in t:
                self.add(WARN, "B06", "в описании не найден год издания", it)
            keys.append(re.sub(r"^[«\"'\s]+", "", t).lower().replace("ё", "е"))
        def grp(s):
            return 0 if re.match(r"[а-я]", s[:1]) else (1 if re.match(r"[a-z]", s[:1]) else 2)
        for a, b in zip(keys, keys[1:]):
            if (grp(a), a) > (grp(b), b):
                self.f.append(Finding(WARN, "B07", "источники не в алфавитном порядке (сначала кириллица, затем латиница)", "список источников", ""))
                break
        if urls > 2 * n / 3 + 0.01:
            self.f.append(Finding(WARN, "B08", f"электронных ресурсов {urls} из {n}: рекомендуется не более 1/3–2/3 списка", "список источников", ""))
        bad = sorted(c for c in cites if c < 1 or c > n)
        if bad:
            self.f.append(Finding(ERR, "B09", f"ссылки на несуществующие источники: {bad} (в списке {n})", "документ", ""))
        unused = [k for k in range(1, n + 1) if k not in cites]
        if unused and cites:
            self.f.append(Finding(WARN, "B10", f"на источники {unused} нет ссылок в тексте: в списке должны быть только цитируемые", "список источников", ""))
        if not cites:
            self.f.append(Finding(WARN, "B10", "в тексте нет ссылок на источники вида [1]", "документ", ""))

    # -------------------------------------------------------------- содержание
    def check_toc(self):
        has_toc_unit = any(it.kind == "struct" and it.meta["name"] == C.TOC for it in self.items)
        toc_items = [it for it in self.items if it.kind == "toc"]
        xml = self.doc.element.body.xml
        if has_toc_unit:
            if not toc_items and "TOC" not in xml:
                self.f.append(Finding(ERR, "O01", "раздел «Содержание» пуст: нужно оглавление (автособираемое)", "содержание", ""))
            else:
                empty_pages = [it for it in toc_items if not re.search(r"\d\s*$", clean(it.text))]
                if empty_pages:
                    self.add(WARN, "O02", "в оглавлении не проставлены номера страниц: обновите поле (F9) или соберите отчёт через build_report.py", empty_pages[0])
                n_head = sum(1 for it in self.items if it.kind in ("section", "appendix") or (it.kind == "struct" and it.meta["name"] not in (C.ABSTRACT, C.TOC)))
                if toc_items and len(toc_items) < n_head:
                    self.f.append(Finding(WARN, "O03", f"в оглавлении {len(toc_items)} строк, заголовков в документе {n_head}", "содержание", ""))
        # курсив для латиницы
        for it in self.items:
            if it.kind not in ("body", "list", "cell", "figcap", "tabcap"):
                continue
            if clean(it.text).upper() == clean(it.text):      # строка ключевых слов / аббревиатуры ПРОПИСНЫМИ
                continue
            nonital = []
            for r in self._runs(it.par):
                if self.pr.run(r, it.par, "italic"):
                    continue
                if (self.pr.run(r, it.par, "name") or "").lower() in CODE_FONTS:
                    continue
                txt = RE_URL.sub(" ", r.text)
                for m in RE_LATIN.finditer(txt):
                    w = m.group(0)
                    if not RE_ROMAN.match(w):
                        nonital.append(w)
            if nonital:
                self.add(WARN, "T07", "английские слова должны быть курсивом: " + ", ".join(nonital[:4]), it)

    # ------------------------------------------------------- вёрстка (Word)
    def check_layout(self, probe: dict):
        tables = probe.get("tables", [])
        body_tables = [e for e in self.doc.element.body.iterchildren() if e.tag == qn("w:tbl")]
        data_caps = {}
        for it in self.items:
            if it.kind == "tabcap":
                nxt = self.items[it.idx] if it.idx < len(self.items) else None
                if nxt is not None and nxt.in_table:
                    data_caps[id(nxt.meta.get("table_el"))] = (it, RE_TAB.match(clean(it.text)).group(1))
        for k, el in enumerate(body_tables):
            if k >= len(tables):
                break
            pages = tables[k].get("rows") or []
            if len(set(pages)) > 1 and id(el) in data_caps:
                it, label = data_caps[id(el)]
                self.add(ERR, "C20", f"таблица {label} переходит на следующую страницу без «Продолжение таблицы {label}» "
                                    f"и повтора шапки: разбейте её (build_report.py делает это автоматически)", it)


# ======================================================================== CLI
def format_findings(findings: list[Finding], max_locs: int = 4) -> str:
    """Группирует одинаковые замечания: код + текст, показывает первые места и счётчик."""
    order = {ERR: 0, WARN: 1}
    groups: dict[tuple, list[Finding]] = {}
    for f in sorted(findings, key=lambda x: (order[x.sev], x.code)):
        groups.setdefault((f.sev, f.code, f.msg), []).append(f)
    out = []
    for (sev, code, msg), fs in groups.items():
        mark = "✗" if sev == ERR else "!"
        if len(fs) == 1:
            f = fs[0]
            loc = f" [{f.loc}]" if f.loc else ""
            snip = f" «{f.snippet}»" if f.snippet else ""
            out.append(f"  {mark} {code}{loc}: {msg}{snip}")
        else:
            locs = [f.loc.replace("абзац ", "") for f in fs if f.loc]
            shown = ", ".join(locs[:max_locs]) + (f" … (+{len(locs) - max_locs})" if len(locs) > max_locs else "")
            first = fs[0].snippet
            tail = f"; например «{first}»" if first else ""
            out.append(f"  {mark} {code} ×{len(fs)}: {msg}\n        абзацы: {shown}{tail}")
    return "\n".join(out)


def run(path: str, work: str = "lr", use_word: bool = False, quiet_ok: bool = True, strict: bool = False,
        font: str = C.FONT, line_spacing: float | None = None, as_json: bool = False) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ch = Checker(path, work=work, font=font, line_spacing=line_spacing)
    probe = None
    if use_word:
        try:
            from . import layout as LY
            probe = LY.probe(path)
        except Exception as e:      # noqa: BLE001
            print(f"  (проверка вёрстки через Word пропущена: {str(e)[:120]})")
    findings = ch.run_all(probe)
    errs = [f for f in findings if f.sev == ERR]
    warns = [f for f in findings if f.sev == WARN]
    if as_json:
        import json
        print(json.dumps([f.__dict__ for f in findings], ensure_ascii=False, indent=2))
    else:
        if findings:
            print(format_findings(findings))
        print(f"\nИтог: ошибок {len(errs)}, предупреждений {len(warns)}"
              + ("  – оформление соответствует правилам ✓" if not findings else ""))
    return 1 if errs or (strict and warns) else 0
