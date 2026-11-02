"""Сборка .docx отчёта из разобранного Markdown. Все правила оформления зашиты здесь."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, Twips

from . import bib
from . import config as C
from . import titlepage
from .inline import Span, Typo, normalize_spans, parse_inline
from .mdparse import Block
from .omml import LatexError, latex_to_omml_xml


class BuildError(Exception):
    pass


@dataclass
class HeadingInfo:
    block: int
    cls: str                # struct|section|appendix
    level: int
    number: str             # '1', '1.2', 'А'
    title: str              # текст без номера
    full: str               # как показывается в оглавлении
    bookmark: str
    name: str = ""          # каноничное имя структурного элемента


@dataclass
class BuildResult:
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    headings: list = field(default_factory=list)
    tables: list = field(default_factory=list)       # [{tid, rows}]


def _set_rfonts(rpr, name: str):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in list(rf.attrib):
        del rf.attrib[a]
    for a in ("ascii", "hAnsi", "cs", "eastAsia"):
        rf.set(qn(f"w:{a}"), name)


def _spacing_line(blank_factor: float) -> int:
    return int(round(blank_factor * 240))


class Builder:
    def __init__(self, meta: dict, blocks: list[Block], sources: list[dict], base_dir: str,
                 layout: dict | None = None, table_breaks: dict | None = None):
        self.meta = meta
        self.blocks = blocks
        self.sources = {s["id"]: bib.normalize_entry(s) for s in sources if "id" in s}
        self.base_dir = base_dir
        self.layout = layout or {}
        self.table_breaks = table_breaks or {}
        self.res = BuildResult()
        self.typo = Typo()
        self.auto_italic = bool(meta.get("auto_italic", True))
        work = str(meta.get("work", meta.get("type", "lr"))).lower()
        if work not in C.PROFILES:
            self.err(0, f"неизвестный вид работы work: {work}. Допустимо: {', '.join(C.PROFILES)}")
            work = "lr"
        self.profile = C.PROFILES[work]
        self.line_spacing = float(meta.get("line_spacing", self.profile["line_spacing"]))
        self.blank_pt = round(C.LINE_PT * self.line_spacing)
        self.section_new_page = bool(meta.get("section_new_page", False))
        self.span_cache: dict[str, list[Span]] = {}
        self.after_gap = False
        self._bm_id = 100
        self._num_id = 100
        self._table_counter = 0

    # --------------------------------------------------------------- сообщения
    def err(self, line, msg):
        self.res.errors.append(f"строка {line}: {msg}" if line else msg)

    def warn(self, msg):
        self.res.warnings.append(msg)

    # ----------------------------------------------------------- парсинг текста
    def spans(self, text: str, line: int = 0) -> list[Span]:
        key = text
        if key not in self.span_cache:
            try:
                self.span_cache[key] = parse_inline(text)
            except ValueError as e:
                self.err(line, str(e))
                self.span_cache[key] = [Span("text", text)]
        return self.span_cache[key]

    def texts_of(self, b: Block) -> list[str]:
        d = b.data
        if b.kind == "para":
            return [d["text"]]
        if b.kind == "list":
            return list(d["items"])
        if b.kind == "figure":
            return [d["caption"]]
        if b.kind == "table":
            return [d["caption"]] + d["header"] + [c for r in d["rows"] for c in r]
        if b.kind == "heading":
            return [d["text"]]
        return []

    # ============================================================== ПРЕДПРОХОД
    def prepass(self):
        blocks = self.blocks
        self.ids: dict[tuple, str] = {}
        self.def_idx: dict[tuple, int] = {}
        self.first_ref: dict[tuple, int] = {}
        self.cite_order: list[str] = []
        self.heads: dict[int, HeadingInfo] = {}
        counters = [0, 0, 0]
        app_idx = -1
        cur_app: str | None = None
        n_fig = n_tab = n_eq = 0
        seen_struct: set[str] = set()
        app_ids: dict[str, int] = {}
        bm = 0

        for bi, b in enumerate(blocks):
            if b.kind == "heading":
                d = b.data
                bm += 1
                bookmark = f"_Toc{900000 + bm}"
                if d["cls"] == "struct":
                    name = d["name"]
                    if name in seen_struct:
                        self.err(b.line, f"структурный элемент «{name}» встречается дважды")
                    seen_struct.add(name)
                    cur_app = None
                    self.heads[bi] = HeadingInfo(bi, "struct", 1, "", name, name, bookmark, name)
                elif d["cls"] == "appendix":
                    app_idx += 1
                    if app_idx >= len(C.APPENDIX_LETTERS):
                        self.err(b.line, "слишком много приложений")
                        app_idx = len(C.APPENDIX_LETTERS) - 1
                    cur_app = C.APPENDIX_LETTERS[app_idx]
                    n_fig = n_tab = n_eq = 0
                    counters = [0, 0, 0]
                    title = d["text"].strip().rstrip(".")
                    self.heads[bi] = HeadingInfo(bi, "appendix", 1, cur_app, title,
                                                 f"{C.APPENDIX} {cur_app} {title}", bookmark, C.APPENDIX)
                    if d.get("id"):
                        self._define(("app", d["id"]), cur_app, bi, b.line)
                        app_ids[d["id"]] = bi
                else:
                    lvl = d["level"]
                    if cur_app:
                        # подразделы внутри приложения: А.1, А.1.1 (заголовок 1 уровня недопустим – это приложение)
                        if lvl == 1:
                            self.err(b.line, "внутри приложения заголовок первого уровня недопустим: используйте «##» для подразделов")
                            continue
                        idx = lvl - 2
                        counters[idx] += 1
                        for k in range(idx + 1, 3):
                            counters[k] = 0
                        num = cur_app + "." + ".".join(str(x) for x in counters[:idx + 1])
                        hl = lvl - 1
                    else:
                        idx = lvl - 1
                        if idx > 0 and counters[idx - 1] == 0:
                            self.err(b.line, f"заголовок «{d['text']}» уровня {lvl} без заголовка уровня {lvl - 1} выше")
                        counters[idx] += 1
                        for k in range(idx + 1, 3):
                            counters[k] = 0
                        num = ".".join(str(x) for x in counters[:idx + 1])
                        hl = lvl
                    title = d["text"].strip().rstrip(".")
                    self.heads[bi] = HeadingInfo(bi, "section", hl, num, title, f"{num} {title}", bookmark)
                    if d.get("id"):
                        self._define(("sec", d["id"]), num, bi, b.line)
                        # номера разделов доступны как [[sec:id]]
            elif b.kind == "figure":
                n_fig += 1
                lab = f"{cur_app}.{n_fig}" if cur_app else str(n_fig)
                b.data["label"] = lab
                if b.data.get("id"):
                    self._define(("fig", b.data["id"]), lab, bi, b.line)
                else:
                    self.err(b.line, "у рисунка нет метки {#fig-имя}: на рисунок обязательно должна быть ссылка [[fig-имя]] в тексте до него")
            elif b.kind == "table":
                n_tab += 1
                lab = f"{cur_app}.{n_tab}" if cur_app else str(n_tab)
                b.data["label"] = lab
                if b.data.get("id"):
                    self._define(("tab", b.data["id"]), lab, bi, b.line)
                else:
                    self.err(b.line, "у таблицы нет метки {#tab-имя}: на таблицу обязательно должна быть ссылка [[tab-имя]] в тексте до неё")
            elif b.kind == "math":
                n_eq += 1
                lab = f"{cur_app}.{n_eq}" if cur_app else str(n_eq)
                b.data["label"] = lab
                if b.data.get("id"):
                    self._define(("eq", b.data["id"]), lab, bi, b.line)

            # ссылки и цитаты в тексте
            for t in self.texts_of(b):
                for sp in self.spans(t, b.line):
                    if sp.kind == "cite":
                        for sid, _extra in sp.ids:
                            if sid not in self.sources:
                                self.err(b.line, f"источник «{sid}» не найден в sources.json")
                            elif sid not in self.cite_order:
                                self.cite_order.append(sid)
                    elif sp.kind == "ref":
                        self.first_ref.setdefault(sp.ref, bi)
                    elif sp.kind == "math":
                        try:
                            latex_to_omml_xml(sp.text)
                        except LatexError as e:
                            self.err(b.line, f"формула ${sp.text}$: {e}")
            if b.kind == "math":
                try:
                    latex_to_omml_xml(b.data["latex"])
                except LatexError as e:
                    self.err(b.line, f"формула «{b.data['latex'][:50]}»: {e}")
            elif b.kind == "figure":
                try:
                    self._find_image(b.data["path"])
                except BuildError as e:
                    self.err(b.line, str(e))

        # проверка ссылок
        for key, bi in self.first_ref.items():
            kind, i = key
            if key not in self.ids:
                self.err(self.blocks[bi].line, f"ссылка [[{i}]] ведёт в никуда: нигде нет метки {{#{i}}}")
        for key, di in self.def_idx.items():
            kind, i = key
            if kind in ("fig", "tab"):
                ri = self.first_ref.get(key)
                if ri is None:
                    self.err(self.blocks[di].line, f"на {'рисунок' if kind == 'fig' else 'таблицу'} «{i}» нет ссылки в тексте "
                                                   f"(добавьте [[{i}]] в абзац ДО неё)")
                elif ri >= di:
                    self.err(self.blocks[di].line, f"ссылка на {'рисунок' if kind == 'fig' else 'таблицу'} «{i}» стоит после неё: "
                                                   f"по правилам ссылка должна быть в тексте ДО рисунка/таблицы")
        # приложения: обязательное/справочное
        for bi, h in self.heads.items():
            if h.cls == "appendix":
                d = self.blocks[bi].data
                if not d.get("mode"):
                    d["mode"] = "обязательное" if (d.get("id") and ("app", d["id"]) in self.first_ref) else "справочное"

        # нумерация источников
        used = [self.sources[s] for s in self.cite_order if s in self.sources]
        self.sorted_sources = bib.sort_sources(used)
        self.src_num = {s["id"]: i + 1 for i, s in enumerate(self.sorted_sources)}
        unused = [sid for sid in self.sources if sid not in self.cite_order]
        if unused:
            self.warn("не процитированы и не попадут в список: " + ", ".join(unused))

        # статистика
        self.res.stats = dict(
            figures=sum(1 for b in blocks if b.kind == "figure"),
            tables=sum(1 for b in blocks if b.kind == "table"),
            equations=sum(1 for b in blocks if b.kind == "math"),
            sources=len(self.sorted_sources),
            appendices=sum(1 for h in self.heads.values() if h.cls == "appendix"),
        )

        # титульный лист
        self.title_data = None
        if titlepage.enabled(self.meta):
            data, missing = titlepage.resolve(self.meta, self.profile["title"])
            if missing:
                self.err(0, "для титульного листа не хватает полей в front matter: " + "; ".join(missing)
                         + ". Значения не угадываются – запросите их у пользователя (или title_page: none, чтобы отказаться от титульника)")
            else:
                self.title_data = data

        # структура
        names = [h.name for h in self.heads.values() if h.cls == "struct"]
        for req in self.profile["required"]:
            if req not in names:
                self.err(0, f"для вида работы «{self.profile['title']}» обязателен раздел «{req}»")
        for rec in self.profile.get("recommended", []):
            if rec not in names:
                self.warn(f"рекомендуется добавить раздел «{rec}»")
        if not any(h.cls == "section" and h.level == 1 for h in self.heads.values()):
            self.err(0, "нет основной части: добавьте разделы «# Название раздела»")
        if C.ABSTRACT in names and not self.meta.get("keywords"):
            self.err(0, "для РЕФЕРАТА нужно поле keywords: в front matter (ключевые слова через запятую)")

        # порядок структурных элементов
        order = [n for n in C.STRUCT_ORDER]
        pos = []
        main_seen = False
        for bi, h in sorted(self.heads.items()):
            if h.cls == "struct":
                pos.append((order.index(h.name), h.name, self.blocks[bi].line))
            elif h.cls == "section" and not main_seen:
                main_seen = True
                pos.append((order.index("ОСНОВНАЯ ЧАСТЬ"), "основная часть", self.blocks[bi].line))
            elif h.cls == "appendix":
                pos.append((order.index(C.APPENDIX), C.APPENDIX, self.blocks[bi].line))
        for a, b in zip(pos, pos[1:]):
            if b[0] < a[0]:
                self.err(b[2], f"нарушен порядок: «{b[1]}» должен стоять раньше «{a[1]}»")
                break

    def _define(self, key, label, bi, line):
        if not key[1].startswith(key[0] + "-"):
            self.err(line, f"идентификатор «{key[1]}» должен начинаться с «{key[0]}-» (например {{#{key[0]}-имя}})")
        if key in self.ids:
            self.err(line, f"идентификатор {key[1]} ({key[0]}) уже использован")
        self.ids[key] = label
        self.def_idx[key] = bi

    # ============================================================== ДОКУМЕНТ
    def build(self, out_path: str) -> BuildResult:
        self.prepass()
        if self.res.errors:
            return self.res
        self.doc = Document()
        self._setup_defaults()
        self._setup_styles()
        self._setup_numbering()
        self._setup_page()
        if self.title_data:
            titlepage.build(self, self.title_data)
            self.after_gap = False
        self._emit_all()
        if self.res.errors:
            return self.res
        self._finish(out_path)
        for what, n in self.typo.fixes.items():
            self.res.notes.append(f"автоисправлено: {what} ({n})")
        self.res.headings = [h for h in self.heads.values()]
        return self.res

    # ------------------------------------------------------------ стили / настройки
    def _setup_defaults(self):
        styles_el = self.doc.styles.element
        dd = styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            rpr = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr"))
            if rpr is not None:
                _set_rfonts(rpr, C.FONT)
                for tag in ("w:sz", "w:szCs", "w:lang"):
                    e = rpr.find(qn(tag))
                    if e is not None:
                        rpr.remove(e)
                sz = OxmlElement("w:sz")
                sz.set(qn("w:val"), str(C.FONT_SIZE * 2))
                szcs = OxmlElement("w:szCs")
                szcs.set(qn("w:val"), str(C.FONT_SIZE * 2))
                lang = OxmlElement("w:lang")
                lang.set(qn("w:val"), "ru-RU")
                lang.set(qn("w:eastAsia"), "ru-RU")
                lang.set(qn("w:bidi"), "ar-SA")
                rpr.extend([sz, szcs, lang])
            ppr = dd.find(qn("w:pPrDefault") + "/" + qn("w:pPr"))
            if ppr is not None:
                for ch in list(ppr):
                    ppr.remove(ch)

    def _reset(self, style):
        el = style.element
        for tag in ("w:pPr", "w:rPr"):
            e = el.find(qn(tag))
            if e is not None:
                el.remove(e)

    def _mk_style(self, name: str, sid: str, base: str | None = "Normal", font_size=None, bold=None, italic=None,
                  align=None, first_indent=None, left_indent=None, line=None, keep_next=None, page_break=None,
                  outline: int | None = None, font: str | None = None, builtin: str | None = None):
        styles = self.doc.styles
        try:
            st = styles[builtin or name]
            self._reset(st)
        except KeyError:
            st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.element.set(qn("w:styleId"), sid)
        if base:
            st.base_style = styles[base]
        st.quick_style = True
        rpr = st.element.get_or_add_rPr()
        if font:
            _set_rfonts(rpr, font)
        f = st.font
        if font_size:
            f.size = Pt(font_size)
        if bold is not None:
            f.bold = bold
        if italic is not None:
            f.italic = italic
        pf = st.paragraph_format
        if align is not None:
            pf.alignment = align
        if first_indent is not None:
            pf.first_line_indent = Cm(first_indent)
        if left_indent is not None:
            pf.left_indent = Cm(left_indent)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        if line is not None:
            pf.line_spacing = line
        if keep_next is not None:
            pf.keep_with_next = keep_next
        if page_break is not None:
            pf.page_break_before = page_break
        pf.widow_control = True
        if outline is not None:
            ppr = st.element.get_or_add_pPr()
            ol = OxmlElement("w:outlineLvl")
            ol.set(qn("w:val"), str(outline))
            ppr.append(ol)
        return st

    def _setup_styles(self):
        J, Ctr, L = WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT
        ls = self.line_spacing
        normal = self.doc.styles["Normal"]
        self._reset(normal)
        rpr = normal.element.get_or_add_rPr()
        _set_rfonts(rpr, C.FONT)
        normal.font.size = Pt(C.FONT_SIZE)
        pf = normal.paragraph_format
        pf.alignment = J
        pf.first_line_indent = Cm(C.INDENT_CM)
        pf.left_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = ls
        pf.widow_control = True

        mk = self._mk_style
        mk("Heading 1", "Heading1", "Normal", C.FONT_SIZE, True, False, J, C.INDENT_CM, None, ls, True, None, 0, C.FONT, builtin="Heading 1")
        mk("Heading 2", "Heading2", "Normal", C.FONT_SIZE, True, False, J, C.INDENT_CM, None, ls, True, None, 1, C.FONT, builtin="Heading 2")
        mk("Heading 3", "Heading3", "Normal", C.FONT_SIZE, False, False, J, C.INDENT_CM, None, ls, True, None, 2, C.FONT, builtin="Heading 3")
        mk("Структурный заголовок", "GostStruct", "Normal", C.FONT_SIZE, False, False, Ctr, 0, None, ls, True, True, 0, C.FONT)
        mk("Заголовок вне оглавления", "GostStructNoToc", "Normal", C.FONT_SIZE, False, False, Ctr, 0, None, ls, True, True, None, C.FONT)
        mk("Подпись рисунка", "GostFigCaption", "Normal", C.FONT_SIZE, False, False, Ctr, 0, None, ls, False, None, None, C.FONT)
        mk("Подпись таблицы", "GostTabCaption", "Normal", C.FONT_SIZE, False, False, L, 0, None, ls, True, None, None, C.FONT)
        mk("Рисунок", "GostFigure", "Normal", C.FONT_SIZE, False, False, Ctr, 0, None, 1.0, True, None, None, C.FONT)
        mk("Код", "GostCode", "Normal", C.CODE_SIZE, False, False, L, 0, None, 1.0, False, None, None, C.CODE_FONT)
        mk("Текст таблицы", "GostTableText", "Normal", C.TABLE_FONT_SIZE, False, False, L, 0, None, 1.0, False, None, None, C.FONT)
        mk("Формула", "GostFormula", "Normal", C.FONT_SIZE, False, False, Ctr, 0, None, 1.0, False, None, None, C.FONT)
        mk("Список", "GostList", "Normal", C.FONT_SIZE, False, False, J, None, None, ls, False, None, None, C.FONT)
        for lvl, left in ((1, 0.0), (2, 0.5), (3, 1.0)):
            st = mk(f"toc {lvl}", f"TOC{lvl}", "Normal", C.FONT_SIZE, False, False, L, 0, left, ls, False, None, None, C.FONT)
            st.paragraph_format.tab_stops.add_tab_stop(Cm(C.TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        foot = self.doc.styles["Footer"]
        self._reset(foot)
        foot.base_style = self.doc.styles["Normal"]
        foot.paragraph_format.alignment = Ctr
        foot.paragraph_format.first_line_indent = Cm(0)
        foot.paragraph_format.line_spacing = 1.0
        _set_rfonts(foot.element.get_or_add_rPr(), C.FONT)
        foot.font.size = Pt(C.FONT_SIZE)

    def _setup_numbering(self):
        numbering = self.doc.part.numbering_part.numbering_definitions._numbering
        ns = nsdecls("w")

        def lvl(fmt, text, font=True):
            rpr = (f'<w:rPr><w:rFonts w:ascii="{C.FONT}" w:hAnsi="{C.FONT}" w:cs="{C.FONT}" w:hint="default"/></w:rPr>' if font else "")
            return (f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:suff w:val="space"/>'
                    f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr><w:ind w:left="0" w:firstLine="{int(C.INDENT_CM * 567 + 0.5)}"/></w:pPr>{rpr}</w:lvl>')

        defs = {90: lvl("bullet", C.DASH), 91: lvl("decimal", "%1)"), 92: lvl("russianLower", "%1)"), 93: lvl("decimal", "%1")}
        first_num = numbering.find(qn("w:num"))
        for aid, body in defs.items():
            el = parse_xml(f'<w:abstractNum {ns} w:abstractNumId="{aid}"><w:multiLevelType w:val="hybridMultilevel"/>{body}</w:abstractNum>')
            if first_num is not None:
                first_num.addprevious(el)
            else:
                numbering.append(el)
        self.num_dash = self._new_num(90, restart=False)

    def _new_num(self, abstract: int, restart=True) -> int:
        numbering = self.doc.part.numbering_part.numbering_definitions._numbering
        self._num_id += 1
        nid = self._num_id
        ov = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' if restart else ""
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{abstract}"/>{ov}</w:num>'))
        return nid

    def _setup_page(self):
        s = self.doc.sections[0]
        s.page_width, s.page_height = Cm(C.PAGE_W_CM), Cm(C.PAGE_H_CM)
        s.top_margin, s.bottom_margin = Cm(C.MARGIN_TOP), Cm(C.MARGIN_BOTTOM)
        s.left_margin, s.right_margin = Cm(C.MARGIN_LEFT), Cm(C.MARGIN_RIGHT)
        s.header_distance, s.footer_distance = Cm(1.0), Cm(0.9)
        fp = s.footer.paragraphs[0]
        fp.style = self.doc.styles["Footer"]
        self._add_field(fp, " PAGE ", "1")
        cp = self.doc.core_properties
        cp.title = str(self.meta.get("title", ""))
        cp.author = str(self.meta.get("author", ""))
        cp.language = "ru-RU"
        cp.comments = "Оформлено по ГОСТ 7.32-2017 / правилам оформления отчётов (skill ghost-report)"

    # ---------------------------------------------------------------- низкоуровневые
    def _add_field(self, par, instr: str, cached: str, fmt=None):
        def r(child):
            run = OxmlElement("w:r")
            run.append(child)
            return run

        def fc(t):
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), t)
            return e

        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        t = OxmlElement("w:t")
        t.text = cached
        for el in (r(fc("begin")), r(it), r(fc("separate")), r(t), r(fc("end"))):
            par._p.append(el)

    def _bookmark(self, par, name: str):
        self._bm_id += 1
        s = OxmlElement("w:bookmarkStart")
        s.set(qn("w:id"), str(self._bm_id))
        s.set(qn("w:name"), name)
        e = OxmlElement("w:bookmarkEnd")
        e.set(qn("w:id"), str(self._bm_id))
        ppr = par._p.find(qn("w:pPr"))
        if ppr is not None:
            ppr.addnext(s)
        else:
            par._p.insert(0, s)
        par._p.append(e)

    def _gap(self, par, wants_before: bool):
        need = self.after_gap or wants_before
        par.paragraph_format.space_before = Pt(self.blank_pt if need else 0)

    def _runs(self, par, spans: list[Span], size: int | None = None, math_size: int | None = None, line: int = 0):
        for sp in spans:
            if sp.kind == "text":
                if not sp.text:
                    continue
                run = par.add_run(sp.text)
                run.bold = True if sp.bold else None
                run.italic = True if sp.italic else None
                if size:
                    run.font.size = Pt(size)
            elif sp.kind == "code":
                run = par.add_run(sp.text)
                rpr = run._r.get_or_add_rPr()
                _set_rfonts(rpr, C.CODE_FONT)
                run.font.size = Pt(C.CODE_SIZE)
            elif sp.kind == "math":
                try:
                    xml = latex_to_omml_xml(sp.text, math_size or size or C.FONT_SIZE, C.MATH_FONT)
                    par._p.append(parse_xml(xml))
                except LatexError as e:
                    self.err(line, f"формула ${sp.text}$: {e}")
            elif sp.kind == "cite":
                parts = []
                for sid, extra in sorted(sp.ids, key=lambda x: self.src_num.get(x[0], 0)):
                    n = self.src_num.get(sid)
                    parts.append(f"{n}, {extra}" if extra else str(n))
                run = par.add_run("[" + ", ".join(parts) + "]")
                if size:
                    run.font.size = Pt(size)
            elif sp.kind == "ref":
                kind, i = sp.ref
                lab = self.ids.get((kind, i), "?")
                run = par.add_run(lab)
                if size:
                    run.font.size = Pt(size)
            elif sp.kind == "br":
                par.add_run().add_break(WD_BREAK.LINE)

    def _text_par(self, text: str, line: int, style: str = "Normal", wants_before=False, wants_after=False,
                  auto_italic: bool | None = None, strip_period=False):
        par = self.doc.add_paragraph(style=style)
        spans = normalize_spans(list(self.spans(text, line)), self.typo,
                                self.auto_italic if auto_italic is None else auto_italic)
        if strip_period and spans and spans[-1].kind == "text":
            spans[-1] = Span("text", spans[-1].text.rstrip().rstrip("."), spans[-1].bold, spans[-1].italic)
        self._gap(par, wants_before)
        self._runs(par, spans, line=line)
        self.after_gap = wants_after
        return par

    # ==================================================================== ВЫВОД
    def _emit_all(self):
        self.last_top = None          # struct|section|appendix
        self.in_appendix = False
        self._cur_struct = None
        for bi, b in enumerate(self.blocks):
            self.bi = bi
            try:
                getattr(self, f"_emit_{b.kind}")(b)
            except (BuildError, OSError) as e:
                self.err(b.line, str(e))

    # ------------------------------------------------------------- заголовки
    def _emit_heading(self, b: Block):
        h = self.heads.get(self.bi)
        if h is None:
            return
        d = b.data
        if h.cls == "struct":
            self._cur_struct = h.name
            self.in_appendix = False
            style = "Заголовок вне оглавления" if h.name in (C.ABSTRACT, C.TOC) else "Структурный заголовок"
            par = self.doc.add_paragraph(style=style)
            par.add_run(h.name)
            par.paragraph_format.space_before = Pt(0)
            self._bookmark(par, h.bookmark)
            self.after_gap = True
            self.last_top = "struct"
            if h.name == C.TOC:
                self._emit_toc()
                self.after_gap = False
            elif h.name == C.ABSTRACT:
                self._emit_abstract_head()
            elif h.name == C.BIB:
                self._emit_bibliography()
        elif h.cls == "appendix":
            self._cur_struct = C.APPENDIX
            self.in_appendix = True
            par = self.doc.add_paragraph(style="Структурный заголовок")
            par.paragraph_format.space_before = Pt(0)
            par.add_run(f"{C.APPENDIX} {h.number}")
            par.add_run().add_break(WD_BREAK.LINE)
            par.add_run(f"({d['mode']})")
            par.add_run().add_break(WD_BREAK.LINE)
            title = normalize_spans(list(self.spans(h.title, b.line)), self.typo, False)
            self._runs(par, title, line=b.line)
            self._bookmark(par, h.bookmark)
            self.after_gap = True
            self.last_top = "appendix"
        else:
            self._cur_struct = "ОСНОВНАЯ ЧАСТЬ" if not self.in_appendix else C.APPENDIX
            style = f"Heading {min(h.level, 3)}"
            par = self.doc.add_paragraph(style=style)
            title = normalize_spans(list(self.spans(h.title, b.line)), self.typo, False)
            par.add_run(f"{h.number} ")
            self._runs(par, title, line=b.line)
            new_page = (h.level == 1 and not self.in_appendix and (self.last_top != "section" or self.section_new_page))
            if new_page:
                par.paragraph_format.page_break_before = True
                par.paragraph_format.space_before = Pt(0)
            else:
                self._gap(par, True)
            self._bookmark(par, h.bookmark)
            self.after_gap = True
            if not self.in_appendix:
                self.last_top = "section"

    # ------------------------------------------------------ реферат / оглавление / список
    def _emit_abstract_head(self):
        s = self.res.stats
        pages = self.layout.get("pages")
        par = self.doc.add_paragraph(style="Normal")
        self._gap(par, False)
        par.add_run("Отчет ")
        self._add_field(par, " NUMPAGES ", str(pages if pages else 1))
        tail = f" с."
        if s["figures"]:
            tail += f", {s['figures']} рис."
        if s["tables"]:
            tail += f", {s['tables']} табл."
        if s["sources"]:
            tail += f", {s['sources']} источн."
        if s["appendices"]:
            tail += f", {s['appendices']} прил."
        par.add_run(tail)
        self.after_gap = False
        kw = [k.strip().upper() for k in re.split(r"[,;]", str(self.meta.get("keywords", ""))) if k.strip()]
        if kw:
            self._gap_par(", ".join(kw))

    def _gap_par(self, text: str):
        par = self.doc.add_paragraph(style="Normal")
        self._gap(par, False)
        par.add_run(text)
        self.after_gap = False

    def _emit_toc(self):
        entries = [h for h in self.heads.values()
                   if h.cls in ("section", "appendix") or (h.cls == "struct" and h.name not in (C.ABSTRACT, C.TOC))]
        entries.sort(key=lambda h: h.block)
        pages = self.layout.get("bookmarks", {})
        if not entries:
            return
        if not pages:
            self.res.notes.append("номера страниц в оглавлении появятся после обновления полей (F9) или при расчёте через Word")
        for k, h in enumerate(entries):
            lvl = 1 if h.cls in ("struct", "appendix") else min(h.level, 3)
            par = self.doc.add_paragraph(style=f"toc {lvl}")
            if k == 0:
                par.paragraph_format.space_before = Pt(self.blank_pt)
                for t, content in (("begin", None), ("instr", ' TOC \\o "1-3" \\h \\z \\u '), ("separate", None)):
                    run = OxmlElement("w:r")
                    if t == "instr":
                        it = OxmlElement("w:instrText")
                        it.set(qn("xml:space"), "preserve")
                        it.text = content
                        run.append(it)
                    else:
                        fc = OxmlElement("w:fldChar")
                        fc.set(qn("w:fldCharType"), t)
                        run.append(fc)
                    par._p.append(run)
            pg = str(pages.get(h.bookmark, ""))
            link = OxmlElement("w:hyperlink")
            link.set(qn("w:anchor"), h.bookmark)
            link.set(qn("w:history"), "1")
            r1 = OxmlElement("w:r")
            t1 = OxmlElement("w:t")
            t1.set(qn("xml:space"), "preserve")
            t1.text = h.full
            r1.append(t1)
            r2 = OxmlElement("w:r")
            r2.append(OxmlElement("w:tab"))
            link.extend([r1, r2])
            for typ, text in (("begin", None), ("instr", f" PAGEREF {h.bookmark} \\h "), ("separate", None), ("text", pg), ("end", None)):
                run = OxmlElement("w:r")
                if typ == "instr":
                    it = OxmlElement("w:instrText")
                    it.set(qn("xml:space"), "preserve")
                    it.text = text
                    run.append(it)
                elif typ == "text":
                    tt = OxmlElement("w:t")
                    tt.text = text
                    run.append(tt)
                else:
                    fc = OxmlElement("w:fldChar")
                    fc.set(qn("w:fldCharType"), typ)
                    run.append(fc)
                link.append(run)
            par._p.append(link)
            if k == len(entries) - 1:
                run = OxmlElement("w:r")
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), "end")
                run.append(fc)
                par._p.append(run)
        self.has_toc_pages = bool(pages)

    def _emit_bibliography(self):
        manual = None
        nxt = self.blocks[self.bi + 1] if self.bi + 1 < len(self.blocks) else None
        if nxt is not None and nxt.kind == "list" and nxt.data["style"] == "digit" and not self.sorted_sources:
            manual = nxt.data["items"]
            self.res.stats["sources"] = len(manual)
        elif nxt is not None and nxt.kind == "list":
            self.warn("под заголовком списка источников есть ручной список – он игнорируется, список строится из sources.json")
        entries: list[str] = []
        if manual is not None:
            entries = manual
            self._skip_next_list = True
        else:
            w: list[str] = []
            for d in self.sorted_sources:
                try:
                    entries.append(bib.format_entry(d, w))
                except bib.BibError as e:
                    self.err(0, str(e))
            for x in dict.fromkeys(w):
                self.warn(x)
        if not entries:
            self.err(0, "список источников пуст: процитируйте источники через [@id] или добавьте sources.json")
            return
        nid = self._new_num(93)
        first = True
        for text in entries:
            par = self.doc.add_paragraph(style="Список")
            ppr = par._p.get_or_add_pPr()
            ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'))
            self._gap(par, False)
            par.add_run(text)
            self.after_gap = False
        self.after_gap = False

    # ---------------------------------------------------------------- абзацы
    def _emit_para(self, b: Block):
        if self._cur_struct == C.TOC:
            self.warn(f"строка {b.line}: текст внутри «Содержание» игнорируется – оглавление строится автоматически")
            return
        if self._cur_struct == C.BIB:
            self.warn(f"строка {b.line}: текст внутри списка источников игнорируется")
            return
        self._text_par(b.data["text"], b.line)

    def _emit_list(self, b: Block):
        if self._cur_struct == C.BIB:
            return
        kind = b.data["style"]
        nid = self.num_dash if kind == "dash" else self._new_num(91 if kind == "digit" else 92)
        for item in b.data["items"]:
            par = self.doc.add_paragraph(style="Список")
            ppr = par._p.get_or_add_pPr()
            ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'))
            self._gap(par, False)
            spans = normalize_spans(list(self.spans(item, b.line)), self.typo, self.auto_italic)
            self._runs(par, spans, line=b.line)
            self.after_gap = False

    def _emit_pagebreak(self, b: Block):
        par = self.doc.add_paragraph(style="Normal")
        par.paragraph_format.first_line_indent = Cm(0)
        par.add_run().add_break(WD_BREAK.PAGE)
        self.after_gap = False

    # --------------------------------------------------------------------- код
    def _emit_code(self, b: Block):
        lines = b.data["lines"] or [""]
        while lines and not lines[-1].strip():
            lines = lines[:-1]
        for k, text in enumerate(lines):
            par = self.doc.add_paragraph(style="Код")
            if k == 0:
                self._gap(par, True)
            else:
                par.paragraph_format.space_before = Pt(0)
            if text:
                par.add_run(text)
            par.paragraph_format.widow_control = False
        self.after_gap = True

    # ----------------------------------------------------------------- рисунки
    def _find_image(self, path: str) -> str:
        cands = [path]
        if not os.path.isabs(path):
            cands = [os.path.join(self.base_dir, path)]
            if self.meta.get("images_dir"):
                cands.append(os.path.join(self.base_dir, str(self.meta["images_dir"]), path))
        for c in cands:
            if os.path.isfile(c):
                return c
        raise BuildError(f"файл рисунка не найден: {path}")

    def _emit_figure(self, b: Block):
        from PIL import Image
        d = b.data
        path = self._find_image(d["path"])
        with Image.open(path) as im:
            wpx, hpx = im.size
            dpi = im.info.get("dpi", (96, 96))[0] or 96
        w_cm = wpx / float(dpi) * 2.54
        h_over_w = hpx / float(wpx)
        want = d.get("width")
        if want:
            m = re.fullmatch(r"([\d.]+)\s*(cm|мм|mm|см|%)?", want)
            if not m:
                raise BuildError(f"непонятная ширина рисунка: {want}")
            val = float(m.group(1))
            unit = m.group(2) or "cm"
            w_cm = val / 10 if unit in ("mm", "мм") else (C.TEXT_WIDTH_CM * val / 100 if unit == "%" else val)
        w_cm = min(w_cm, C.TEXT_WIDTH_CM)
        if w_cm * h_over_w > C.MAX_IMAGE_HEIGHT_CM:
            w_cm = C.MAX_IMAGE_HEIGHT_CM / h_over_w
        par = self.doc.add_paragraph(style="Рисунок")
        self._gap(par, True)
        run = par.add_run()
        shape = run.add_picture(path, width=Cm(w_cm))
        cap_plain = re.sub(r"[*`$]", "", d["caption"])
        shape._inline.docPr.set("descr", cap_plain)
        shape._inline.docPr.set("name", f"Рисунок {d['label']}")
        cap = self.doc.add_paragraph(style="Подпись рисунка")
        cap.paragraph_format.space_before = Pt(0)
        spans = normalize_spans(list(self.spans(d["caption"], b.line)), self.typo, self.auto_italic)
        if spans and spans[-1].kind == "text":
            spans[-1] = Span("text", spans[-1].text.rstrip().rstrip("."), spans[-1].bold, spans[-1].italic)
        cap.add_run(f"Рисунок {d['label']} {C.DASH} ")
        self._runs(cap, spans, line=b.line)
        self.after_gap = True

    # ------------------------------------------------------------------ формулы
    def _emit_math(self, b: Block):
        d = b.data
        try:
            xml = latex_to_omml_xml(d["latex"], C.FONT_SIZE, C.MATH_FONT)
        except LatexError as e:
            raise BuildError(f"формула: {e}  [{d['latex']}]")
        tbl = self.doc.add_table(rows=1, cols=2)
        self._table_counter += 1
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        total = int(C.TEXT_WIDTH_CM * 567)
        wnum = int(1.9 * 567)
        widths = [total - wnum, wnum]
        for i, w in enumerate(widths):
            tbl.columns[i].width = Twips(w)
            tbl.rows[0].cells[i].width = Twips(w)
        c0, c1 = tbl.rows[0].cells
        p0 = c0.paragraphs[0]
        p0.style = self.doc.styles["Формула"]
        p0._p.append(parse_xml(xml))
        p1 = c1.paragraphs[0]
        p1.style = self.doc.styles["Формула"]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.add_run(f"({d['label']})")
        for c in (c0, c1):
            tcpr = c._tc.get_or_add_tcPr()
            tcpr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
        self._table_gap(tbl, wants_before=True)
        self.after_gap = True

    def _table_gap(self, tbl, wants_before: bool):
        """Пустая строка перед таблицей реализуется отступом первого абзаца таблицы."""
        need = self.after_gap or wants_before
        if need:
            for cell in tbl.rows[0].cells:
                cell.paragraphs[0].paragraph_format.space_before = Pt(self.blank_pt)

    # ------------------------------------------------------------------ таблицы
    def _emit_table(self, b: Block):
        d = b.data
        tid = d.get("id") or f"t{self._table_counter}"
        header, rows, aligns = d["header"], d["rows"], d["aligns"]
        ncols = len(header)
        cap = self.doc.add_paragraph(style="Подпись таблицы")
        self._gap(cap, True)
        spans = normalize_spans(list(self.spans(d["caption"], b.line)), self.typo, self.auto_italic)
        if spans and spans[-1].kind == "text":
            spans[-1] = Span("text", spans[-1].text.rstrip().rstrip("."), spans[-1].bold, spans[-1].italic)
        cap.add_run(f"Таблица {d['label']} {C.DASH} ")
        self._runs(cap, spans, line=b.line)

        # веса столбцов
        def plain(s):
            return re.sub(r"[*`$]", "", s)
        weights = []
        for c in range(ncols):
            longest = max([len(plain(header[c]))] + [len(plain(r[c])) for r in rows] + [1])
            weights.append(min(max(longest, 5), 40))
        total = int(C.TEXT_WIDTH_CM * 567)
        widths = [max(int(total * w / sum(weights)), 700) for w in weights]

        # разбиение на части (по границам страниц из Word)
        breaks = sorted(set(self.table_breaks.get(tid, [])))
        breaks = [x for x in breaks if 0 < x < len(rows)]
        bounds = [0] + breaks + [len(rows)]
        keep_together = len(rows) <= 18 and not breaks
        for part in range(len(bounds) - 1):
            lo, hi = bounds[part], bounds[part + 1]
            if part > 0:
                cont = self.doc.add_paragraph(style="Подпись таблицы")
                cont.paragraph_format.space_before = Pt(0)
                cont.paragraph_format.page_break_before = True
                cont.add_run(f"Продолжение таблицы {d['label']}")
            tbl = self.doc.add_table(rows=1 + (hi - lo), cols=ncols)
            self._table_counter += 1
            tbl.style = self.doc.styles["Table Grid"]
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = True
            tblPr = tbl._tbl.tblPr
            tblw = tblPr.find(qn("w:tblW"))
            if tblw is None:
                tblw = OxmlElement("w:tblW")
                tblPr.append(tblw)
            tblw.set(qn("w:type"), "pct")
            tblw.set(qn("w:w"), "5000")
            tblPr.append(parse_xml(f'<w:tblCaption {nsdecls("w")} w:val="{tid}|{lo}"/>'))
            for i, w in enumerate(widths):
                tbl.columns[i].width = Twips(w)
            all_rows = [header] + rows[lo:hi]
            for ri, row_cells in enumerate(all_rows):
                row = tbl.rows[ri]
                trpr = row._tr.get_or_add_trPr()
                trpr.append(OxmlElement("w:cantSplit"))
                if ri == 0:
                    trpr.append(OxmlElement("w:tblHeader"))
                for ci, text in enumerate(row_cells):
                    cell = row.cells[ci]
                    cell.width = Twips(widths[ci])
                    par = cell.paragraphs[0]
                    par.style = self.doc.styles["Текст таблицы"]
                    al = "center" if ri == 0 else aligns[ci]
                    par.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                                     "right": WD_ALIGN_PARAGRAPH.RIGHT}[al]
                    keep = ri == 0 or (keep_together and ri < len(all_rows) - 1)
                    if keep:
                        par.paragraph_format.keep_with_next = True
                    spans = normalize_spans(list(self.spans(text, b.line)), self.typo, self.auto_italic)
                    self._runs(par, spans, size=C.TABLE_FONT_SIZE, math_size=C.TABLE_FONT_SIZE, line=b.line)
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            self.res.tables.append(dict(tid=tid, start=lo, rows=hi - lo))
        self.after_gap = True

    # ---------------------------------------------------------------- завершение
    def _finish(self, out_path: str):
        pages = self.layout.get("bookmarks") or {}
        if self.title_data and pages and self.heads:
            first = self.heads[min(self.heads)].bookmark
            if pages.get(first, 2) != 2:
                self.warn("титульный лист не умещается на одной странице: сократите тему/название дисциплины/направление")
        settings = self.doc.settings.element
        if not self.layout.get("bookmarks") and any(h.name == C.TOC for h in self.heads.values() if h.cls == "struct"):
            uf = OxmlElement("w:updateFields")
            uf.set(qn("w:val"), "true")
            settings.append(uf)
        self.doc.save(out_path)
