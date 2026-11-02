"""Мини-конвертер LaTeX -> OMML (встроенные формулы Word) без внешних зависимостей.

Поддерживается подмножество: дроби, степени/индексы, корни, суммы/интегралы/пределы, греческие буквы,
стрелки и отношения, \\left..\\right, \\text/\\mathrm, \\overline/\\vec/\\hat, матрицы и cases.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "varepsilon": "ε", "zeta": "ζ",
    "eta": "η", "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν",
    "xi": "ξ", "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ", "sigma": "σ", "varsigma": "ς", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ",
    "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}
SYMBOLS = {
    "cdot": "⋅", "times": "×", "pm": "±", "mp": "∓", "div": "÷", "ast": "∗", "star": "⋆", "circ": "∘",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠", "approx": "≈", "equiv": "≡",
    "sim": "∼", "simeq": "≃", "cong": "≅", "propto": "∝", "ll": "≪", "gg": "≫",
    "to": "→", "rightarrow": "→", "leftarrow": "←", "gets": "←", "leftrightarrow": "↔",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔", "implies": "⇒", "iff": "⇔", "mapsto": "↦",
    "uparrow": "↑", "downarrow": "↓",
    "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "subseteq": "⊆", "supset": "⊃", "supseteq": "⊇",
    "cup": "∪", "cap": "∩", "setminus": "∖", "emptyset": "∅", "varnothing": "∅",
    "infty": "∞", "partial": "∂", "nabla": "∇", "forall": "∀", "exists": "∃", "neg": "¬", "lnot": "¬",
    "land": "∧", "wedge": "∧", "lor": "∨", "vee": "∨", "oplus": "⊕", "otimes": "⊗", "ominus": "⊖",
    "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱",
    "angle": "∠", "perp": "⊥", "parallel": "∥", "degree": "°", "prime": "′", "hbar": "ℏ", "ell": "ℓ",
    "Re": "ℜ", "Im": "ℑ", "aleph": "ℵ", "top": "⊤", "bot": "⊥", "vdash": "⊢", "models": "⊨",
    "langle": "⟨", "rangle": "⟩", "lceil": "⌈", "rceil": "⌉", "lfloor": "⌊", "rfloor": "⌋",
    "mid": "∣", "nmid": "∤", "vert": "|", "Vert": "‖", "lvert": "|", "rvert": "|", "lVert": "‖", "rVert": "‖",
    "because": "∵", "therefore": "∴", "triangle": "△", "square": "□", "blacksquare": "■", "qed": "∎", "|": "‖", "{": "{", "}": "}", "%": "%", "$": "$", "&": "&", "#": "#", "_": "_",
}
SPACES = {",": " ", ";": " ", ":": " ", "!": "", " ": " ", "quad": " ", "qquad": "  "}
FUNCS = {"sin", "cos", "tan", "tg", "ctg", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "arctg", "sinh", "cosh",
         "tanh", "ln", "log", "lg", "exp", "det", "dim", "ker", "deg", "gcd", "mod", "arg", "sgn", "min", "max",
         "sup", "inf", "lim", "limsup", "liminf", "Pr", "rank", "tr"}
LIMIT_FUNCS = {"lim", "limsup", "liminf", "max", "min", "sup", "inf", "arg"}
NARY = {"sum": "∑", "prod": "∏", "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮", "bigcup": "⋃",
        "bigcap": "⋂", "coprod": "∐", "bigoplus": "⨁", "bigvee": "⋁", "bigwedge": "⋀"}
ACCENTS = {"hat": "̂", "tilde": "̃", "vec": "⃗", "dot": "̇", "ddot": "̈", "check": "̌", "breve": "̆", "acute": "́", "grave": "̀"}
BLACKBOARD = {"R": "ℝ", "N": "ℕ", "Z": "ℤ", "Q": "ℚ", "C": "ℂ", "P": "ℙ", "E": "𝔼"}
DELIMS_OPEN = {"(": "(", "[": "[", "\\{": "{", "\\lbrace": "{", "|": "|", "\\|": "‖", "\\langle": "⟨",
               "\\lfloor": "⌊", "\\lceil": "⌈", ".": ""}
DELIMS_CLOSE = {")": ")", "]": "]", "\\}": "}", "\\rbrace": "}", "|": "|", "\\|": "‖", "\\rangle": "⟩",
                "\\rfloor": "⌋", "\\rceil": "⌉", ".": ""}


class LatexError(ValueError):
    pass


class Atom:
    """Элемент формулы: либо текстовый run (text,sty), либо готовый XML."""
    __slots__ = ("text", "sty", "xml", "op")

    def __init__(self, text=None, sty=None, xml=None, op=False):
        self.text, self.sty, self.xml, self.op = text, sty, xml, op


class Converter:
    def __init__(self, size_pt: int = 14, math_font: str = "Cambria Math"):
        self.sz = int(size_pt * 2)
        self.font = math_font
        self.s = ""
        self.i = 0

    # ------------------------------------------------------------- вывод XML
    def run(self, text: str, sty: str | None = None) -> str:
        rpr = f'<m:rPr><m:sty m:val="{sty}"/></m:rPr>' if sty else ""
        return (f'<m:r>{rpr}<w:rPr><w:rFonts w:ascii="{self.font}" w:hAnsi="{self.font}" w:cs="{self.font}"/>'
                f'<w:sz w:val="{self.sz}"/><w:szCs w:val="{self.sz}"/></w:rPr><m:t xml:space="preserve">{escape(text)}</m:t></m:r>')

    def atom_xml(self, a: Atom) -> str:
        return a.xml if a.xml is not None else self.run(a.text, a.sty)

    def seq_xml(self, atoms: list[Atom]) -> str:
        out, buf = [], None          # buf = [text, sty]
        for a in atoms:
            if a.xml is None:
                if buf is not None and buf[1] == a.sty:
                    buf[0] += a.text
                else:
                    if buf is not None:
                        out.append(self.run(buf[0], buf[1]))
                    buf = [a.text, a.sty]
            else:
                if buf is not None:
                    out.append(self.run(buf[0], buf[1]))
                    buf = None
                out.append(a.xml)
        if buf is not None:
            out.append(self.run(buf[0], buf[1]))
        return "".join(out)

    def e(self, atoms: list[Atom]) -> str:
        return f"<m:e>{self.seq_xml(atoms)}</m:e>"

    # -------------------------------------------------------------- парсер
    def peek(self) -> str:
        return self.s[self.i] if self.i < len(self.s) else ""

    def skip_ws(self):
        while self.i < len(self.s) and self.s[self.i] in " \t\n":
            self.i += 1

    def read_cmd(self) -> str:
        """После '\\': имя команды или один спецсимвол."""
        j = self.i
        if j < len(self.s) and self.s[j].isalpha():
            while j < len(self.s) and self.s[j].isalpha():
                j += 1
        else:
            j = min(j + 1, len(self.s))
        name = self.s[self.i:j]
        self.i = j
        return name

    def parse_group(self) -> list[Atom]:
        """Читает {…} или один атом."""
        self.skip_ws()
        if self.peek() == "{":
            self.i += 1
            atoms = self.parse_seq(until="}")
            if self.peek() != "}":
                raise LatexError("не закрыта фигурная скобка")
            self.i += 1
            return atoms
        a = self.parse_atom(single=True)
        return [a] if a is not None else []

    def parse_seq(self, until: str | None = None, end_cmds: tuple = ()) -> list[Atom]:
        atoms: list[Atom] = []
        while self.i < len(self.s):
            ch = self.peek()
            if until and ch == until:
                break
            if ch == "\\":
                save = self.i
                self.i += 1
                name = self.read_cmd()
                self.i = save
                if name in end_cmds:
                    break
            a = self.parse_atom()
            if a is None:
                continue
            a = self.parse_scripts(a, atoms)
            atoms.append(a)
        return atoms

    def parse_scripts(self, base: Atom, prev: list[Atom]) -> Atom:
        sub = sup = None
        while True:
            self.skip_ws()
            ch = self.peek()
            if ch == "_" and sub is None:
                self.i += 1
                sub = self.parse_group()
            elif ch == "^" and sup is None:
                self.i += 1
                sup = self.parse_group()
            elif ch == "'" and sup is None:
                self.i += 1
                sup = [Atom("′")]
            else:
                break
        if sub is None and sup is None:
            return base
        bx = self.atom_xml(base)
        if base.xml is not None and base.op == "nary":      # крупный оператор: пределы уже внутри
            return base
        if sub is not None and sup is not None:
            return Atom(xml=f"<m:sSubSup><m:e>{bx}</m:e><m:sub>{self.seq_xml(sub)}</m:sub><m:sup>{self.seq_xml(sup)}</m:sup></m:sSubSup>")
        if sub is not None:
            return Atom(xml=f"<m:sSub><m:e>{bx}</m:e><m:sub>{self.seq_xml(sub)}</m:sub></m:sSub>")
        return Atom(xml=f"<m:sSup><m:e>{bx}</m:e><m:sup>{self.seq_xml(sup)}</m:sup></m:sSup>")

    def read_limits(self):
        sub = sup = None
        while True:
            self.skip_ws()
            ch = self.peek()
            if ch == "_" and sub is None:
                self.i += 1
                sub = self.parse_group()
            elif ch == "^" and sup is None:
                self.i += 1
                sup = self.parse_group()
            else:
                return sub, sup

    def operand(self) -> list[Atom]:
        self.skip_ws()
        if self.i >= len(self.s) or self.peek() in "}&" or self.s.startswith("\\\\", self.i):
            return []
        a = self.parse_atom()
        if a is None:
            return []
        a = self.parse_scripts(a, [])
        return [a]

    def parse_atom(self, single: bool = False) -> Atom | None:
        ch = self.peek()
        if ch == "":
            return None
        if ch in " \t\n":
            self.i += 1
            return None
        if ch == "{":
            atoms = self.parse_group()
            return Atom(xml=self.seq_xml(atoms)) if atoms else None
        if ch == "}":
            raise LatexError("лишняя закрывающая скобка }")
        if ch == "\\":
            self.i += 1
            return self.parse_command()
        self.i += 1
        if ch.isdigit():
            if single:
                return Atom(ch, "p")
            j = self.i
            while j < len(self.s) and (self.s[j].isdigit() or (self.s[j] in ".," and j + 1 < len(self.s) and self.s[j + 1].isdigit())):
                j += 1
            num = ch + self.s[self.i:j]
            self.i = j
            return Atom(num, "p")
        if ch.isalpha():
            return Atom(ch, None)
        if ch in "+-=<>*/:;,.!?()[]|~":
            if ch == "-":
                ch = "−"
            if ch == "*":
                ch = "∗"
            if ch == "~":
                return Atom(" ", "p")
            return Atom(ch, "p", op=ch in "+−=<>/")
        return Atom(ch, None)

    # ------------------------------------------------------------- команды
    def parse_command(self) -> Atom | None:
        name = self.read_cmd()
        if name == "" or name == " ":
            return Atom(" ", "p")
        if name == "bmod":
            return Atom("\u2005mod\u2005", "p")
        if name in ("pmod", "pod"):
            arg = self.parse_group()
            head = [Atom("(", "p")] + ([Atom("mod", "p"), Atom("\u2005", "p")] if name == "pmod" else [])
            return Atom(xml=self.seq_xml(head + arg + [Atom(")", "p")]))
        if name in SPACES:
            t = SPACES[name]
            return Atom(t, "p") if t else None
        if name == "\\":
            return Atom(" ", "p")
        if name in GREEK:
            return Atom(GREEK[name], None)
        if name in NARY:
            return self.nary(NARY[name])
        if name in SYMBOLS:
            return Atom(SYMBOLS[name], "p", op=name in ("cdot", "times", "pm", "mp", "leq", "le", "geq", "ge", "neq", "ne",
                                                          "approx", "equiv", "to", "rightarrow", "leftarrow", "Rightarrow",
                                                          "Leftrightarrow", "leftrightarrow", "in", "notin", "subset",
                                                          "subseteq", "cup", "cap", "land", "lor", "sim", "propto", "div",
                                                          "implies", "iff", "mapsto", "ll", "gg", "oplus", "otimes"))
        if name in ("frac", "dfrac", "tfrac"):
            num, den = self.parse_group(), self.parse_group()
            return Atom(xml=f"<m:f><m:num>{self.seq_xml(num)}</m:num><m:den>{self.seq_xml(den)}</m:den></m:f>")
        if name in ("binom",):
            n_, k_ = self.parse_group(), self.parse_group()
            inner = (f'<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>{self.seq_xml(n_)}</m:num>'
                     f'<m:den>{self.seq_xml(k_)}</m:den></m:f>')
            return Atom(xml=f"<m:d><m:e>{inner}</m:e></m:d>")
        if name == "sqrt":
            self.skip_ws()
            deg = None
            if self.peek() == "[":
                j = self.s.find("]", self.i)
                if j < 0:
                    raise LatexError("не закрыта [ в \\sqrt")
                sub = Converter(self.sz // 2, self.font)
                deg = sub.convert_atoms(self.s[self.i + 1:j])
                self.i = j + 1
            body = self.parse_group()
            if deg:
                return Atom(xml=f"<m:rad><m:deg>{self.seq_xml(deg)}</m:deg><m:e>{self.seq_xml(body)}</m:e></m:rad>")
            return Atom(xml=f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{self.seq_xml(body)}</m:e></m:rad>')
        if name in ("text", "textrm", "mathrm", "operatorname", "mbox", "textit", "textbf", "mathbf", "mathit", "boldsymbol", "bm", "mathcal", "mathbb", "mathsf", "mathtt"):
            if name in ("text", "textrm", "mbox") or name in ("mathrm", "operatorname"):
                self.skip_ws()
                if self.peek() == "{":
                    j = self._match_brace(self.i)
                    raw = self.s[self.i + 1:j]
                    self.i = j + 1
                else:
                    raw, self.i = self.peek(), self.i + 1
                if name in ("text", "textrm", "mbox"):
                    return Atom(raw, "p")
                return Atom(raw, "p")
            atoms = self.parse_group()
            if name == "mathbb":
                txt = "".join(BLACKBOARD.get(a.text, a.text) if a.text else "" for a in atoms if a.xml is None)
                return Atom(txt, "p")
            style = {"textit": "i", "mathit": "i", "textbf": "b", "mathbf": "b", "boldsymbol": "bi", "bm": "bi"}.get(name, "p")
            txt = "".join(a.text or "" for a in atoms if a.xml is None)
            return Atom(txt, style)
        if name in ("overline", "bar", "underline"):
            body = self.parse_group()
            pos = "bot" if name == "underline" else "top"
            return Atom(xml=f'<m:bar><m:barPr><m:pos m:val="{pos}"/></m:barPr><m:e>{self.seq_xml(body)}</m:e></m:bar>')
        if name in ACCENTS or name in ("widehat", "widetilde", "overrightarrow"):
            chr_ = {"widehat": "̂", "widetilde": "̃", "overrightarrow": "⃗"}.get(name) or ACCENTS[name]
            body = self.parse_group()
            return Atom(xml=f'<m:acc><m:accPr><m:chr m:val="{chr_}"/></m:accPr><m:e>{self.seq_xml(body)}</m:e></m:acc>')
        if name in ("overbrace", "underbrace"):
            body = self.parse_group()
            pos = "top" if name == "overbrace" else "bot"
            ch = "⏞" if name == "overbrace" else "⏟"
            return Atom(xml=f'<m:groupChr><m:groupChrPr><m:chr m:val="{ch}"/><m:pos m:val="{pos}"/></m:groupChrPr><m:e>{self.seq_xml(body)}</m:e></m:groupChr>')
        if name == "left":
            return self.delimiters()
        if name == "right":
            raise LatexError("\\right без \\left")
        if name == "begin":
            return self.environment()
        if name in ("big", "Big", "bigg", "Bigg", "bigl", "bigr", "displaystyle", "textstyle", "limits", "nolimits"):
            return None
        if name in FUNCS:
            if name in LIMIT_FUNCS:
                sub, sup = self.read_limits()
                if sub is not None and sup is None:
                    return Atom(xml=f"<m:limLow><m:e>{self.run(name, 'p')}</m:e><m:lim>{self.seq_xml(sub)}</m:lim></m:limLow>")
                if sub is not None or sup is not None:
                    return self._scripts_atom(Atom(name, "p"), sub, sup)
            return Atom(name, "p")
        if name in ("%", "&", "#", "_", "$", "{", "}"):
            return Atom(name, "p")
        raise LatexError(f"неподдерживаемая команда \\{name}")

    def _scripts_atom(self, base: Atom, sub, sup) -> Atom:
        bx = self.atom_xml(base)
        if sub is not None and sup is not None:
            return Atom(xml=f"<m:sSubSup><m:e>{bx}</m:e><m:sub>{self.seq_xml(sub)}</m:sub><m:sup>{self.seq_xml(sup)}</m:sup></m:sSubSup>")
        if sub is not None:
            return Atom(xml=f"<m:sSub><m:e>{bx}</m:e><m:sub>{self.seq_xml(sub)}</m:sub></m:sSub>")
        return Atom(xml=f"<m:sSup><m:e>{bx}</m:e><m:sup>{self.seq_xml(sup)}</m:sup></m:sSup>")

    def _match_brace(self, start: int) -> int:
        depth = 0
        for j in range(start, len(self.s)):
            if self.s[j] == "{":
                depth += 1
            elif self.s[j] == "}":
                depth -= 1
                if depth == 0:
                    return j
        raise LatexError("не закрыта фигурная скобка")

    def nary(self, chr_: str) -> Atom:
        sub, sup = self.read_limits()
        operand = self.operand()
        is_int = chr_ in "∫∬∭∮"
        pr = f'<m:chr m:val="{chr_}"/><m:limLoc m:val="{"subSup" if is_int else "undOvr"}"/>'
        if sub is None:
            pr += '<m:subHide m:val="1"/>'
        if sup is None:
            pr += '<m:supHide m:val="1"/>'
        xml = (f"<m:nary><m:naryPr>{pr}</m:naryPr><m:sub>{self.seq_xml(sub or [])}</m:sub>"
               f"<m:sup>{self.seq_xml(sup or [])}</m:sup><m:e>{self.seq_xml(operand)}</m:e></m:nary>")
        return Atom(xml=xml, op="nary")

    def delimiters(self) -> Atom:
        def read_delim(table):
            self.skip_ws()
            if self.peek() == "\\":
                self.i += 1
                tok = "\\" + self.read_cmd()
            else:
                tok = self.peek()
                self.i += 1
            if tok not in table:
                raise LatexError(f"неизвестный разделитель «{tok}»")
            return table[tok]

        beg = read_delim(DELIMS_OPEN)
        inner = self.parse_seq(end_cmds=("right",))
        if not self.s.startswith("\\right", self.i):
            raise LatexError("\\left без \\right")
        self.i += len("\\right")
        end = read_delim(DELIMS_CLOSE)
        return Atom(xml=f'<m:d><m:dPr><m:begChr m:val="{beg}"/><m:endChr m:val="{end}"/></m:dPr><m:e>{self.seq_xml(inner)}</m:e></m:d>')

    def environment(self) -> Atom:
        self.skip_ws()
        j = self._match_brace(self.i)
        env = self.s[self.i + 1:j]
        self.i = j + 1
        rows: list[list[list[Atom]]] = [[]]
        cell_atoms: list[Atom] = []
        while True:
            seg = self.parse_seq(end_cmds=("end", "\\"), until="&")
            cell_atoms.extend(seg)
            if self.peek() == "&":
                self.i += 1
                rows[-1].append(cell_atoms)
                cell_atoms = []
                continue
            if self.s.startswith("\\\\", self.i):
                self.i += 2
                rows[-1].append(cell_atoms)
                cell_atoms = []
                rows.append([])
                continue
            if self.s.startswith("\\end", self.i):
                self.i += 4
                self.skip_ws()
                k = self._match_brace(self.i)
                if self.s[self.i + 1:k] != env:
                    raise LatexError(f"\\begin{{{env}}} закрыт как \\end{{{self.s[self.i + 1:k]}}}")
                self.i = k + 1
                rows[-1].append(cell_atoms)
                break
            raise LatexError(f"не закрыта среда {env}")
        rows = [r for r in rows if any(c for c in r)]
        EM = Atom(chr(0x2003), "p")
        ALIGN = Atom("&", "p")          # точка выравнивания строк массива уравнений в OMML
        if env == "cases":
            # строки прижаты влево (& в начале), без maxDist: иначе Word растягивает блок на всю строку
            def case_row(r):
                atoms = [ALIGN]
                for k, c in enumerate(r):
                    atoms += ([EM] if k else []) + c
                return atoms
            eq = "".join(f"<m:e>{self.seq_xml(case_row(r))}</m:e>" for r in rows)
            return Atom(xml=f'<m:d><m:dPr><m:begChr m:val="{{"/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>{eq}</m:eqArr></m:e></m:d>')
        if env in ("align", "aligned", "align*", "gather", "gathered", "eqnarray", "split"):
            def align_row(r):
                atoms = []
                for k, c in enumerate(r):
                    atoms += ([ALIGN] if k else []) + c      # & из LaTeX -> точка выравнивания
                return atoms
            eq = "".join(f"<m:e>{self.seq_xml(align_row(r))}</m:e>" for r in rows)
            return Atom(xml=f"<m:eqArr>{eq}</m:eqArr>")
        brackets = {"pmatrix": ("(", ")"), "bmatrix": ("[", "]"), "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"),
                    "Vmatrix": ("‖", "‖"), "matrix": ("", ""), "smallmatrix": ("", ""), "array": ("", "")}
        if env not in brackets:
            raise LatexError(f"неподдерживаемая среда {env}")
        b, e_ = brackets[env]
        ncols = max(len(r) for r in rows)
        mrows = ""
        for r in rows:
            cells = r + [[] for _ in range(ncols - len(r))]
            mrows += "<m:mr>" + "".join(f"<m:e>{self.seq_xml(c)}</m:e>" for c in cells) + "</m:mr>"
        mat = f"<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"{ncols}\"/><m:mcJc m:val=\"center\"/></m:mcPr></m:mc></m:mcs></m:mPr>{mrows}</m:m>"
        if not b and not e_:
            return Atom(xml=mat)
        return Atom(xml=f'<m:d><m:dPr><m:begChr m:val="{b}"/><m:endChr m:val="{e_}"/></m:dPr><m:e>{mat}</m:e></m:d>')

    # ---------------------------------------------------------------- API
    def convert_atoms(self, latex: str) -> list[Atom]:
        saved = (self.s, self.i)
        self.s, self.i = latex.strip(), 0
        atoms = self.parse_seq()
        if self.i < len(self.s):
            raise LatexError(f"неожиданный символ «{self.s[self.i]}» в позиции {self.i}")
        self.s, self.i = saved
        return atoms

    def convert(self, latex: str) -> str:
        """Возвращает XML m:oMath."""
        atoms = self.convert_atoms(latex)
        return f'<m:oMath xmlns:m="{M_NS}" xmlns:w="{W_NS}">{self.seq_xml(atoms)}</m:oMath>'


def latex_to_omml_xml(latex: str, size_pt: int = 14, math_font: str = "Cambria Math") -> str:
    return Converter(size_pt, math_font).convert(latex)
