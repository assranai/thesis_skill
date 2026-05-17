#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ⚠️ 新团队接手前必须先读: ../HANDOFF_遗嘱_必读.md
"""OMML formula generator from LaTeX."""

import re
from xml.sax.saxutils import escape

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GREEK_MAP = {
    "alpha": "α", "beta": "β", "gamma": "γ",
    "delta": "δ", "epsilon": "ε", "varepsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ",
    "vartheta": "θ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν",
    "xi": "ξ", "omicron": "ο", "pi": "π",
    "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ",
    "sigma": "σ", "varsigma": "ς", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ",
    "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ",
    "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π",
    "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ",
    "Psi": "Ψ", "Omega": "Ω",
}

OP_MAP = {
    "le": "≤", "ge": "≥", "leq": "≤", "geq": "≥",
    "times": "×", "div": "÷", "pm": "±",
    "rightarrow": "→", "leftarrow": "←",
    "Rightarrow": "⇒", "Leftarrow": "⇐",
    "leftrightarrow": "↔", "Leftrightarrow": "⇔",
    "infty": "∞", "partial": "∂", "nabla": "∇",
    "exists": "∃", "forall": "∀",
    "in": "∈", "notin": "∉",
    "subset": "⊂", "supset": "⊃",
    "subseteq": "⊆", "supseteq": "⊇",
    "cup": "∪", "cap": "∩",
    "emptyset": "∅", "varnothing": "∅",
    "cdots": "⋯", "ldots": "…",
    "vdots": "⋮", "ddots": "⋱",
    "approx": "≈", "neq": "≠", "ne": "≠",
    "equiv": "≡", "sim": "∼", "propto": "∝",
    "perp": "⊥", "angle": "∠", "triangle": "△",
    "surd": "√", "imath": "ı", "jmath": "ȷ",
    "ell": "ℓ", "hbar": "ℏ", "Re": "ℜ", "Im": "ℑ",
    "aleph": "ℵ", "wp": "℘",
    "otimes": "⊗", "oplus": "⊕", "oslash": "⊘", "odot": "⊙",
    "circ": "∘", "bullet": "∙",
    "dagger": "†", "ddagger": "‡",
    "neg": "¬", "land": "∧", "lor": "∨",
    "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
    "sum": "∑", "prod": "∏", "coprod": "∐",
    "prime": "′", "degree": "°",
    "cdot": "·", "ast": "∗", "star": "☆",
    "bigcirc": "○", "bigtriangleup": "△", "bigtriangledown": "▽",
    "backslash": "\\",
}

FUNC_NAMES = {
    "max", "min", "arg", "argmax", "argmin",
    "sin", "cos", "tan", "cot", "sec", "csc",
    "sinh", "cosh", "tanh", "coth",
    "arcsin", "arccos", "arctan",
    "log", "ln", "lg",
    "exp", "det", "dim", "hom", "ker", "tr",
    "deg", "Pr", "Var", "Cov", "Corr",
    "lim", "limsup", "liminf",
    "sup", "inf",
    "mod", "bmod", "pmod",
    "gcd", "lcm",
    "erf", "erfc",
}


def _cr():
    """Return Cambria Math run properties XML fragment."""
    return (
        '<w:rPr>'
        '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
        '<w:sz w:val="24"/>'
        '</w:rPr>'
    )


def _ctrl_pr():
    """Return ctrlPr with Cambria Math font."""
    return (
        '<m:ctrlPr>'
        '<w:rPr>'
        '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
        '<w:sz w:val="24"/>'
        '</w:rPr>'
        '</m:ctrlPr>'
    )


def _make_run(text):
    """Create an OMML text run <m:r> with Cambria Math font."""
    if not text:
        return ""
    escaped = escape(text)
    return "<m:r>" + _cr() + "<m:t>" + escaped + "</m:t></m:r>"


def _make_run_sty(text, sty="p"):
    """Create an OMML text run with a style attribute."""
    if not text:
        return ""
    escaped = escape(text)
    return (
        '<m:r><m:rPr><m:sty m:val="' + sty + '"/></m:rPr>'
        + _cr()
        + "<m:t>" + escaped + "</m:t></m:r>"
    )


def _make_subscript(base, sub):
    """Create OMML subscript <m:sSub>."""
    return (
        "<m:sSub>"
        + '<m:sSubPr>' + _ctrl_pr() + '</m:sSubPr>'
        + ("<m:e>" + base + "</m:e>" if base else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "<m:sub>" + sub + "</m:sub>"
        + "</m:sSub>"
    )


def _make_superscript(base, sup):
    """Create OMML superscript <m:sSup>."""
    return (
        "<m:sSup>"
        + '<m:sSupPr>' + _ctrl_pr() + '</m:sSupPr>'
        + ("<m:e>" + base + "</m:e>" if base else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "<m:sup>" + sup + "</m:sup>"
        + "</m:sSup>"
    )


def _make_subsup(base, sub, sup):
    """Create OMML combined sub/sup <m:sSubSup>."""
    return (
        "<m:sSubSup>"
        + '<m:sSubSupPr>' + _ctrl_pr() + '</m:sSubSupPr>'
        + ("<m:e>" + base + "</m:e>" if base else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "<m:sub>" + sub + "</m:sub>"
        + "<m:sup>" + sup + "</m:sup>"
        + "</m:sSubSup>"
    )


def _make_fraction(num, den):
    """Create OMML fraction <m:f>."""
    return (
        "<m:f>"
        + '<m:fPr>' + _ctrl_pr() + '</m:fPr>'
        + ("<m:num>" + num + "</m:num>" if num else "<m:num><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:num>")
        + ("<m:den>" + den + "</m:den>" if den else "<m:den><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:den>")
        + "</m:f>"
    )


def _make_nary(chr_val, sub, sup, expr, lim_loc="undOvr"):
    """Create OMML n-ary operator."""
    return (
        "<m:nary>"
        + "<m:naryPr>"
        + '<m:chr m:val="' + escape(chr_val) + '"/>'
        + '<m:limLoc m:val="' + lim_loc + '"/>'
        + _ctrl_pr()
        + "</m:naryPr>"
        + "<m:sub>" + sub + "</m:sub>"
        + "<m:sup>" + sup + "</m:sup>"
        + ("<m:e>" + expr + "</m:e>" if expr else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "</m:nary>"
    )


def _make_rad(deg, expr):
    """Create OMML radical <m:rad> (square root or nth root)."""
    return (
        "<m:rad>"
        + '<m:radPr>'
        + ('' if deg else '<m:degHide m:val="1"/>')
        + _ctrl_pr()
        + '</m:radPr>'
        + (("<m:deg>" + deg + "</m:deg>") if deg else "<m:deg/>")
        + ("<m:e>" + expr + "</m:e>" if expr else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "</m:rad>"
    )

def _make_func(func_name, lower_limit=None, arg_expr=""):
    """Create OMML function with optional lower limit."""
    if lower_limit:
        fname_content = (
            "<m:limLow>"
            + '<m:limLowPr>' + _ctrl_pr() + '</m:limLowPr>'
            + "<m:e>" + _make_run_sty(func_name) + "</m:e>"
            + "<m:lim>" + lower_limit + "</m:lim>"
            + "</m:limLow>"
        )
    else:
        fname_content = _make_run_sty(func_name)

    return (
        "<m:func>"
        + '<m:funcPr>' + _ctrl_pr() + '</m:funcPr>'
        + "<m:fName>" + fname_content + "</m:fName>"
        + ("<m:e>" + arg_expr + "</m:e>" if arg_expr else "<m:e><m:r><m:t xml:space=\"preserve\"> </m:t></m:r></m:e>")
        + "</m:func>"
    )


def _tokenize_latex(latex_str):
    """Tokenize a LaTeX math expression into a list of tokens."""
    tokens = []
    i = 0
    while i < len(latex_str):
        c = latex_str[i]
        if c == "\\":
            j = i + 1
            if j < len(latex_str) and latex_str[j] == " ":
                tokens.append("\\ ")
                i += 2
                continue
            while j < len(latex_str) and latex_str[j].isalpha():
                j += 1
            if j > i + 1:
                tokens.append(latex_str[i:j])
                i = j
            else:
                tokens.append(c)
                i += 1
        elif c in "{}_^":
            tokens.append(c)
            i += 1
        elif c.isspace():
            i += 1
        else:
            tokens.append(c)
            i += 1
    return tokens



class OMMLParser:
    """Recursive descent parser for LaTeX math -> OMML XML."""

    def __init__(self, tokens):
        self.tokens = tokens
        self.pos = 0

    def peek(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def consume(self):
        tok = self.tokens[self.pos] if self.pos < len(self.tokens) else None
        self.pos += 1
        return tok

    def expect(self, expected):
        tok = self.consume()
        if tok != expected:
            raise ValueError("Expected " + repr(expected) + ", got " + repr(tok))
        return tok

    def parse(self):
        parts = self._parse_expression()
        return self._join_parts(parts)

    def _join_parts(self, parts):
        return "".join(p for p in parts if p)

    def _parse_expression(self, stop_tokens=None):
        if stop_tokens is None:
            stop_tokens = set()
        parts = []
        while self.pos < len(self.tokens):
            tok = self.peek()
            if tok in stop_tokens:
                break
            if tok == "}":
                break

            # Handle subscript/superscript: combine with preceding atom as base
            if tok == "_":
                self.consume()
                sub = self._parse_sub_sup_body()
                if self.peek() == "^":
                    self.consume()
                    sup = self._parse_sub_sup_body()
                    if parts:
                        base = parts.pop()
                        parts.append(_make_subsup(base, sub, sup))
                    else:
                        parts.append(_make_subscript("", sub))
                        parts.append(_make_superscript("", sup))
                else:
                    if parts:
                        base = parts.pop()
                        parts.append(_make_subscript(base, sub))
                    else:
                        parts.append(_make_subscript("", sub))
                continue

            if tok == "^":
                self.consume()
                sup = self._parse_sub_sup_body()
                if self.peek() == "_":
                    self.consume()
                    sub = self._parse_sub_sup_body()
                    if parts:
                        base = parts.pop()
                        parts.append(_make_subsup(base, sub, sup))
                    else:
                        parts.append(_make_superscript("", sup))
                        parts.append(_make_subscript("", sub))
                else:
                    if parts:
                        base = parts.pop()
                        parts.append(_make_superscript(base, sup))
                    else:
                        parts.append(_make_superscript("", sup))
                continue

            part = self._parse_atom()
            if part:
                parts.append(part)
        return parts

    def _parse_atom(self):
        tok = self.peek()
        if tok is None:
            return ""
        if tok == "{":
            return self._parse_group()
        elif tok.startswith("\\"):
            return self._parse_command()
        else:
            self.consume()
            return _make_run(tok)

    def _parse_group(self):
        self.expect("{")
        parts = self._parse_expression(stop_tokens={"}"})
        self.expect("}")
        return self._join_parts(parts)

    def _parse_subscript(self, base=None):
        """Parse subscript. If base is None, uses empty base."""
        self.expect("_")
        sub = self._parse_sub_sup_body()
        base_str = base if base is not None else ""
        if self.peek() == "^":
            self.consume()
            sup = self._parse_sub_sup_body()
            return _make_subsup(base_str, sub, sup)
        return _make_subscript(base_str, sub)

    def _parse_superscript(self, base=None):
        """Parse superscript. If base is None, uses empty base."""
        self.expect("^")
        sup = self._parse_sub_sup_body()
        base_str = base if base is not None else ""
        if self.peek() == "_":
            self.consume()
            sub = self._parse_sub_sup_body()
            return _make_subsup(base_str, sub, sup)
        return _make_superscript(base_str, sup)

    def _parse_sub_sup_body(self):
        if self.peek() == "{":
            return self._parse_group()
        else:
            tok = self.consume()
            if tok and tok.startswith("\\"):
                return self._parse_command_body(tok)
            return _make_run(tok) if tok else ""

    def _parse_command(self):
        cmd = self.consume()
        return self._parse_command_body(cmd)

    def _parse_command_body(self, cmd):
        name = cmd[1:] if cmd.startswith("\\") else cmd

        # Greek letters
        if name in GREEK_MAP:
            return _make_run(GREEK_MAP[name])

        # Summation / Product / Coproduct (must check before OP_MAP)
        if name in ("sum", "prod", "coprod"):
            chr_map = {"sum": "\u2211", "prod": "\u220f", "coprod": "\u2210"}
            chr_val = chr_map[name]
            sub = ""
            sup = ""
            if self.peek() == "_":
                self.consume()
                sub = self._parse_sub_sup_body()
            if self.peek() == "^":
                self.consume()
                sup = self._parse_sub_sup_body()
            expr = ""
            if self.peek() == "{":
                expr = self._parse_group()
            elif self.peek() and self.peek() not in "}_^":
                expr = self._parse_atom()
            return _make_nary(chr_val, sub, sup, expr)

        # Integral (must check before OP_MAP)
        if name in ("int", "iint", "iiint", "oint"):
            chr_map = {"int": "\u222b", "iint": "\u222c",
                       "iiint": "\u222d", "oint": "\u222e"}
            chr_val = chr_map[name]
            sub = ""
            sup = ""
            if self.peek() == "_":
                self.consume()
                sub = self._parse_sub_sup_body()
            if self.peek() == "^":
                self.consume()
                sup = self._parse_sub_sup_body()
            expr = ""
            if self.peek() == "{":
                expr = self._parse_group()
            elif self.peek() and self.peek() not in "_^":
                expr = self._parse_atom()
            return _make_nary(chr_val, sub, sup, expr, lim_loc="subSup")

        # Operators
        if name in OP_MAP:
            return _make_run(OP_MAP[name])

        # Bold commands — parse content with full LaTeX command support, then wrap in bold
        if name in ("mathbf", "textbf", "bm"):
            if self.peek() == "{":
                content = self._parse_group()  # Full LaTeX parsing (handles \beta, \times etc.)
                if not content:
                    return ""
                # Wrap each <m:r> in bold style: inject <m:sty m:val="b"/> into rPr
                import re as _re2
                content = _re2.sub(
                    r'(<m:r>)',
                    r'<m:r><m:rPr><m:sty m:val="b"/></m:rPr>',
                    content
                )
                return content
            elif self.peek():
                return _make_run_sty(self.consume() or "", sty="b")
            return ""
        
        # Other font commands (mathrm, mathit, mathcal, etc.) — parse content, no bold
        if name in ("mathrm", "mathit", "mathcal",
                    "mathsf", "mathtt", "mathbb", "mathfrak"):
            if self.peek() == "{":
                return self._parse_group()
            elif self.peek():
                return _make_run(self.consume() or "")
            return ""

        # Text mode — consolidate multi-char text into a single OMML run
        # (sty="p" for upright Roman style, not math italic).  
        # This avoids per-character <m:r> runs that can cause spacing artifacts
        # and bloated XML in Word's OMML renderer.
        if name == "text":
            if self.peek() == "{":
                self.consume()  # consume "{"
                chars = []
                while self.pos < len(self.tokens):
                    tok = self.peek()
                    if tok == "}":
                        self.consume()
                        break
                    self.consume()
                    chars.append(tok)
                return _make_run_sty(''.join(chars), sty="p") if chars else ""
            elif self.peek():
                return _make_run_sty(self.consume() or "", sty="p")
            return ""

        # Fractions
        if name == "frac":
            self.expect("{")
            num_parts = self._parse_expression(stop_tokens={"}"})
            self.expect("}")
            self.expect("{")
            den_parts = self._parse_expression(stop_tokens={"}"})
            self.expect("}")
            num = self._join_parts(num_parts)
            den = self._join_parts(den_parts)
            return _make_fraction(num, den)

        # Functions (max, min, sin, etc.)
        if name in FUNC_NAMES:
            if self.peek() == "_":
                self.consume()
                if self.peek() == "{":
                    lim_content = self._parse_group()
                else:
                    lim_content = _make_run(self.consume() or "")
                arg = ""
                if self.peek() == "{":
                    arg = self._parse_group()
                return _make_func(name, lower_limit=lim_content, arg_expr=arg)
            else:
                arg = ""
                if self.peek() == "{":
                    arg = self._parse_group()
                return _make_func(name, arg_expr=arg)

        # \\limits
        if name == "limits":
            return ""

        # \\arg\\max, \\arg\\min
        if name == "arg":
            if self.peek() and self.peek().startswith("\\"):
                next_cmd = self.consume()
                next_name = next_cmd[2:]
                if next_name in ("max", "min"):
                    lower_limit = None
                    if self.peek() == "_":
                        self.consume()
                        lower_limit = self._parse_sub_sup_body()
                    arg = ""
                    if self.peek() == "{":
                        arg = self._parse_group()
                    return _make_func("arg\\" + next_name, lower_limit=lower_limit, arg_expr=arg)
                else:
                    return _make_run("arg") + self._parse_command_body(next_cmd)
            return _make_run("arg")

        # Accents (simplified: render content only)
        if name in ("hat", "widehat", "bar", "overline", "tilde", "widetilde",
                    "dot", "ddot", "vec", "check", "breve"):
            if self.peek() == "{":
                return self._parse_group()
            elif self.peek():
                return _make_run(self.consume() or "")
            return ""

        # \\sqrt
        if name in ("sqrt", "surd"):
            deg = ""
            if self.peek() == "[":
                self.consume()
                deg_body = []
                while self.peek() and self.peek() != "]":
                    deg_body.append(self.consume())
                if self.peek() == "]":
                    self.consume()
                deg = ''.join(deg_body)
            if self.peek() == "{":
                expr = self._parse_group()
                return _make_rad(deg, expr)
            return _make_rad("", "")

        # \\left, \\right, \\big, etc.
        if name in ("left", "right", "big", "Big", "bigg", "Bigg",
                    "bigl", "Bigl", "biggl", "Biggl",
                    "bigr", "Bigr", "biggr", "Biggr",
                    "bigm", "Bigm", "biggm", "Biggm"):
            if self.peek():
                self.consume()
            return ""

        # Style commands
        if name in ("displaystyle", "textstyle", "scriptstyle",
                    "scriptscriptstyle"):
            return ""

        # \\tag, \\label, \\nonumber
        if name in ("tag", "label", "nonumber", "notag"):
            if self.peek() == "{":
                self._parse_group()
            return ""

        # Spacing
        if name in ("quad", "qquad"):
            return _make_run("  ")
        if name in (",", ":", ";", "!"):
            return _make_run(" ")
        if name == " ":
            return _make_run(" ")

        # Special characters
        if name == "|":
            return _make_run("\u2016")
        if name == "#":
            return _make_run("#")
        if name == "%":
            return _make_run("%")
        if name == "&":
            return _make_run("&")
        if name == "_":
            return _make_run("_")
        if name == "{":
            return _make_run("{")
        if name == "}":
            return _make_run("}")

        # Unknown command: render as-is
        return _make_run(cmd)



def latex_to_omath(latex_str, display=False, eq_number=None):
    """Convert LaTeX math expression to OMML XML string.

    Args:
        latex_str: LaTeX math expression (without $ delimiters).
        display: If True, wrap in m:oMathPara container for display math.
        eq_number: Optional equation number string (e.g. "2.1").
                   Only used when display=True.

    Returns:
        OMML XML string, or "" if conversion fails.
    """
    if not latex_str or not latex_str.strip():
        return ""

    # Pre-process non-alphabetic LaTeX commands that tokenizer can't handle
    latex_str = latex_str.replace(r'\|', '\u2016')  # norm → double vertical bar ‖
    latex_str = latex_str.replace(r'\{', '{').replace(r'\}', '}')

    try:
        latex_str = latex_str.strip()
        tokens = _tokenize_latex(latex_str)
        parser = OMMLParser(tokens)
        inner_omml = parser.parse()

        if not inner_omml:
            return ""

        if eq_number and display:
            eq_num_text = "#" + eq_number
            omml = (
                "<m:oMath>"
                + "<m:eqArr>"
                + "<m:eqArrPr>"
                + '<m:maxDist m:val="1"/>'
                + _ctrl_pr()
                + "</m:eqArrPr>"
                + "<m:e>"
                + inner_omml
                + _make_run_sty(eq_num_text)
                + "</m:e>"
                + "</m:eqArr>"
                + "</m:oMath>"
            )
        else:
            omml = "<m:oMath>" + inner_omml + "</m:oMath>"

        if display:
            omml = "<m:oMathPara>" + omml + "</m:oMathPara>"

        return omml

    except Exception:
        return ""


def _fallback_math_to_unicode(text):
    """Convert LaTeX to plain Unicode when OMML conversion fails."""
    if not text:
        return ""

    result = text
    # Fix corrupted \b in \bm (backspace character from bad JSON generation)
    result = result.replace('\x08', '\\bm')

    # Replace Greek letters (longest first to avoid partial matches)
    for cmd, unichar in sorted(GREEK_MAP.items(), key=lambda x: -len(x[0])):
        result = result.replace("\\" + cmd, unichar)

    # Replace operators
    for cmd, unichar in sorted(OP_MAP.items(), key=lambda x: -len(x[0])):
        result = result.replace("\\" + cmd, unichar)

    # Font commands (handle both \\cmd{...} and \cmd{...} forms)
    result = re.sub(r"\\{1,2}mathrm\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}mathbf\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}textbf\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}bm\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}text\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}mathit\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}mathcal\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}mathbb\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\{1,2}mathfrak\{([^}]*)\}", r"\1", result)

    # Fractions
    result = re.sub(r"\\{1,2}frac\{([^}]*)\}\{([^}]*)\}", r"\1/\2", result)

    # Subscripts/superscripts
    result = re.sub(r"\{(\w+)\}_\{(\w+)\}", r"\1_{\2}", result)
    result = re.sub(r"\{(\w+)\}\^\{(\w+)\}", r"\1^\2", result)

    # Remove remaining braces
    result = result.replace("{", "").replace("}", "")

    # Functions
    for func in sorted(FUNC_NAMES, key=lambda x: -len(x)):
        result = result.replace("\\" + func, func)

    # Remove \\left, \\right, \\big, etc.
    result = re.sub(r"\\(left|right|big[lr]?|Big[lr]?|bigg[lr]?|Bigg[lr]?)", "", result)
    result = re.sub(r"\\(displaystyle|textstyle|scriptstyle|scriptscriptstyle)", "", result)
    result = re.sub(r"\\(tag|label|nonumber|notag|limits)\\b", "", result)
    result = re.sub(r"\\(quad|qquad|\\\\,|\\\\:|\\\\;|\\\\)", " ", result)

    # Remove stray backslashes
    result = re.sub(r"\\([#%&_])", r"\1", result)

    # Clean up multiple spaces
    result = re.sub(r"  +", " ", result).strip()

    return result



def convert_inline_math(text):
    """
    Convert all $...$ patterns in text to OMML or Unicode fallback.
    Returns (converted_text, has_omml, omml_fragments) tuple.

    - converted_text: text with $...$ replaced by Unicode fallback where OMML fails
    - has_omml: True if any $...$ was successfully converted to OMML
    - omml_fragments: list of (position, omml_xml) for successful OMML conversions
    """
    # Fix corrupted \b in \bm (backspace character from bad JSON generation)
    text = text.replace('\x08', '\\bm')
    # Pre-process non-alphabetic LaTeX commands that tokenizer can't handle
    text = text.replace(r'\|', '\u2016')  # norm → double vertical bar ‖
    text = text.replace(r'\{', '{').replace(r'\}', '}')
    parts = re.split(r'(\$[^$]+\$)', text)
    result_parts = []
    has_omml = False
    omml_frags = []
    pos = 0
    for part in parts:
        if part.startswith('$') and part.endswith('$'):
            inner = part[1:-1]
            omml = latex_to_omath(inner, display=False)
            if omml:
                has_omml = True
                omml_frags.append((pos, omml))
                result_parts.append('')
            else:
                unicode_text = _fallback_math_to_unicode(inner)
                result_parts.append(unicode_text)
        else:
            result_parts.append(part)
        pos += len(part)
    return ''.join(result_parts), has_omml, omml_frags


if __name__ == "__main__":
    print("=" * 60)
    print("OMML Generator - Test Suite")
    print("=" * 60)

    test_cases = [
        ("Simple variable theta", r"\theta", False, None),
        ("Subscript x_1", "x_1", False, None),
        ("Subscript F_max", r"F_{\max}", False, None),
        ("Superscript R^n", "R^n", False, None),
        ("Fraction 1/N", r"\frac{1}{N}", False, None),
        ("Fraction N_f/N", r"\frac{N_f}{N}", False, None),
        ("Summation sum_{j=1}^{N}", r"\sum_{j=1}^{N}", False, None),
        ("Max with limit", r"\max_{0 \le t \le T}", False, None),
        ("Bold X", r"\bm{X}", False, None),
        ("Display equation with number", r"\frac{N_f}{N}", True, "2.10"),
        ("Display equation theta=f(X)", r"\theta=f(X)", True, "2.1"),
        ("Operator le", r"\le", False, None),
        ("Operator rightarrow", r"\rightarrow", False, None),
        ("Complex: P_f approx 1/N sum I_F(x_j)",
         r"P_f \approx \frac{1}{N} \sum_{j=1}^{N} I_F(x_j)", False, None),
    ]

    for desc, latex, display, eq_num in test_cases:
        print()
        print("--- " + desc + " ---")
        print("  LaTeX: $" + latex + "$")
        result = latex_to_omath(latex, display=display, eq_number=eq_num)
        if result:
            display_str = result[:200] + "..." if len(result) > 200 else result
            print("  OMML:  " + display_str)
        else:
            fallback = _fallback_math_to_unicode(latex)
            print("  OMML:  (conversion failed)")
            print("  Unicode fallback: " + fallback)

    print()
    print("=" * 60)
    print("Fallback test")
    print("=" * 60)
    fb = _fallback_math_to_unicode(r"\theta \rightarrow \infty")
    print("  Input:  \\theta \\rightarrow \\infty")
    print("  Output: " + fb)
