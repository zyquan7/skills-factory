#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Deterministic OpenXML Word (.docx) Rendering Engine for Chinese Academic Paper Translations (v3).
Fixes & Features:
1. Native Word Office Math Markup Language (OMML <m:oMath>) Engine:
   - Uses <m:d> delimiter elements for curly braces \{...\}, norms \|...\|, parentheses (...), and brackets [...]
     so zero literal backslashes (\{ or \}) ever appear in Word or LibreOffice.
   - Uses BMP Unicode math glyphs and <m:scr m:val="script"/> for calligraphic letters (\mathcal{N}, \mathcal{T})
     to avoid UTF-16 astral surrogate truncation in equation renderers.
   - Supports \frac{a}{b} (<m:f>), \sum_{sub}^{sup} (<m:nary>), \sqrt{x} (<m:rad>), subscripts (<m:sSub>),
     superscripts (<m:sSup>), combined sub-superscripts (<m:sSubSup>), and automatic bare-math detection.
2. Authentic Chinese Academic Three-Line Table (简明三线表) Engine:
   - Explicitly writes <w:tblGrid><w:gridCol w:w="..."/></w:tblGrid> in DXA (1 cm = 567 dxa) AND <w:tcW> on
     every cell (including merged spanner cells) + <w:noWrap/> on method/category columns so neither headers
     nor method names ever wrap into multi-line vertical fragments.
   - Enforces exact 1.5pt top rule (sz=12), 0.5pt spanner sub-rule (sz=4), 0.75pt header bottom rule (sz=6),
     1.5pt table bottom rule (sz=12), and nil internal/vertical borders.
"""

import os
import sys
import re
import json
import argparse
from typing import Dict, List, Optional, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls, nsmap

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
nsmap["m"] = MATH_NS

# Exact Chinese Font Size Mapping (Pt)
PT_XIAO_ER = 18.0   # 小二号 (18pt)
PT_SAN_HAO = 16.0   # 三号 (16pt)
PT_XIAO_SAN = 15.0  # 小三号 (15pt)
PT_SI_HAO = 14.0    # 四号 (14pt)
PT_XIAO_SI = 12.0   # 小四号 (12pt)
PT_WU_HAO = 10.5    # 五号 (10.5pt)

ASCII_PATTERN = re.compile(r'([\x00-\x7F]+)')
INLINE_MATH_PATTERN = re.compile(r'\$([^\$]+)\$')

# Strictly 2-byte BMP Unicode mathematical symbols (zero 4-byte astral surrogates)
LATEX_SYMBOLS: Dict[str, str] = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ϵ", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\iota": "ι", r"\kappa": "κ", r"\lambda": "λ",
    r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ", r"\pi": "π",
    r"\rho": "ρ", r"\sigma": "σ", r"\tau": "τ", r"\upsilon": "υ",
    r"\phi": "ϕ", r"\varphi": "φ", r"\chi": "χ", r"\psi": "ψ",
    r"\omega": "ω", r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ",
    r"\Lambda": "Λ", r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ",
    r"\Phi": "Φ", r"\Psi": "Ψ", r"\Omega": "Ω",
    r"\in": "∈", r"\notin": "∉", r"\subset": "⊂", r"\subseteq": "⊆",
    r"\cup": "∪", r"\cap": "∩", r"\emptyset": "∅",
    r"\ge": "≥", r"\geq": "≥", r"\le": "≤", r"\leq": "≤",
    r"\neq": "≠", r"\ne": "≠", r"\approx": "≈", r"\sim": "∼",
    r"\equiv": "≡", r"\propto": "∝", r"\times": "×", r"\cdot": "⋅",
    r"\pm": "±", r"\mp": "∓", r"\div": "÷", r"\ast": "∗",
    r"\oplus": "⊕", r"\otimes": "⊗", r"\top": "⊤", r"\bot": "⊥",
    r"\infty": "∞", r"\partial": "∂", r"\nabla": "∇", r"\forall": "∀",
    r"\exists": "∃", r"\rightarrow": "→", r"\to": "→", r"\leftarrow": "←",
    r"\Rightarrow": "⇒", r"\Leftrightarrow": "⇔", r"\mapsto": "↦",
    r"\dots": "…", r"\ldots": "…", r"\cdots": "⋯", r"\circ": "°",
    r"\ell": "ℓ",
    # BMP Double-Struck & Script letters (U+2100..U+214F)
    r"\mathbb{R}": "ℝ", r"\mathbb{N}": "ℕ", r"\mathbb{Z}": "ℤ",
    r"\mathbb{E}": "E", r"\mathbb{P}": "P",
    r"\mathcal{I}": "ℐ", r"\mathcal{L}": "ℒ", r"\mathcal{H}": "ℋ",
    r"\mathcal{N}": "N", r"\mathcal{T}": "T", r"\mathcal{P}": "P",
    r"\mathcal{O}": "O", r"\mathcal{D}": "D", r"\mathcal{X}": "X",
    r"\|": "‖", r"\Vert": "‖",
    r"\%": "%", r"\quad": "   ", r"\qquad": "     ", r"\,": " ", r"\;": " ", r"\!": "",
}


# ==============================================================================
# 1. Native Word OMML (<m:oMath>) Equation Builder Engine
# ==============================================================================
def _m_el(tag: str) -> OxmlElement:
    return parse_xml(f'<m:{tag} xmlns:m="{MATH_NS}" xmlns:w="{W_NS}"/>')


def _create_math_run(
    text: str,
    italic: bool = True,
    size_pt: float = PT_XIAO_SI,
    script_style: Optional[str] = None,
) -> OxmlElement:
    """
    Creates an <m:r> element with Cambria Math / Times New Roman + SimSun font binding.
    """
    mr = _m_el("r")
    mrPr = _m_el("rPr")
    has_mrpr = False
    if not italic and not script_style:
        mrPr.append(_m_el("nor"))
        has_mrpr = True
    if script_style:
        mscr = _m_el("scr")
        mscr.set(qn("m:val"), script_style)
        mrPr.append(mscr)
        has_mrpr = True
    if has_mrpr:
        mr.append(mrPr)

    wrPr = parse_xml(
        f'<w:rPr xmlns:w="{W_NS}">'
        f'<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math" w:eastAsia="宋体"/>'
        f'<w:sz w:val="{int(size_pt * 2)}"/>'
        f'<w:szCs w:val="{int(size_pt * 2)}"/>'
        f'</w:rPr>'
    )
    mr.append(wrPr)

    mt = _m_el("t")
    mt.text = text
    if text.startswith(" ") or text.endswith(" "):
        mt.set(qn("xml:space"), "preserve")
    mr.append(mt)
    return mr


def _create_delimiter_element(
    beg_chr: str, end_chr: str, inner_nodes: List[OxmlElement], size_pt: float = PT_XIAO_SI
) -> OxmlElement:
    r"""
    Creates a native OMML <m:d> delimiter element for \{...\}, (...), [...], or \|...\|
    so curly braces and norms render as proper mathematical delimiters without backslashes.
    """
    md = _m_el("d")
    mdPr = _m_el("dPr")
    mbeg = _m_el("begChr")
    mbeg.set(qn("m:val"), beg_chr)
    mend = _m_el("endChr")
    mend.set(qn("m:val"), end_chr)
    mdPr.append(mbeg)
    mdPr.append(mend)
    md.append(mdPr)
    me = _m_el("e")
    for node in inner_nodes:
        me.append(node)
    md.append(me)
    return md


def _extract_braced_group(s: str, pos: int) -> Tuple[str, int]:
    """Extracts a {...} group or a single token/compound math structure starting at pos in string s."""
    while pos < len(s) and s[pos] == " ":
        pos += 1
    if pos >= len(s):
        return "", pos
    if s[pos] == "{":
        depth = 0
        start = pos + 1
        i = pos
        while i < len(s):
            if s[i] == "\\" and i + 1 < len(s) and s[i + 1] in ("{", "}"):
                i += 2
                continue
            if s[i] == "{":
                depth += 1
            elif s[i] == "}":
                depth -= 1
                if depth == 0:
                    return s[start:i], i + 1
            i += 1
        return s[start:], len(s)
    elif s[pos] == "\\":
        m = re.match(r"^(\\[a-zA-Z]+)", s[pos:])
        if m:
            cmd = m.group(1)
            next_p = pos + len(cmd)
            # 1. 字体与运算符命令：\mathrm, \text, \operatorname, \mathbf, \mathcal, \mathbb
            if cmd in (r"\mathrm", r"\text", r"\operatorname", r"\mathbf", r"\mathcal", r"\mathbb") and next_p < len(s) and s[next_p] == "{":
                inner, end_p = _extract_braced_group(s, next_p)
                return f"{cmd}{{{inner}}}", end_p
            # 2. 分式命令：\frac{num}{den} 完整提取两组参数，防止被加项出现空分式槽位
            if cmd == r"\frac":
                num_s, p1 = _extract_braced_group(s, next_p)
                den_s, p2 = _extract_braced_group(s, p1)
                return f"\\frac{{{num_s}}}{{{den_s}}}", p2
            # 3. 根号与上标重音：\sqrt, \tilde, \hat, \bar 等提取单个参数组
            if cmd in (r"\sqrt", r"\tilde", r"\widetilde", r"\hat", r"\widehat", r"\bar", r"\overline", r"\vec", r"\dot"):
                arg_s, p1 = _extract_braced_group(s, next_p)
                return f"{cmd}{{{arg_s}}}", p1
            # 4. 大型算子：\sum, \prod, \int 等递归提取其上下标与被运算项
            if cmd in (r"\sum", r"\prod", r"\int", r"\coprod", r"\bigcap", r"\bigcup"):
                cur = next_p
                parts = [cmd]
                while cur < len(s) and s[cur] in (" ", "_", "^"):
                    if s[cur] == " ":
                        cur += 1
                    else:
                        op = s[cur]
                        lim_s, cur = _extract_braced_group(s, cur + 1)
                        parts.append(f"{op}{{{lim_s}}}")
                body_s, cur = _extract_braced_group(s, cur)
                if body_s:
                    parts.append(f"{{{body_s}}}")
                return "".join(parts), cur
            return cmd, next_p
        return s[pos:pos + 2], pos + 2
    elif s[pos] == "(":
        close_p = _find_matching_delim(s, pos, "(", ")")
        if close_p != -1:
            return s[pos:close_p + 1], close_p + 1
        return s[pos], pos + 1
    else:
        return s[pos], pos + 1


def _find_matching_delim(s: str, start_idx: int, open_tok: str, close_tok: str) -> int:
    """Finds the matching closing delimiter index in string s."""
    depth = 0
    i = start_idx
    while i < len(s):
        if s.startswith(open_tok, i):
            depth += 1
            i += len(open_tok)
            continue
        if s.startswith(close_tok, i):
            depth -= 1
            if depth == 0:
                return i
            i += len(close_tok)
            continue
        i += 1
    return -1


def _replace_latex_symbols(s: str) -> str:
    out = s
    for k in sorted(LATEX_SYMBOLS.keys(), key=len, reverse=True):
        out = out.replace(k, LATEX_SYMBOLS[k])
    return out


def _attach_sub_sup_if_present(
    expr: str, i: int, base_nodes: List[OxmlElement], size_pt: float
) -> Tuple[List[OxmlElement], int]:
    """Checks if current position i has _sub and/or ^sup and wraps base_nodes in <m:sSub>/<m:sSup>/<m:sSubSup>."""
    n = len(expr)
    sub_str, sup_str = None, None
    while i < n and expr[i] in ("_", "^"):
        if expr[i] == "_":
            sub_str, i = _extract_braced_group(expr, i + 1)
        elif expr[i] == "^":
            sup_str, i = _extract_braced_group(expr, i + 1)

    if sub_str is not None and sup_str is not None:
        ms = _m_el("sSubSup")
        me = _m_el("e")
        for bn in base_nodes:
            me.append(bn)
        msub = _m_el("sub")
        for sn in _build_omml_nodes(sub_str, size_pt=max(8.0, size_pt - 2.0)):
            msub.append(sn)
        msup = _m_el("sup")
        for spn in _build_omml_nodes(sup_str, size_pt=max(8.0, size_pt - 2.0)):
            msup.append(spn)
        ms.append(me)
        ms.append(msub)
        ms.append(msup)
        return [ms], i
    elif sub_str is not None:
        ms = _m_el("sSub")
        me = _m_el("e")
        for bn in base_nodes:
            me.append(bn)
        msub = _m_el("sub")
        for sn in _build_omml_nodes(sub_str, size_pt=max(8.0, size_pt - 2.0)):
            msub.append(sn)
        ms.append(me)
        ms.append(msub)
        return [ms], i
    elif sup_str is not None:
        ms = _m_el("sSup")
        me = _m_el("e")
        for bn in base_nodes:
            me.append(bn)
        msup = _m_el("sup")
        for spn in _build_omml_nodes(sup_str, size_pt=max(8.0, size_pt - 2.0)):
            msup.append(spn)
        ms.append(me)
        ms.append(msup)
        return [ms], i
    return base_nodes, i


def _build_omml_nodes(expr: str, size_pt: float = PT_XIAO_SI, display_mode: bool = False) -> List[OxmlElement]:
    expr = expr.strip()
    nodes: List[OxmlElement] = []
    i = 0
    n = len(expr)

    while i < n:
        if expr[i].isspace():
            i += 1
            continue

        # 0. Handle \{ ... \} set braces via native <m:d> delimiter element
        if expr.startswith(r"\{", i):
            close_pos = _find_matching_delim(expr, i, r"\{", r"\}")
            if close_pos != -1:
                inner_str = expr[i + 2:close_pos]
                i = close_pos + 2
                inner_nodes = _build_omml_nodes(inner_str, size_pt=size_pt, display_mode=display_mode)
                delim_el = _create_delimiter_element("{", "}", inner_nodes, size_pt=size_pt)
                wrapped, i = _attach_sub_sup_if_present(expr, i, [delim_el], size_pt)
                nodes.extend(wrapped)
                continue

        # Handle \| ... \| norm bars via native <m:d> delimiter element
        if expr.startswith(r"\|", i):
            close_pos = expr.find(r"\|", i + 2)
            if close_pos != -1:
                inner_str = expr[i + 2:close_pos]
                i = close_pos + 2
                inner_nodes = _build_omml_nodes(inner_str, size_pt=size_pt, display_mode=display_mode)
                delim_el = _create_delimiter_element("‖", "‖", inner_nodes, size_pt=size_pt)
                wrapped, i = _attach_sub_sup_if_present(expr, i, [delim_el], size_pt)
                nodes.extend(wrapped)
                continue

        # 1. Handle \frac{num}{den}
        if expr.startswith(r"\frac", i):
            i += len(r"\frac")
            num_str, i = _extract_braced_group(expr, i)
            den_str, i = _extract_braced_group(expr, i)
            mf = _m_el("f")
            mnum = _m_el("num")
            mden = _m_el("den")
            for child in _build_omml_nodes(num_str, size_pt=size_pt, display_mode=display_mode):
                mnum.append(child)
            for child in _build_omml_nodes(den_str, size_pt=size_pt, display_mode=display_mode):
                mden.append(child)
            mf.append(mnum)
            mf.append(mden)
            nodes.append(mf)
            continue

        # 2. Handle \sum_{sub}^{sup} or \sum_{sub}
        if expr.startswith(r"\sum", i):
            i += len(r"\sum")
            sub_str, sup_str = "", ""
            while i < n and expr[i] in (" ", "_", "^"):
                if expr[i] == " ":
                    i += 1
                elif expr[i] == "_":
                    sub_str, i = _extract_braced_group(expr, i + 1)
                elif expr[i] == "^":
                    sup_str, i = _extract_braced_group(expr, i + 1)

            mnary = _m_el("nary")
            mnaryPr = _m_el("naryPr")
            mchr = _m_el("chr")
            mchr.set(qn("m:val"), "∑")
            mlimLoc = _m_el("limLoc")
            mlimLoc.set(qn("m:val"), "undOvr" if display_mode else "subSup")
            mnaryPr.append(mchr)
            mnaryPr.append(mlimLoc)
            if not sup_str:
                msupHide = _m_el("supHide")
                msupHide.set(qn("m:val"), "1")
                mnaryPr.append(msupHide)
            if not sub_str:
                msubHide = _m_el("subHide")
                msubHide.set(qn("m:val"), "1")
                mnaryPr.append(msubHide)
            mnary.append(mnaryPr)

            msub = _m_el("sub")
            for child in _build_omml_nodes(sub_str, size_pt=max(8.0, size_pt - 2.0)):
                msub.append(child)
            msup = _m_el("sup")
            for child in _build_omml_nodes(sup_str, size_pt=max(8.0, size_pt - 2.0)):
                msup.append(child)
            me = _m_el("e")

            while i < n and expr[i].isspace():
                i += 1
            if i < n and expr[i] == "{":
                summand_str, i = _extract_braced_group(expr, i)
            else:
                start_sm = i
                paren_depth = 0
                brace_depth = 0
                while i < n:
                    ch = expr[i]
                    if ch in "({[":
                        paren_depth += 1
                        i += 1
                    elif ch in ")}]":
                        if paren_depth > 0:
                            paren_depth -= 1
                            i += 1
                        else:
                            break
                    elif paren_depth == 0 and brace_depth == 0:
                        if ch in ("+", "=", ","):
                            break
                        if ch in ("-", "−") and i > start_sm:
                            break
                        if expr.startswith((r"\quad", r"\qquad"), i):
                            break
                        i += 1
                    else:
                        i += 1
                summand_str = expr[start_sm:i].strip()

            summand_nodes = _build_omml_nodes(summand_str, size_pt=size_pt) if summand_str else []
            if not summand_nodes:
                summand_nodes = [_create_math_run("", size_pt=size_pt)]
            for child in summand_nodes:
                me.append(child)

            mnary.append(msub)
            mnary.append(msup)
            mnary.append(me)
            nodes.append(mnary)
            continue

        # 3. Handle \sqrt{...}
        if expr.startswith(r"\sqrt", i):
            i += len(r"\sqrt")
            arg_str, i = _extract_braced_group(expr, i)
            mrad = _m_el("rad")
            mradPr = _m_el("radPr")
            mdegHide = _m_el("degHide")
            mdegHide.set(qn("m:val"), "1")
            mradPr.append(mdegHide)
            mdeg = _m_el("deg")
            me = _m_el("e")
            for child in _build_omml_nodes(arg_str, size_pt=size_pt):
                me.append(child)
            mrad.append(mradPr)
            mrad.append(mdeg)
            mrad.append(me)
            nodes.append(mrad)
            continue

        # 3b. Handle math accents \tilde{...}, \hat{...}, \bar{...}, \overline{...}, \vec{...}, \dot{...}
        accent_map = {
            r"\tilde": "̃", r"\widetilde": "̃",
            r"\hat": "̂", r"\widehat": "̂",
            r"\bar": "̄", r"\overline": "̅",
            r"\vec": "⃗", r"\dot": "̇",
        }
        matched_acc = None
        for acc_cmd, acc_chr in sorted(accent_map.keys() if False else accent_map.items(), key=lambda kv: len(kv[0]), reverse=True):
            if expr.startswith(acc_cmd, i) and (i + len(acc_cmd) == n or not expr[i + len(acc_cmd)].isalpha()):
                matched_acc = (acc_cmd, acc_chr)
                break
        if matched_acc is not None:
            acc_cmd, acc_chr = matched_acc
            i += len(acc_cmd)
            arg_str, i = _extract_braced_group(expr, i)
            macc = _m_el("acc")
            maccPr = _m_el("accPr")
            mchr = _m_el("chr")
            mchr.set(qn("m:val"), acc_chr)
            maccPr.append(mchr)
            me = _m_el("e")
            for child in _build_omml_nodes(arg_str, size_pt=size_pt):
                me.append(child)
            macc.append(maccPr)
            macc.append(me)
            wrapped, i = _attach_sub_sup_if_present(expr, i, [macc], size_pt)
            nodes.extend(wrapped)
            continue

        # 4. Extract next base atom
        is_upright = False
        if expr.startswith((r"\mathrm", r"\text", r"\operatorname", r"\mathbf"), i):
            m = re.match(r"^(\\[a-zA-Z]+)", expr[i:])
            cmd = m.group(1)
            i += len(cmd)
            base_str, i = _extract_braced_group(expr, i)
            is_upright = True
        elif expr[i] == "{":
            base_str, i = _extract_braced_group(expr, i)
        elif expr[i] == "\\":
            matched_sym = None
            for k in sorted(LATEX_SYMBOLS.keys(), key=len, reverse=True):
                if expr.startswith(k, i):
                    matched_sym = LATEX_SYMBOLS[k]
                    i += len(k)
                    break
            if matched_sym is not None:
                base_str = matched_sym
                is_upright = not matched_sym.isalpha()
            else:
                m = re.match(r"^(\\[a-zA-Z]+)", expr[i:])
                if m:
                    cmd = m.group(1)
                    i += len(cmd)
                    replaced = _replace_latex_symbols(cmd)
                    base_str = replaced.lstrip("\\") if replaced.startswith("\\") else replaced
                    is_upright = True
                else:
                    base_str = expr[i]
                    i += 1
                    is_upright = True
        else:
            m_fn = re.match(r"^(mean|min|max|exp|log|sin|cos|Norm|TVaR|IQR|PRO|AUROC|AUPR|PCA|SVD|FPR|span)\b", expr[i:])
            if m_fn:
                base_str = m_fn.group(1)
                i += len(base_str)
                is_upright = True
            elif expr[i].isdigit():
                m_num = re.match(r"^[0-9]+(?:\.[0-9]+)?%?", expr[i:])
                base_str = m_num.group(0)
                i += len(base_str)
                is_upright = True
            else:
                base_str = expr[i]
                i += 1
                if not base_str.isalpha():
                    is_upright = True

        if is_upright:
            clean_base = _replace_latex_symbols(base_str)
            if clean_base in ("=", "∈", "≥", "≤", "≈", "∼", "→", "+", "−", "-"):
                clean_base = f" {clean_base} "
            base_nodes = [_create_math_run(clean_base, italic=False, size_pt=size_pt)]
        else:
            clean_base = _replace_latex_symbols(base_str)
            if len(clean_base) == 1 or base_str == expr:
                base_nodes = [_create_math_run(clean_base, italic=True, size_pt=size_pt)]
            else:
                base_nodes = _build_omml_nodes(base_str, size_pt=size_pt)

        wrapped, i = _attach_sub_sup_if_present(expr, i, base_nodes, size_pt)
        nodes.extend(wrapped)

    return nodes


def create_omath_element(latex_expr: str, size_pt: float = PT_XIAO_SI, display_mode: bool = False) -> OxmlElement:
    s = latex_expr.strip().strip("$")
    # 1. 清理定界符控制词与间距
    s = s.replace(r"\left", "").replace(r"\right", "").replace(r"\middle|", "|").replace(r"\middle", "")
    s = s.replace(r"\%", "%").replace(r"\;", " ").replace(r"\,", " ").replace(r"\!", "")
    s = s.replace(r"\lceil", "⌈").replace(r"\rceil", "⌉").replace(r"\langle", "⟨").replace(r"\rangle", "⟩")

    # 2. 转换常见数学字体与宏
    s = re.sub(r"\\mathcal\{([^\}]+)\}", lambda m: LATEX_SYMBOLS.get(m.group(0), m.group(1)), s)
    s = re.sub(r"\\mathbb\{([^\}]+)\}", lambda m: LATEX_SYMBOLS.get(m.group(0), m.group(1)), s)
    s = re.sub(r"\\text\{([^\}]+)\}", r"\\mathrm{\1}", s)
    s = re.sub(r"\\operatorname\{([^\}]+)\}", r"\\mathrm{\1}", s)

    # 3. 未知转义降级（保留已识别命令，其余剥离反斜杠避免解析断裂）
    known_cmds = set(LATEX_SYMBOLS.keys()) | {
        r"\frac", r"\sum", r"\prod", r"\int", r"\coprod", r"\bigcap", r"\bigcup",
        r"\sqrt", r"\mathrm", r"\mathbf",
        r"\tilde", r"\widetilde", r"\hat", r"\widehat", r"\bar", r"\overline", r"\vec", r"\dot",
        r"\quad", r"\qquad",
    }
    s = re.sub(r"\\[a-zA-Z]+", lambda m: m.group(0) if (m.group(0) in known_cmds) else m.group(0)[1:], s)

    omath = _m_el("oMath")
    for child in _build_omml_nodes(s, size_pt=size_pt, display_mode=display_mode):
        omath.append(child)

    # 4. 行间公式优化：非分式内部的大型求和号上下限统一设置为上下叠放 (undOvr)
    if display_mode:
        f_nodes = set()
        for f_el in omath.findall(f".//{{{MATH_NS}}}f"):
            for desc in f_el.iter():
                f_nodes.add(desc)
        for nary in omath.findall(f".//{{{MATH_NS}}}nary"):
            naryPr = nary.find(f"{{{MATH_NS}}}naryPr")
            if naryPr is not None and nary not in f_nodes:
                limLoc = naryPr.find(f"{{{MATH_NS}}}limLoc")
                if limLoc is not None:
                    limLoc.set(qn("m:val"), "undOvr")
    return omath


# ==============================================================================
# 2. Mixed Chinese/English + Inline Native Math Run Splitter
# ==============================================================================
def set_run_fonts(
    run,
    cn_font: str = "宋体",
    en_font: str = "Times New Roman",
    size_pt: float = PT_XIAO_SI,
    bold: bool = False,
    italic: bool = False,
    color_rgb: Tuple[int, int, int] = (0, 0, 0),
) -> None:
    run.font.name = en_font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:cs"), en_font)
    rFonts.set(qn("w:eastAsia"), cn_font)

    # 显式绑定 zh-CN 语言与 en-US 英文，杜绝 Word 默认回退为 ja-JP 触发大面积假性红波浪线
    for old_lang in rPr.findall(qn("w:lang")):
        rPr.remove(old_lang)
    rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="en-US" w:eastAsia="zh-CN"/>'))


BARE_SUB_SUP_PATTERN = re.compile(
    r'(?<![a-zA-Z0-9_])([A-Za-zα-ωΑ-Ω\\]+(?:_\{[^\}]+\}|_[A-Za-z0-9α-ω]+|\^\{[^\}]+\}|\^[A-Za-z0-9⊤\+]+)+(?:\([A-Za-z0-9,\s_]+\))?)'
)


def _auto_wrap_bare_math(text: str) -> str:
    segments = INLINE_MATH_PATTERN.split(text)
    out = []
    for idx, seg in enumerate(segments):
        if idx % 2 == 1:
            out.append(f"${seg}$")
        else:
            def _repl(m):
                tok = m.group(1)
                if tok.lower().startswith(("mvtec", "visa", "dinov2", "subspacead", "table", "fig")):
                    return tok
                tok = re.sub(r'_([a-zA-Z]{2,})', r'_{\\mathrm{\1}}', tok)
                return f"${tok}$"
            out.append(BARE_SUB_SUP_PATTERN.sub(_repl, seg))
    return "".join(out)


def add_mixed_runs(
    paragraph,
    text: str,
    cn_font: str = "宋体",
    en_font: str = "Times New Roman",
    size_pt: float = PT_XIAO_SI,
    bold: bool = False,
    italic: bool = False,
) -> None:
    if not text:
        return

    processed_text = _auto_wrap_bare_math(text)
    math_split = INLINE_MATH_PATTERN.split(processed_text)

    for idx, chunk in enumerate(math_split):
        if not chunk:
            continue
        if idx % 2 == 1:
            omath = create_omath_element(chunk, size_pt=size_pt, display_mode=False)
            paragraph._p.append(omath)
        else:
            parts = ASCII_PATTERN.split(chunk)
            for part in parts:
                if not part:
                    continue
                r = paragraph.add_run(part)
                set_run_fonts(
                    r, cn_font=cn_font, en_font=en_font,
                    size_pt=size_pt, bold=bold, italic=italic
                )


def set_outline_level(p, outline_level: Optional[int] = None) -> None:
    """
    Sets or clears the Word OpenXML paragraph outline level (<w:pPr><w:outlineLvl w:val="N"/></w:pPr>).
    outline_level is 1-based (1 -> Level 1 / w:val="0", 2 -> Level 2 / w:val="1", ..., 9 -> Level 9 / w:val="8").
    When outline_level is None, removes any <w:outlineLvl> tag so the paragraph remains Body Text (正文文本).
    """
    pPr = p._p.get_or_add_pPr()
    for old_el in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(old_el)
    if outline_level is not None:
        lvl_idx = max(0, min(8, int(outline_level) - 1))
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{lvl_idx}"/>'))


def get_outline_level(p) -> Optional[int]:
    """
    Returns the 1-based Word outline level (1..9) of paragraph p if <w:outlineLvl> is set, else None.
    """
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        return None
    val = ol.get(qn("w:val"))
    if val is None or not str(val).isdigit():
        return None
    lvl_idx = int(val)
    return (lvl_idx + 1) if 0 <= lvl_idx <= 8 else None


def format_paragraph(
    p,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    line_spacing: float = 1.5,
    first_line_indent_pt: float = 0.0,
    left_indent_pt: float = 0.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    keep_with_next: bool = False,
    outline_level: Optional[int] = None,
) -> None:
    pf = p.paragraph_format
    p.alignment = align
    pf.line_spacing = line_spacing
    pf.first_line_indent = Pt(first_line_indent_pt) if first_line_indent_pt > 0 else Pt(0)
    pf.left_indent = Pt(left_indent_pt) if left_indent_pt > 0 else Pt(0)
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.keep_with_next = keep_with_next
    pf.widow_control = True
    set_outline_level(p, outline_level)


# ==============================================================================
# 3. Rule 2.12: Native OMML Display Equation + Flush-Right Equation Number
# ==============================================================================
def _set_table_grid_and_widths(tbl, col_widths_cm: List[float]) -> List[int]:
    """
    Sets <w:tblGrid><w:gridCol w:w="..."/></w:tblGrid> in DXA (1 cm = 567 dxa)
    so Word, WPS, and LibreOffice strictly honor exact column widths!
    """
    dxa_widths = [int(round(w_cm * 567.0)) for w_cm in col_widths_cm]
    tblGrid = tbl._tbl.tblGrid
    if tblGrid is not None:
        for old_col in list(tblGrid):
            tblGrid.remove(old_col)
        for dxa in dxa_widths:
            gc = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{dxa}"/>')
            tblGrid.append(gc)
    return dxa_widths


def _set_cell_dxa_width(cell, dxa_w: int, no_wrap: bool = False) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old_w in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old_w)
    tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{dxa_w}" w:type="dxa"/>'))
    if no_wrap:
        for old_nw in tcPr.findall(qn("w:noWrap")):
            tcPr.remove(old_nw)
        tcPr.append(parse_xml(f'<w:noWrap {nsdecls("w")}/>'))


def add_numbered_equation(doc, formula_latex: str, eq_number: str) -> None:
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblStyle", "w:tblW", "w:tblBorders", "w:tblLayout"):
        for old_el in tblPr.findall(qn(tag)):
            tblPr.remove(old_el)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="9072" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        '</w:tblBorders>'
    ))

    # 拓宽中间公式列至 13.8 cm（两侧各 1.1 cm 对称居中，总宽 16.0 cm / 9072 dxa），保证公式与编号单行排版
    col_widths_cm = [1.1, 13.8, 1.1]
    dxa_widths = _set_table_grid_and_widths(tbl, col_widths_cm)
    row = tbl.rows[0]
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for idx, dxa_w in enumerate(dxa_widths):
        cell = row.cells[idx]
        _set_cell_dxa_width(cell, dxa_w)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        tcPr = cell._tc.get_or_add_tcPr()
        for old_tb in tcPr.findall(qn("w:tcBorders")):
            tcPr.remove(old_tb)
        tcPr.append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
            '</w:tcBorders>'
        ))
        # 消除单元格默认左右内边距，释放最大物理净宽度
        for old_mar in tcPr.findall(qn("w:tcMar")):
            tcPr.remove(old_mar)
        tcPr.append(parse_xml(
            f'<w:tcMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>'
        ))

    p_left = row.cells[0].paragraphs[0]
    format_paragraph(p_left, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.15, space_before_pt=3, space_after_pt=3)

    # Center cell: Split multi-clause equations separated by \quad into separate <m:oMath> blocks separated by spaces
    p_eq = row.cells[1].paragraphs[0]
    format_paragraph(p_eq, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15, space_before_pt=4, space_after_pt=4)
    clean_latex = formula_latex.strip("$").strip()
    eq_pt = 11.0 if len(clean_latex) > 115 else PT_XIAO_SI
    clauses = [c.strip() for c in re.split(r'\\qquad|\\quad', clean_latex) if c.strip()]
    for c_idx, clause in enumerate(clauses):
        omath = create_omath_element(clause, size_pt=eq_pt, display_mode=True)
        p_eq._p.append(omath)
        if c_idx < len(clauses) - 1:
            r_sp = p_eq.add_run("   ")
            set_run_fonts(r_sp, cn_font="宋体", en_font="Times New Roman", size_pt=eq_pt)

    p_num = row.cells[2].paragraphs[0]
    format_paragraph(p_num, align=WD_ALIGN_PARAGRAPH.RIGHT, line_spacing=1.15, space_before_pt=4, space_after_pt=4)
    num_str = eq_number if eq_number.startswith("(") else f"({eq_number})"
    add_mixed_runs(p_num, num_str, cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, italic=False)


# ==============================================================================
# 4. Rule 2.11: Authentic Chinese Academic Three-Line Table (简明三线表)
# ==============================================================================
def _set_cell_borders(
    cell,
    top: Optional[Tuple[str, int]] = None,
    bottom: Optional[Tuple[str, int]] = None,
) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)

    top_xml = (
        f'<w:top w:val="{top[0]}" w:sz="{top[1]}" w:space="0" w:color="000000"/>'
        if top else '<w:top w:val="nil"/>'
    )
    bot_xml = (
        f'<w:bottom w:val="{bottom[0]}" w:sz="{bottom[1]}" w:space="0" w:color="000000"/>'
        if bottom else '<w:bottom w:val="nil"/>'
    )
    borders_xml = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'{top_xml}'
        '<w:left w:val="nil"/>'
        f'{bot_xml}'
        '<w:right w:val="nil"/>'
        '</w:tcBorders>'
    )
    tcPr.append(borders_xml)


def add_three_line_table(
    doc,
    table_num: str,
    table_title: str,
    headers: List[str],
    rows_data: List[List[str]],
    group_headers: Optional[List[Tuple[str, int]]] = None,
    footnote: Optional[str] = None,
) -> None:
    clean_title = table_title.rstrip("。.：:；;")
    cap_p = doc.add_paragraph()
    format_paragraph(
        cap_p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5,
        space_before_pt=8, space_after_pt=3, keep_with_next=True
    )
    add_mixed_runs(cap_p, f"{table_num} {clean_title}", cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)

    num_cols = len(headers)
    total_header_rows = 2 if group_headers else 1
    tbl = doc.add_table(rows=total_header_rows, cols=num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # Configure Table-Level Properties: Fixed layout, 16.0 cm (9072 dxa) full width, three-line outer borders
    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblStyle", "w:tblW", "w:tblBorders", "w:tblCellMar", "w:tblLayout"):
        for old_el in tblPr.findall(qn(tag)):
            tblPr.remove(old_el)

    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="9072" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        '<w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/>'
        '<w:insideV w:val="nil"/>'
        '</w:tblBorders>'
    ))
    tblPr.append(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        '<w:top w:w="45" w:type="dxa"/>'
        '<w:left w:w="45" w:type="dxa"/>'
        '<w:bottom w:w="45" w:type="dxa"/>'
        '<w:right w:w="45" w:type="dxa"/>'
        '</w:tblCellMar>'
    ))

    # Compute smart proportional column widths summing to 16.0 cm
    col_weights = []
    for c_idx in range(num_cols):
        hdr_clean = str(headers[c_idx]).replace("$", "").replace("\\", "")
        max_len = sum(1.85 if ord(ch) > 127 else 1.0 for ch in hdr_clean)
        for r_vals in rows_data:
            if c_idx < len(r_vals):
                val_clean = str(r_vals[c_idx]).replace("*", "").replace("$", "").replace("\\", "")
                c_len = sum(1.85 if ord(ch) > 127 else 1.0 for ch in val_clean)
                max_len = max(max_len, c_len)
        col_weights.append(max(4.5, float(max_len) + 1.8))

    total_w_units = sum(col_weights)
    col_widths_cm = [16.0 * (w / total_w_units) for w in col_weights]
    dxa_widths = _set_table_grid_and_widths(tbl, col_widths_cm)

    # Optional Row 0: Grouped Spanner Header Row (e.g. MVTec-AD / VisA)
    if group_headers:
        grp_row = tbl.rows[0]
        grp_trPr = grp_row._tr.get_or_add_trPr()
        grp_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        grp_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        col_cursor = 0
        for grp_label, span_count in group_headers:
            start_cell = grp_row.cells[col_cursor]
            end_cell = grp_row.cells[col_cursor + span_count - 1]
            span_dxa = sum(dxa_widths[col_cursor:col_cursor + span_count])
            has_sub_rule = bool(grp_label.strip()) and span_count > 1

            for c_i in range(col_cursor, col_cursor + span_count):
                _set_cell_dxa_width(grp_row.cells[c_i], dxa_widths[c_i])
                _set_cell_borders(
                    grp_row.cells[c_i],
                    top=("single", 12),
                    bottom=("single", 4) if has_sub_rule else None,
                )
            if span_count > 1:
                start_cell.merge(end_cell)
                _set_cell_dxa_width(start_cell, span_dxa)
                _set_cell_borders(
                    start_cell,
                    top=("single", 12),
                    bottom=("single", 4) if has_sub_rule else None,
                )

            start_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = start_cell.paragraphs[0]
            format_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15, space_before_pt=2, space_after_pt=2, keep_with_next=True)
            add_mixed_runs(p, grp_label, cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO, bold=True)
            col_cursor += span_count

    # Main Column Header Row (with 0.75pt bottom border '栏目线')
    hdr_row = tbl.rows[total_header_rows - 1]
    hdr_trPr = hdr_row._tr.get_or_add_trPr()
    hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    hdr_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for c_idx, h_text in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        _set_cell_dxa_width(cell, dxa_widths[c_idx], no_wrap=True)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_borders(
            cell,
            top=("single", 12) if not group_headers else None,
            bottom=("single", 6),
        )
        p = cell.paragraphs[0]
        format_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15, space_before_pt=2, space_after_pt=2, keep_with_next=True)
        add_mixed_runs(p, str(h_text), cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO, bold=True)

    # Data Rows (supports __GROUP__ sub-header rows and standard data rows; last row has 1.5pt bottom border)
    n_rows = len(rows_data)
    for r_idx, row_vals in enumerate(rows_data):
        row = tbl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_last = (r_idx == n_rows - 1)

        if len(row_vals) >= 2 and str(row_vals[0]) == "__GROUP__":
            grp_title = str(row_vals[1])
            for c_i in range(num_cols):
                _set_cell_dxa_width(row.cells[c_i], dxa_widths[c_i])
                _set_cell_borders(
                    row.cells[c_i],
                    top=("single", 4) if r_idx > 0 else None,
                    bottom=("single", 4),
                )
            start_c = row.cells[0]
            start_c.merge(row.cells[num_cols - 1])
            _set_cell_dxa_width(start_c, sum(dxa_widths))
            _set_cell_borders(
                start_c,
                top=("single", 4) if r_idx > 0 else None,
                bottom=("single", 4),
            )
            p = start_c.paragraphs[0]
            format_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.15, space_before_pt=2, space_after_pt=2, keep_with_next=True)
            add_mixed_runs(p, grp_title, cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO, bold=True, italic=True)
            continue

        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            _set_cell_dxa_width(cell, dxa_widths[c_idx], no_wrap=True)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_borders(
                cell,
                top=None,
                bottom=("single", 12) if is_last else None,
            )
            p = cell.paragraphs[0]
            format_paragraph(
                p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15,
                space_before_pt=1.5, space_after_pt=1.5,
                keep_with_next=(n_rows <= 6 and not is_last)
            )
            val_str = str(val)
            is_bold = val_str.startswith("**") and val_str.endswith("**")
            clean_val = val_str.strip("*")
            add_mixed_runs(p, clean_val, cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO, bold=is_bold)

    if footnote:
        fn_p = doc.add_paragraph()
        format_paragraph(fn_p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.25, space_before_pt=2, space_after_pt=6)
        add_mixed_runs(fn_p, footnote, cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)


def add_figure_with_caption(
    doc,
    image_path: str,
    fig_num: str,
    fig_title: str,
    subfig_notes: Optional[List[str]] = None,
    width_inches: float = 5.4,
) -> None:
    if image_path and os.path.isfile(image_path):
        img_p = doc.add_paragraph()
        format_paragraph(
            img_p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.15,
            space_before_pt=6, space_after_pt=3, keep_with_next=True
        )
        img_p.add_run().add_picture(image_path, width=Inches(width_inches))

    clean_title = fig_title.rstrip("。.：:；;")
    cap_p = doc.add_paragraph()
    format_paragraph(
        cap_p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5,
        space_before_pt=2, space_after_pt=2 if subfig_notes else 6
    )
    add_mixed_runs(cap_p, f"{fig_num} {clean_title}", cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)

    if subfig_notes:
        for idx, note_line in enumerate(subfig_notes):
            np_p = doc.add_paragraph()
            format_paragraph(
                np_p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.35,
                space_before_pt=0, space_after_pt=6 if idx == len(subfig_notes) - 1 else 1
            )
            add_mixed_runs(np_p, note_line, cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)


def _has_page_break(p) -> bool:
    if p._p.find(f".//{{{W_NS}}}sectPr") is not None:
        return True
    for br in p._p.findall(f".//{{{W_NS}}}br"):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _ensure_page_break_before(doc) -> None:
    if len(doc.paragraphs) > 0 and not _has_page_break(doc.paragraphs[-1]):
        doc.add_page_break()


def _attach_bookmark(p, bm_id: int, bm_name: str) -> None:
    bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bm_name}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(bm_start)
    else:
        p._p.insert(0, bm_start)
    p._p.append(bm_end)


def _set_right_dot_tab(p, pos_dxa: int = 9072) -> None:
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn("w:tabs")):
        pPr.remove(old_tabs)
    pPr.append(parse_xml(
        f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="{pos_dxa}"/></w:tabs>'
    ))


def _add_fld_char(p, fld_type: str, dirty: bool = False) -> None:
    dirty_attr = ' w:dirty="true"' if dirty else ""
    p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="{fld_type}"{dirty_attr}/></w:r>'))


def _add_instr_text(p, instr: str) -> None:
    r = parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"></w:instrText></w:r>')
    r.find(qn("w:instrText")).text = instr
    p._p.append(r)


def _add_tab_run(p, cn_font: str = "宋体", en_font: str = "Times New Roman", size_pt: float = PT_XIAO_SI) -> None:
    sz_val = int(size_pt * 2)
    p._p.append(parse_xml(
        f'<w:r {nsdecls("w")}>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{en_font}" w:hAnsi="{en_font}" w:cs="{en_font}" w:eastAsia="{cn_font}"/>'
        f'<w:sz w:val="{sz_val}"/><w:szCs w:val="{sz_val}"/>'
        f'</w:rPr>'
        f'<w:tab/>'
        f'</w:r>'
    ))


def _configure_toc_styles_and_footer(doc, auto_update_fields: bool = False) -> None:
    styles_el = doc.styles.element
    for lvl, (cn_font, sz_half, indent_dxa) in {1: ("黑体", 28, 0), 2: ("宋体", 24, 480), 3: ("宋体", 24, 960)}.items():
        style_id = f"TOC{lvl}"
        style_name = f"toc {lvl}"
        for old in list(styles_el.findall(qn("w:style"))):
            name_el = old.find(qn("w:name"))
            if old.get(qn("w:styleId")) in (style_id, f"TOC {lvl}") or (
                name_el is not None and name_el.get(qn("w:val"), "").lower() == style_name
            ):
                styles_el.remove(old)
        styles_el.append(parse_xml(
            f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="{style_id}">'
            f'<w:name w:val="{style_name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
            f'<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
            f'<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9072"/></w:tabs>'
            f'<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'<w:ind w:left="{indent_dxa}" w:firstLine="0"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="{cn_font}"/>'
            f'<w:lang w:val="en-US" w:eastAsia="zh-CN"/>'
            f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr></w:style>'
        ))

    settings_el = doc.settings.element
    # 纠正主题语言为简体中文 zh-CN，杜绝 Word 默认回退到 ja-JP 进行错误语法校对
    for old_tfl in settings_el.findall(qn("w:themeFontLang")):
        settings_el.remove(old_tfl)
    settings_el.append(parse_xml(f'<w:themeFontLang {nsdecls("w")} w:val="en-US" w:eastAsia="zh-CN"/>'))

    # 关闭文档拼写与语法红色波浪线报错标记（避免专有名词乱飘红）
    for tag in ("w:hideSpellingErrors", "w:hideGrammaticalErrors"):
        for old_el in settings_el.findall(qn(tag)):
            settings_el.remove(old_el)
        settings_el.append(parse_xml(f'<{tag} {nsdecls("w")}/>'))

    # 安全的 updateFields 配置：默认 false，避免 Word 拦截弹窗“可能引用外部文件”
    for old_uf in settings_el.findall(qn("w:updateFields")):
        settings_el.remove(old_uf)
    val_str = "true" if auto_update_fields else "false"
    settings_el.append(parse_xml(f'<w:updateFields {nsdecls("w")} w:val="{val_str}"/>'))

    # Front-matter section (Section 0): empty header (no bottom border) and empty centered footer (no page number)
    for sec in doc.sections:
        sec.different_first_page_header_footer = False
        pgMar = sec._sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:header"), "851")
            pgMar.set(qn("w:footer"), "992")
        hdr = sec.header
        hdr.is_linked_to_previous = False
        p_hd = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
        p_hd_pPr = p_hd._p.get_or_add_pPr()
        for old_bdr in p_hd_pPr.findall(qn("w:pBdr")):
            p_hd_pPr.remove(old_bdr)
        p_hd_pPr.append(parse_xml(
            f'<w:pBdr {nsdecls("w")}><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr>'
        ))
        ftr = sec.footer
        ftr.is_linked_to_previous = False
        for child in list(ftr._element):
            ftr._element.remove(child)
        ftr._element.append(parse_xml(
            f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'
        ))


def _start_main_body_section_with_page_one(doc) -> None:
    """
    Inserts a Next-Page Section Break before the Main Body, unlinks the footer from Front-Matter,
    sets <w:pgNumType w:start="1"/> so Main Body starts at Page 1, and configures the reference
    document's exact <w:sdt> 'Page Numbers (Bottom of Page)' centered 9pt (sz=18) PAGE field footer.
    """
    from docx.enum.section import WD_SECTION
    body_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    body_sec.page_width = Cm(21.0)
    body_sec.page_height = Cm(29.7)
    body_sec.top_margin = Cm(2.5)
    body_sec.bottom_margin = Cm(2.5)
    body_sec.left_margin = Cm(2.5)
    body_sec.right_margin = Cm(2.5)
    body_sec.different_first_page_header_footer = False

    sectPr = body_sec._sectPr
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is not None:
        pgMar.set(qn("w:header"), "851")
        pgMar.set(qn("w:footer"), "992")

    for old_pg in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(old_pg)
    sectPr.append(parse_xml(f'<w:pgNumType {nsdecls("w")} w:start="1"/>'))

    hdr = body_sec.header
    hdr.is_linked_to_previous = False
    p_hd = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    p_hd_pPr = p_hd._p.get_or_add_pPr()
    for old_bdr in p_hd_pPr.findall(qn("w:pBdr")):
        p_hd_pPr.remove(old_bdr)
    p_hd_pPr.append(parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr>'
    ))

    ftr = body_sec.footer
    ftr.is_linked_to_previous = False
    for child in list(ftr._element):
        ftr._element.remove(child)
    sdt_footer = parse_xml(
        f'<w:sdt {nsdecls("w")}>'
        f'<w:sdtPr>'
        f'<w:docPartObj>'
        f'<w:docPartGallery w:val="Page Numbers (Bottom of Page)"/>'
        f'<w:docPartUnique/>'
        f'</w:docPartObj>'
        f'</w:sdtPr>'
        f'<w:sdtContent>'
        f'<w:p>'
        f'<w:pPr>'
        f'<w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs>'
        f'<w:snapToGrid w:val="0"/>'
        f'<w:jc w:val="center"/>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText xml:space="preserve">PAGE   \\* MERGEFORMAT</w:instrText></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="zh-CN"/></w:rPr><w:t>1</w:t></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        f'</w:p>'
        f'</w:sdtContent>'
        f'</w:sdt>'
    )
    ftr._element.append(sdt_footer)


def _collect_document_outline_headings(data: Dict) -> List[Dict]:
    headings: List[Dict] = []
    bm_counter = 1001
    for block in data.get("body_blocks", []):
        if block.get("type") == "heading":
            lvl = int(block.get("level", 1))
            text = block.get("text", "")
            stripped_t = re.sub(r"[\s　]+", "", text)
            if block.get("outline_level") is not None:
                outline_lvl = int(block["outline_level"])
            elif stripped_t in ("参考文献", "附录"):
                outline_lvl = 1
            elif re.match(r"^附录[A-Za-z](?:$|[^.0-9])", stripped_t) and lvl == 1:
                outline_lvl = 2
            else:
                outline_lvl = max(1, min(9, lvl))
            headings.append({"outline_level": outline_lvl, "text": text, "bm_id": bm_counter, "bm_name": f"_Toc{bm_counter}"})
            bm_counter += 1
    if data.get("references"):
        headings.append({"outline_level": 1, "text": "参考文献", "bm_id": bm_counter, "bm_name": f"_Toc{bm_counter}"})
        bm_counter += 1
    appendices = data.get("appendices", [])
    if appendices:
        app_sec_title = data.get("appendix_section_title", "附　　　录")
        headings.append({"outline_level": 1, "text": app_sec_title, "bm_id": bm_counter, "bm_name": f"_Toc{bm_counter}"})
        bm_counter += 1
        for app_idx, app in enumerate(appendices):
            app_code = app.get("code", chr(ord("A") + app_idx))
            app_title = app.get("title", "")
            app_outline_lvl = int(app.get("outline_level", 2))
            heading_label = f"附录 {app_code} {app_title}".rstrip() if app_code else app_title
            headings.append({"outline_level": app_outline_lvl, "text": heading_label, "bm_id": bm_counter, "bm_name": f"_Toc{bm_counter}"})
            bm_counter += 1
            for block in app.get("blocks", []):
                if block.get("type") == "heading":
                    sub_txt = block.get("text", "")
                    sub_lvl = int(block.get("level", 1))
                    if block.get("outline_level") is not None:
                        sub_outline_lvl = int(block["outline_level"])
                    elif re.match(r"^[A-Za-z]\.\d+\.\d+", sub_txt.strip()):
                        sub_outline_lvl = 4
                    elif re.match(r"^[A-Za-z]\.\d+", sub_txt.strip()):
                        sub_outline_lvl = 3
                    else:
                        sub_outline_lvl = min(9, sub_lvl + 2 if sub_lvl <= 2 else sub_lvl)
                    headings.append({"outline_level": sub_outline_lvl, "text": sub_txt, "bm_id": bm_counter, "bm_name": f"_Toc{bm_counter}"})
                    bm_counter += 1
    return headings


def _refresh_toc_and_apply_rule26(docx_path: str, expected_toc_headings: List[Dict]) -> bool:
    abs_path = os.path.abspath(docx_path)
    com_updated = False
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            wdoc = app.Documents.Open(abs_path, False, False)
            wdoc.Repaginate()
            for toc in wdoc.TablesOfContents:
                toc.Update()
            wdoc.Fields.Update()
            wdoc.Save()
            wdoc.Close(False)
            com_updated = True
        finally:
            app.Quit()
    except Exception as exc:
        print(f"[TOC Engine] Word COM auto-update skipped ({exc}); using OpenXML TOC field + updateFields.")

    doc = docx.Document(abs_path)
    if com_updated:
        for old_uf in doc.settings.element.findall(qn("w:updateFields")):
            doc.settings.element.remove(old_uf)

    level_by_norm_text = {re.sub(r"[\s　]+", "", h["text"]): h["outline_level"] for h in expected_toc_headings}
    for p in doc.paragraphs:
        instr_concat = " ".join((n.text or "") for n in p._p.findall(f".//{{{W_NS}}}instrText"))
        if "HYPERLINK" not in instr_concat and "PAGEREF" not in instr_concat:
            continue
        style_name = (p.style.name or "").lower() if p.style else ""
        m_st = re.search(r"toc\s*([1-9])", style_name)
        lvl = int(m_st.group(1)) if m_st else None
        full_txt = "".join((t.text or "") for t in p._p.findall(f".//{{{W_NS}}}t"))
        entry_txt = re.sub(r"\s*\d+\s*$", "", full_txt).strip()
        norm_entry = re.sub(r"[\s　]+", "", entry_txt)
        if norm_entry in level_by_norm_text:
            lvl = level_by_norm_text[norm_entry]
        elif lvl is None:
            lvl = 3 if re.match(r"^(\d+\.\d+\.\d+|[A-Za-z]\.\d+)", entry_txt) else (2 if re.match(r"^(\d+\.\d+|附录\s*[A-Za-z])", entry_txt) else 1)

        cn_font, sz_pt, left_pt = ("黑体", PT_SI_HAO, 0.0) if lvl == 1 else (("宋体", PT_XIAO_SI, 24.0) if lvl == 2 else ("宋体", PT_XIAO_SI, 48.0))
        format_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, left_indent_pt=left_pt, outline_level=None)
        _set_right_dot_tab(p, 9072)
        sz_half = int(sz_pt * 2)
        for r_el in list(p._p.findall(f".//{{{W_NS}}}r")):
            rPr = r_el.find(qn("w:rPr"))
            if rPr is None:
                rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
                r_el.insert(0, rPr)
            for tag in ("w:rStyle", "w:rFonts", "w:sz", "w:szCs", "w:b", "w:bCs", "w:u", "w:color"):
                for old in rPr.findall(qn(tag)):
                    rPr.remove(old)
            rPr.append(parse_xml(
                f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="{cn_font}"/>'
            ))
            rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_half}"/>'))
            rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_half}"/>'))
            rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="000000"/>'))

    toc_end_idx, first_body_idx = None, None
    for i, p in enumerate(doc.paragraphs):
        instr_concat = " ".join((n.text or "") for n in p._p.findall(f".//{{{W_NS}}}instrText"))
        if "HYPERLINK" in instr_concat or "PAGEREF" in instr_concat:
            toc_end_idx = i
        if first_body_idx is None and get_outline_level(p) is not None:
            first_body_idx = i
    if toc_end_idx is not None and first_body_idx is not None and first_body_idx > toc_end_idx:
        if not any(_has_page_break(doc.paragraphs[k]) for k in range(toc_end_idx, first_body_idx)):
            doc.paragraphs[toc_end_idx].add_run().add_break(docx.enum.text.WD_BREAK.PAGE)

    doc.save(abs_path)
    return com_updated


def build_docx_from_schema(data: Dict, output_path: str) -> None:
    doc = docx.Document()

    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(2.5)
        sec.right_margin = Cm(2.5)

    auto_update_fields = bool(data.get("auto_update_fields", False))
    _configure_toc_styles_and_footer(doc, auto_update_fields=auto_update_fields)
    all_outline_headings = _collect_document_outline_headings(data)
    toc_max_level = int(data.get("toc_max_level", 3))
    auto_toc_headings = [h for h in all_outline_headings if h["outline_level"] <= toc_max_level]
    heading_cursor = 0

    def _next_heading_bookmark() -> Tuple[int, str]:
        nonlocal heading_cursor
        if heading_cursor < len(all_outline_headings):
            h_info = all_outline_headings[heading_cursor]
            heading_cursor += 1
            return h_info["bm_id"], h_info["bm_name"]
        bm_id = 2000 + heading_cursor
        heading_cursor += 1
        return bm_id, f"_Toc{bm_id}"

    # Rule 2.1: Cover Block
    cover = data.get("cover", {})
    if cover:
        p_cn_t = doc.add_paragraph()
        format_paragraph(p_cn_t, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=18, space_after_pt=8)
        add_mixed_runs(p_cn_t, cover.get("cn_title", ""), cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_ER, bold=True)

        p_en_t = doc.add_paragraph()
        format_paragraph(p_en_t, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=4, space_after_pt=16)
        add_mixed_runs(p_en_t, cover.get("en_title", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_SAN_HAO, bold=True)

        for field_label, field_val in cover.get("fields", []):
            p_f = doc.add_paragraph()
            format_paragraph(p_f, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=3, space_after_pt=3)
            add_mixed_runs(p_f, f"{field_label}：{field_val}", cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SAN, bold=False)

        doc.add_page_break()

    # Rules 2.2 & 2.3: Chinese Title, Author, Affiliation, Abstract & Keywords (Page Break After!)
    cn_abs = data.get("cn_abstract_block", {})
    if cn_abs:
        _ensure_page_break_before(doc)
        p_t = doc.add_paragraph()
        format_paragraph(p_t, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=6, space_after_pt=4)
        add_mixed_runs(p_t, cn_abs.get("title", ""), cn_font="黑体", en_font="Times New Roman", size_pt=PT_SAN_HAO, bold=False)

        p_a = doc.add_paragraph()
        format_paragraph(p_a, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=2, space_after_pt=2)
        add_mixed_runs(p_a, cn_abs.get("authors", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_SI_HAO)

        p_aff = doc.add_paragraph()
        format_paragraph(p_aff, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=2, space_after_pt=10)
        add_mixed_runs(p_aff, cn_abs.get("affiliation", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)

        p_ab = doc.add_paragraph()
        format_paragraph(p_ab, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, first_line_indent_pt=24.0, space_before_pt=4, space_after_pt=4)
        add_mixed_runs(p_ab, "摘要：", cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)
        add_mixed_runs(p_ab, cn_abs.get("abstract", ""), cn_font="楷体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)

        p_kw = doc.add_paragraph()
        format_paragraph(p_kw, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, first_line_indent_pt=0.0, space_before_pt=2, space_after_pt=12)
        add_mixed_runs(p_kw, "关键词：", cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)
        kw_str = "；".join(cn_abs.get("keywords", []))
        add_mixed_runs(p_kw, kw_str, cn_font="楷体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)

        doc.add_page_break()

    # Rules 2.4 & 2.5: English Title, Author, Affiliation, Abstract & Key Words (Page Break Before & After!)
    en_abs = data.get("en_abstract_block", {})
    if en_abs:
        _ensure_page_break_before(doc)
        p_et = doc.add_paragraph()
        format_paragraph(p_et, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=10, space_after_pt=4)
        add_mixed_runs(p_et, en_abs.get("title", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_SAN_HAO, bold=True)

        p_ea = doc.add_paragraph()
        format_paragraph(p_ea, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=2, space_after_pt=2)
        add_mixed_runs(p_ea, en_abs.get("authors", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_SI_HAO)

        p_eaff = doc.add_paragraph()
        format_paragraph(p_eaff, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=2, space_after_pt=10)
        add_mixed_runs(p_eaff, en_abs.get("affiliation", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)

        p_eab = doc.add_paragraph()
        format_paragraph(p_eab, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, first_line_indent_pt=24.0, space_before_pt=4, space_after_pt=4)
        add_mixed_runs(p_eab, "Abstract: ", cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=True)
        add_mixed_runs(p_eab, en_abs.get("abstract", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)

        p_ekw = doc.add_paragraph()
        format_paragraph(p_ekw, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, first_line_indent_pt=0.0, space_before_pt=2, space_after_pt=12)
        add_mixed_runs(p_ekw, "Key Words: ", cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=True)
        ekw_str = "; ".join(en_abs.get("keywords", []))
        add_mixed_runs(p_ekw, ekw_str, cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=False)

        doc.add_page_break()

    # Rule 2.6: Native Word Automatic Table of Contents (TOC \o "1-3" \h \z \u)
    has_toc_section = bool(data.get("toc", True)) and len(auto_toc_headings) > 0
    if has_toc_section:
        _ensure_page_break_before(doc)
        p_toc_title = doc.add_paragraph()
        format_paragraph(p_toc_title, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5, space_before_pt=0, space_after_pt=0)
        add_mixed_runs(p_toc_title, "目　　　录", cn_font="黑体", en_font="Times New Roman", size_pt=PT_SAN_HAO, bold=False)

        p_blank = doc.add_paragraph()
        format_paragraph(p_blank, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0, space_before_pt=0, space_after_pt=0)
        add_mixed_runs(p_blank, " ", cn_font="宋体", en_font="Times New Roman", size_pt=PT_WU_HAO)

        for idx, h_item in enumerate(auto_toc_headings):
            lvl = h_item["outline_level"]
            text = h_item["text"]
            bm_name = h_item["bm_name"]
            p_item = doc.add_paragraph()
            cn_f, sz_p, indent_p = ("黑体", PT_SI_HAO, 0.0) if lvl == 1 else (("宋体", PT_XIAO_SI, 24.0) if lvl == 2 else ("宋体", PT_XIAO_SI, 48.0))
            format_paragraph(p_item, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5, left_indent_pt=indent_p)
            _set_right_dot_tab(p_item, 9072)

            if idx == 0:
                _add_fld_char(p_item, "begin", dirty=auto_update_fields)
                _add_instr_text(p_item, f' TOC \\o "1-{toc_max_level}" \\h \\z \\u ')
                _add_fld_char(p_item, "separate")

            _add_fld_char(p_item, "begin")
            _add_instr_text(p_item, f" HYPERLINK \\l {bm_name} ")
            _add_fld_char(p_item, "separate")
            add_mixed_runs(p_item, text, cn_font=cn_f, en_font="Times New Roman", size_pt=sz_p)
            _add_tab_run(p_item, cn_font=cn_f, en_font="Times New Roman", size_pt=sz_p)

            _add_fld_char(p_item, "begin")
            _add_instr_text(p_item, f" PAGEREF {bm_name} \\h ")
            _add_fld_char(p_item, "separate")
            r_pg = p_item.add_run("1")
            set_run_fonts(r_pg, cn_font=cn_f, en_font="Times New Roman", size_pt=sz_p)
            _add_fld_char(p_item, "end")
            _add_fld_char(p_item, "end")

            if idx == len(auto_toc_headings) - 1:
                _add_fld_char(p_item, "end")

    # Rule 2.17: Start Main Body in a new section with Page Number starting at 1
    if len(doc.paragraphs) > 0:
        _start_main_body_section_with_page_one(doc)

    # Rule 2.7 & Rule 2.16: Main Body Sections (with Word Outline Levels 1..9)
    for block in data.get("body_blocks", []):
        btype = block.get("type", "paragraph")
        if btype == "heading":
            lvl = int(block.get("level", 1))
            text = block.get("text", "")
            stripped_t = re.sub(r"[\s　]+", "", text)
            if "outline_level" in block and block["outline_level"] is not None:
                outline_lvl = int(block["outline_level"])
            elif stripped_t in ("参考文献", "附录"):
                outline_lvl = 1
            elif re.match(r"^附录[A-Za-z](?:$|[^.0-9])", stripped_t) and lvl == 1:
                outline_lvl = 2
            else:
                outline_lvl = max(1, min(9, lvl))

            p_h = doc.add_paragraph()
            if lvl == 1:
                format_paragraph(
                    p_h, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
                    space_before_pt=8, space_after_pt=4, keep_with_next=True,
                    outline_level=outline_lvl,
                )
                add_mixed_runs(p_h, text, cn_font="黑体", en_font="Times New Roman", size_pt=PT_SI_HAO)
            elif lvl == 2:
                format_paragraph(
                    p_h, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
                    space_before_pt=6, space_after_pt=2, keep_with_next=True,
                    outline_level=outline_lvl,
                )
                add_mixed_runs(p_h, text, cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_SI)
            else:
                format_paragraph(
                    p_h, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
                    space_before_pt=4, space_after_pt=2, keep_with_next=True,
                    outline_level=outline_lvl,
                )
                add_mixed_runs(p_h, text, cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI, bold=True)
            bm_id, bm_name = _next_heading_bookmark()
            _attach_bookmark(p_h, bm_id, bm_name)
        elif btype == "paragraph":
            p_b = doc.add_paragraph()
            p_outline = int(block["outline_level"]) if block.get("outline_level") is not None else None
            format_paragraph(
                p_b, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5,
                first_line_indent_pt=24.0, outline_level=p_outline,
            )
            add_mixed_runs(p_b, block.get("text", ""), cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI)
        elif btype == "equation":
            add_numbered_equation(doc, block.get("formula", ""), block.get("number", ""))
        elif btype == "table":
            add_three_line_table(
                doc,
                table_num=block.get("table_num", ""),
                table_title=block.get("table_title", ""),
                headers=block.get("headers", []),
                rows_data=block.get("rows", []),
                group_headers=block.get("group_headers"),
                footnote=block.get("footnote"),
            )
        elif btype == "figure":
            add_figure_with_caption(
                doc,
                image_path=block.get("image_path", ""),
                fig_num=block.get("fig_num", ""),
                fig_title=block.get("fig_title", ""),
                subfig_notes=block.get("subfig_notes"),
                width_inches=block.get("width_inches", 5.2),
            )

    # Rules 2.8, 2.14 & 2.16: References (If present, "参考文献" heading has Outline Level 1 and starts on a new page)
    refs = data.get("references", [])
    if refs:
        _ensure_page_break_before(doc)
        p_ref_h = doc.add_paragraph()
        format_paragraph(
            p_ref_h, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
            space_before_pt=12, space_after_pt=6, keep_with_next=True,
            outline_level=1,
        )
        add_mixed_runs(p_ref_h, "参考文献", cn_font="黑体", en_font="Times New Roman", size_pt=PT_SI_HAO)
        bm_id, bm_name = _next_heading_bookmark()
        _attach_bookmark(p_ref_h, bm_id, bm_name)

        for ref_str in refs:
            p_r = doc.add_paragraph()
            format_paragraph(p_r, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, first_line_indent_pt=0.0)
            add_mixed_runs(p_r, ref_str, cn_font="宋体", en_font="Times New Roman", size_pt=PT_XIAO_SI)

    # Rules 2.9, 2.15 & 2.16: Appendices (Starts on a new page)
    # If present: top-level "附　　　录" / "附录" heading is Outline Level 1 (val="0"),
    # each appendix item ("附录 A ...", "附录 B ...", "附录 C ...") is Outline Level 2 (val="1"),
    # and sub-headings inside an appendix shift to Outline Level 3, 4, ...
    appendices = data.get("appendices", [])
    if appendices:
        _ensure_page_break_before(doc)
        app_sec_title = data.get("appendix_section_title", "附　　　录")
        p_app_sec = doc.add_paragraph()
        format_paragraph(
            p_app_sec, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
            space_before_pt=8, space_after_pt=4, keep_with_next=True,
            outline_level=1,
        )
        add_mixed_runs(p_app_sec, app_sec_title, cn_font="黑体", en_font="Times New Roman", size_pt=PT_SI_HAO)
        bm_id, bm_name = _next_heading_bookmark()
        _attach_bookmark(p_app_sec, bm_id, bm_name)

        for app_idx, app in enumerate(appendices):
            if app_idx > 0 and app.get("page_break", False):
                doc.add_page_break()
            app_code = app.get("code", chr(ord("A") + app_idx))
            app_title = app.get("title", "")
            app_outline_lvl = int(app.get("outline_level", 2))

            p_app_top = doc.add_paragraph()
            format_paragraph(
                p_app_top, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
                space_before_pt=6, space_after_pt=4, keep_with_next=True,
                outline_level=app_outline_lvl,
            )
            heading_label = f"附录 {app_code} {app_title}".rstrip() if app_code else app_title
            add_mixed_runs(p_app_top, heading_label, cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_SI)
            bm_id, bm_name = _next_heading_bookmark()
            _attach_bookmark(p_app_top, bm_id, bm_name)

            for block in app.get("blocks", []):
                btype = block.get("type", "paragraph")
                if btype == "heading":
                    sub_txt = block.get("text", "")
                    sub_lvl = int(block.get("level", 1))
                    if "outline_level" in block and block["outline_level"] is not None:
                        sub_outline_lvl = int(block["outline_level"])
                    elif re.match(r"^[A-Za-z]\.\d+\.\d+", sub_txt.strip()):
                        sub_outline_lvl = 4
                    elif re.match(r"^[A-Za-z]\.\d+", sub_txt.strip()):
                        sub_outline_lvl = 3
                    else:
                        sub_outline_lvl = min(9, sub_lvl + 2 if sub_lvl <= 2 else sub_lvl)

                    p_ah = doc.add_paragraph()
                    format_paragraph(
                        p_ah, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5,
                        space_before_pt=4, space_after_pt=2, keep_with_next=True,
                        outline_level=sub_outline_lvl,
                    )
                    add_mixed_runs(p_ah, sub_txt, cn_font="黑体", en_font="Times New Roman", size_pt=PT_XIAO_SI)
                    bm_id, bm_name = _next_heading_bookmark()
                    _attach_bookmark(p_ah, bm_id, bm_name)
                elif btype == "paragraph":
                    p_ap = doc.add_paragraph()
                    p_ap_outline = int(block["outline_level"]) if block.get("outline_level") is not None else None
                    format_paragraph(
                        p_ap, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5,
                        first_line_indent_pt=24.0, outline_level=p_ap_outline,
                    )
                    add_mixed_runs(p_ap, block.get("text", ""), cn_font="楷体", en_font="Times New Roman", size_pt=PT_XIAO_SI)
                elif btype == "equation":
                    add_numbered_equation(doc, block.get("formula", ""), block.get("number", ""))
                elif btype == "table":
                    add_three_line_table(
                        doc,
                        table_num=block.get("table_num", ""),
                        table_title=block.get("table_title", ""),
                        headers=block.get("headers", []),
                        rows_data=block.get("rows", []),
                        group_headers=block.get("group_headers"),
                        footnote=block.get("footnote"),
                    )
                elif btype == "figure":
                    add_figure_with_caption(
                        doc,
                        image_path=block.get("image_path", ""),
                        fig_num=block.get("fig_num", ""),
                        fig_title=block.get("fig_title", ""),
                        subfig_notes=block.get("subfig_notes"),
                        width_inches=block.get("width_inches", 5.2),
                    )
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    if has_toc_section:
        _refresh_toc_and_apply_rule26(output_path, auto_toc_headings)

    # 如果未显式开启 auto_update_fields，确保最终文档将 w:updateFields 设为 false 并移除所有 w:dirty，避免打开安全弹窗
    if not auto_update_fields:
        import zipfile, shutil
        tmp_clean = output_path + ".tmp"
        with zipfile.ZipFile(output_path, "r") as zin, zipfile.ZipFile(tmp_clean, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_b = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data_b = data_b.replace(b' w:dirty="true"', b"")
                elif item.filename == "word/settings.xml" and b'<w:updateFields w:val="true"/>' in data_b:
                    data_b = data_b.replace(b'<w:updateFields w:val="true"/>', b'<w:updateFields w:val="false"/>')
                zout.writestr(item, data_b)
        shutil.move(tmp_clean, output_path)

    print(f"[DOCX Engine v4] Saved strictly formatted document (with native OMML Math, Auto-TOC, Outline Levels & Fixed-Grid Three-Line Tables) to: {output_path}")


def verify_docx_xml_audit(docx_path: str) -> Dict:
    """
    Mode B Programmatic OpenXML Auditor for Pure-Text Models:
    Inspects the generated .docx XML tree without requiring visual image inspection.
    Verifies native <m:oMath> counts, checks for any remaining raw underscore math,
    validates three-line table border/width compliance, audits Word Outline Levels (<w:outlineLvl>),
    front-matter page breaks (CN Abstract -> EN Abstract -> TOC -> Body), and native Word Automatic TOC.
    """
    doc = docx.Document(docx_path)
    root = doc._element
    omath_nodes = root.findall(f".//{{{MATH_NS}}}oMath")
    f_nodes = root.findall(f".//{{{MATH_NS}}}f")
    nary_nodes = root.findall(f".//{{{MATH_NS}}}nary")
    sub_nodes = root.findall(f".//{{{MATH_NS}}}sSub") + root.findall(f".//{{{MATH_NS}}}sSubSup")

    raw_underscore_hits = []
    outline_tree = []
    outline_level_counts: Dict[str, int] = {}

    cn_kw_idx, en_title_idx, en_kw_idx, toc_title_idx, first_h1_idx = None, None, None, None, None
    ref_title_idx, last_body_idx, app_title_idx, last_ref_idx = None, None, None, None
    has_toc_field = False
    toc_hyperlink_count = 0
    toc_pageref_count = 0
    toc_entries_with_pages = []

    for idx, p in enumerate(doc.paragraphs):
        txt = "".join((t.text or "") for t in p._p.findall(f".//{{{W_NS}}}t"))
        norm_txt = re.sub(r"[\s　]+", "", txt)
        if "_" in txt:
            raw_underscore_hits.append(txt[:60])

        if txt.startswith("关键词："):
            cn_kw_idx = idx
        elif txt.startswith("Key Words:"):
            en_kw_idx = idx
        elif norm_txt == "目录":
            toc_title_idx = idx
        elif en_title_idx is None and cn_kw_idx is not None and toc_title_idx is None and txt.strip():
            en_title_idx = idx
        elif norm_txt == "参考文献":
            ref_title_idx = idx
        elif norm_txt in ("附录", "附　　录") or re.match(r"^附录[A-Za-z]", norm_txt):
            if app_title_idx is None:
                app_title_idx = idx

        # 记录正文最后一项段落与参考文献最后一项段落位置
        if first_h1_idx is not None and ref_title_idx is None and app_title_idx is None and idx > first_h1_idx:
            if txt.strip() and not _has_page_break(p):
                last_body_idx = idx
        elif ref_title_idx is not None and app_title_idx is None and idx > ref_title_idx:
            if txt.strip() and not _has_page_break(p):
                last_ref_idx = idx

        instr_joined = " ".join((n.text or "") for n in p._p.findall(f".//{{{W_NS}}}instrText"))
        if "TOC " in instr_joined and "\\u" in instr_joined:
            has_toc_field = True
        if "HYPERLINK" in instr_joined and "_Toc" in instr_joined:
            toc_hyperlink_count += 1
        if "PAGEREF" in instr_joined and "_Toc" in instr_joined:
            toc_pageref_count += 1
            m_pg = re.match(r"^(.*?)\s*(\d+)\s*$", txt.strip())
            toc_entries_with_pages.append({
                "entry": m_pg.group(1).strip() if m_pg else txt.strip(),
                "page": int(m_pg.group(2)) if m_pg else None,
            })

        ol = get_outline_level(p)
        if ol is not None:
            if first_h1_idx is None:
                first_h1_idx = idx
            outline_tree.append({"outline_level": ol, "xml_val": ol - 1, "text": txt})
            key = f"level_{ol}"
            outline_level_counts[key] = outline_level_counts.get(key, 0) + 1

    def _has_break_between(start_i: Optional[int], end_i: Optional[int]) -> bool:
        if start_i is None or end_i is None or end_i <= start_i:
            return True
        return any(_has_page_break(doc.paragraphs[i]) for i in range(start_i, end_i))

    pb_cn_to_next = _has_break_between(cn_kw_idx, en_title_idx if en_title_idx is not None else toc_title_idx)
    pb_before_toc = _has_break_between(en_kw_idx if en_kw_idx is not None else cn_kw_idx, toc_title_idx) if toc_title_idx is not None else True
    pb_after_toc = _has_break_between(toc_title_idx, first_h1_idx) if toc_title_idx is not None else True
    pb_body_to_refs = _has_break_between(last_body_idx, ref_title_idx) if (ref_title_idx is not None and last_body_idx is not None) else True
    pb_to_appendices = _has_break_between(last_ref_idx if last_ref_idx is not None else last_body_idx, app_title_idx) if app_title_idx is not None else True

    document_page_breaks = {
        "cn_abstract_to_en_abstract_or_next": pb_cn_to_next,
        "preceding_section_to_toc": pb_before_toc,
        "toc_to_main_body": pb_after_toc,
        "main_body_to_references": pb_body_to_refs,
        "references_or_body_to_appendices": pb_to_appendices,
        "all_page_breaks_valid": (pb_cn_to_next and pb_before_toc and pb_after_toc and pb_body_to_refs and pb_to_appendices),
    }

    native_auto_toc = {
        "has_toc_field": has_toc_field,
        "toc_hyperlink_count": toc_hyperlink_count,
        "toc_pageref_count": toc_pageref_count,
        "toc_entries_with_pages": toc_entries_with_pages,
        "auto_toc_valid": (
            (toc_title_idx is None) or (has_toc_field and toc_hyperlink_count > 0 and toc_pageref_count == toc_hyperlink_count)
        ),
    }

    # Check Rule 2.17 Page Numbering Specification (Front-Matter no page number; Main Body starts at Page 1, centered 9pt PAGE field)
    front_sec_has_page_field = False
    body_sec_starts_at_1 = False
    body_footer_has_sdt_page = False
    body_footer_centered_9pt = False
    if len(doc.sections) >= 2:
        front_ftr_xml = doc.sections[0].footer._element.xml
        front_sec_has_page_field = ("PAGE" in front_ftr_xml)
        body_sectPr = doc.sections[1]._sectPr
        pg_num_el = body_sectPr.find(qn("w:pgNumType"))
        body_sec_starts_at_1 = (pg_num_el is not None and pg_num_el.get(qn("w:start")) == "1")
        body_ftr_el = doc.sections[1].footer._element
        body_ftr_xml = body_ftr_el.xml
        body_footer_has_sdt_page = (
            ("Page Numbers (Bottom of Page)" in body_ftr_xml or "w:sdt" in body_ftr_xml)
            and "PAGE" in body_ftr_xml
            and "MERGEFORMAT" in body_ftr_xml
        )
        body_footer_centered_9pt = ('w:jc w:val="center"' in body_ftr_xml and 'w:sz w:val="18"' in body_ftr_xml)

    first_toc_page_is_1 = (len(toc_entries_with_pages) == 0 or toc_entries_with_pages[0].get("page") == 1)
    page_numbering_spec = {
        "sections_count": len(doc.sections),
        "front_matter_has_no_page_number": (not front_sec_has_page_field),
        "main_body_starts_at_page_1": body_sec_starts_at_1,
        "first_body_heading_toc_page_is_1": first_toc_page_is_1,
        "body_footer_has_sdt_page_field": body_footer_has_sdt_page,
        "body_footer_centered_9pt_xiaowu": body_footer_centered_9pt,
        "page_numbering_valid": (
            len(doc.sections) >= 2
            and (not front_sec_has_page_field)
            and body_sec_starts_at_1
            and first_toc_page_is_1
            and body_footer_has_sdt_page
            and body_footer_centered_9pt
        ),
    }

    # Check Rule 2.16 Outline Level compliance for References & Appendices (when present)
    outline_rules_ok = len(outline_tree) > 0
    for item in outline_tree:
        norm_t = re.sub(r"[\s　]+", "", item["text"])
        if norm_t in ("参考文献", "附录"):
            if item["outline_level"] != 1:
                outline_rules_ok = False
        elif re.match(r"^附录[A-Za-z](?:$|[^.0-9])", norm_t):
            if item["outline_level"] != 2:
                outline_rules_ok = False

    content_tables_checked = 0
    three_line_compliant = 0
    for tbl in doc.tables:
        tblPr = tbl._tbl.tblPr
        borders = tblPr.find(qn("w:tblBorders"))
        if borders is not None:
            top_el = borders.find(qn("w:top"))
            bot_el = borders.find(qn("w:bottom"))
            left_el = borders.find(qn("w:left"))
            if top_el is not None and top_el.get(qn("w:val")) == "single":
                content_tables_checked += 1
                if (
                    top_el.get(qn("w:sz")) == "12"
                    and bot_el is not None and bot_el.get(qn("w:sz")) == "12"
                    and (left_el is None or left_el.get(qn("w:val")) in ("nil", "none"))
                ):
                    three_line_compliant += 1

    # Check for empty OMML slots (<m:num>, <m:den>, <m:e>) that render as ⬚ dashed placeholders in Word
    empty_math_slots = []
    for tag in ("e", "num", "den"):
        for el in root.findall(f".//{{{MATH_NS}}}{tag}"):
            if len(list(el)) == 0 and not (el.text and el.text.strip()):
                empty_math_slots.append(tag)

    settings_xml = ""
    try:
        settings_xml = doc.settings.element.xml
    except Exception:
        pass
    has_ja_jp = 'w:eastAsia="ja-JP"' in settings_xml
    has_zh_cn = 'w:eastAsia="zh-CN"' in settings_xml
    spelling_hidden = "w:hideSpellingErrors" in settings_xml

    report = {
        "docx_path": docx_path,
        "total_paragraphs": len(doc.paragraphs),
        "native_omath_equations_count": len(omath_nodes),
        "omath_fractions_count": len(f_nodes),
        "omath_summations_count": len(nary_nodes),
        "omath_subscripts_count": len(sub_nodes),
        "raw_underscore_paragraphs_count": len(raw_underscore_hits),
        "empty_math_slots_count": len(empty_math_slots),
        "empty_math_slots_details": empty_math_slots,
        "language_theme_has_ja_jp": has_ja_jp,
        "language_theme_has_zh_cn": has_zh_cn,
        "hide_spelling_errors_configured": spelling_hidden,
        "content_tables_count": content_tables_checked,
        "three_line_compliant_tables_count": three_line_compliant,
        "document_page_breaks": document_page_breaks,
        "front_matter_page_breaks": document_page_breaks,
        "page_numbering_spec": page_numbering_spec,
        "native_auto_toc": native_auto_toc,
        "outline_paragraphs_count": len(outline_tree),
        "outline_level_counts": outline_level_counts,
        "outline_tree": outline_tree,
        "outline_compliance": outline_rules_ok,
        "audit_passed": (
            len(omath_nodes) > 0
            and len(raw_underscore_hits) == 0
            and len(empty_math_slots) == 0
            and not has_ja_jp
            and content_tables_checked == three_line_compliant
            and outline_rules_ok
            and document_page_breaks["all_page_breaks_valid"]
            and page_numbering_spec["page_numbering_valid"]
            and native_auto_toc["auto_toc_valid"]
        ),
    }
    print("[OpenXML Programmatic Audit Report]:", json.dumps(report, indent=2, ensure_ascii=False))
    assert report["audit_passed"], f"Audit failed! Details: {report}"
    return report


if __name__ == "__main__":
    import subprocess
    parser = argparse.ArgumentParser(description="Build Chinese Academic Thesis/Translation DOCX (Rules 2.1-2.16)")
    parser.add_argument("--input", type=str, required=True, help="Path to translated_paper_data.json")
    parser.add_argument("--output", type=str, required=True, help="Output .docx file path")
    parser.add_argument("--render-qa-dir", type=str, default="", help="Mode A (Vision): Directory to render 150-DPI Visual QA page PNGs")
    parser.add_argument("--verify-xml-audit", action="store_true", help="Mode B (Pure-Text): Run programmatic OpenXML audit on generated .docx")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        schema_data = json.load(f)
    build_docx_from_schema(schema_data, args.output)

    if args.render_qa_dir:
        os.makedirs(args.render_qa_dir, exist_ok=True)
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", args.render_qa_dir, os.path.abspath(args.output)],
            check=True
        )
        pdf_name = os.path.splitext(os.path.basename(args.output))[0] + ".pdf"
        pdf_path = os.path.join(args.render_qa_dir, pdf_name)
        subprocess.run(
            ["pdftoppm", "-png", "-r", "150", pdf_path, os.path.join(args.render_qa_dir, "qa_page")],
            check=True
        )
        print(f"[Visual QA] Rendered QA page PNGs in: {args.render_qa_dir}")

    if args.verify_xml_audit:
        verify_docx_xml_audit(args.output)
