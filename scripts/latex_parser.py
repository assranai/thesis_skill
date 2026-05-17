#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ⚠️ 新团队接手前必须先读: ../HANDOFF_遗嘱_必读.md
"""
LaTeX Parser for Chinese Engineering Thesis
Parses all .tex files and outputs structured JSON.
"""

import os
import re
import json
import glob

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = r"E:\thesis\skill_test\latex模板"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "thesis_content.json")
FIGURE_DIR = os.path.join(BASE_DIR, "figure")

# ============================================================
# FILE READING
# ============================================================
def read_file(filepath):
    """Read a UTF-8 text file, return content or empty string on error."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        print(f"  [WARN] Cannot read {filepath}: {e}")
        return ""


# ============================================================
# METADATA PARSING
# ============================================================
def parse_metadata(com_info_text, bachelor_info_text):
    """Extract metadata from com_info.tex and bachelor_info.tex."""
    meta = {}

    # Author
    m = re.search(r'\\thesisauthor\s*\{(.+?)\}\{(.+?)\}', com_info_text, re.DOTALL)
    if m:
        meta["author_cn"] = m.group(1).strip()
        meta["author_en"] = m.group(2).strip()

    # Teacher
    m = re.search(r'\\teacher\s*\{(.+?)\}\{(.+?)\}', com_info_text, re.DOTALL)
    if m:
        meta["teacher_cn"] = m.group(1).strip()
        meta["teacher_en"] = m.group(2).strip()

    # School
    m = re.search(r'\\school\s*\{(.+?)\}\{(.+?)\}', com_info_text, re.DOTALL)
    if m:
        meta["school_cn"] = m.group(1).strip()
        meta["school_en"] = m.group(2).strip()

    # Major
    m = re.search(r'\\major\s*\{(.+?)\}\{(.+?)\}', com_info_text, re.DOTALL)
    if m:
        meta["major_cn"] = m.group(1).strip()
        meta["major_en"] = m.group(2).strip()

    # Thesis title (4 brace groups: cn_main, cn_sub, en_main, en_sub)
    # Use [^}]* instead of .+? to prevent crossing brace boundaries, with \s* between groups for newlines
    m = re.search(r'\\thesistitle\s*\{([^}]*)\}\s*\{([^}]*)\}\s*\{([^}]*)\}\s*\{([^}]*)\}', com_info_text)
    if m:
        cn_title = m.group(1).strip()
        cn_sub = m.group(2).strip()
        meta["thesis_title_cn"] = cn_title if cn_title else cn_sub
        en_title = m.group(3).strip()
        en_sub = m.group(4).strip()
        meta["thesis_title_en"] = en_title if en_title else en_sub

    # Student ID
    m = re.search(r'\\studentID\s*\{(.+?)\}', bachelor_info_text)
    if m:
        meta["student_id"] = m.group(1).strip()

    # Class
    m = re.search(r'\\class\s*\{(.+?)\}', bachelor_info_text)
    if m:
        meta["class"] = m.group(1).strip()

    # Category (中图分类号)
    m = re.search(r'\\category\s*\{(.+?)\}', com_info_text)
    if m:
        meta["category"] = m.group(1).strip()

    return meta


# ============================================================
# ABSTRACT PARSING
# ============================================================
def parse_abstract(text):
    """Extract Chinese and English abstracts with keywords."""
    result = {"abstract_cn": "", "abstract_cn_keywords": "", "abstract_en": "", "abstract_en_keywords": ""}

    # Chinese abstract
    m = re.search(r'\\begin\{cabstract\}(.*?)\\end\{cabstract\}', text, re.DOTALL)
    if m:
        result["abstract_cn"] = m.group(1).strip()

    # English abstract
    m = re.search(r'\\begin\{eabstract\}(.*?)\\end\{eabstract\}', text, re.DOTALL)
    if m:
        result["abstract_en"] = m.group(1).strip()

    # Keywords from com_info
    m = re.search(r'\\ckeyword\s*\{(.+?)\}', text)
    if m:
        result["abstract_cn_keywords"] = m.group(1).strip()

    m = re.search(r'\\ekeyword\s*\{(.+?)\}', text)
    if m:
        result["abstract_en_keywords"] = m.group(1).strip()

    return result


# ============================================================
# LATEX TEXT CLEANING
# ============================================================
def clean_latex_text(text):
    r"""
    Clean LaTeX commands from paragraph text.
    - \upcite{keys} -> [REF:key1,key2]
    - \ref{label} -> [REF:{label}]
    - \textbf{text} -> text
    - \textit{text} -> text
    - \textsf{text} -> text
    - $...$ -> keep as-is
    - \bm{X} -> $\bm{X}$
    - \` and \'' -> Chinese quotes
    - Skip formatting commands
    """
    # Replace \upcite{key1,key2,...} with [REF:key1,key2]
    text = re.sub(r'\\upcite\{([^}]+)\}', lambda m: '[REF:' + m.group(1) + ']', text)

    # Replace \ref{label} with [REF:{label}]
    text = re.sub(r'\\ref\{([^}]+)\}', r'[REF:{\1}]', text)

    # Replace \textbf{text} -> text
    text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)

    # Replace \textit{text} -> text
    text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)

    # Replace \textsf{text} -> text
    text = re.sub(r'\\textsf\{([^}]*)\}', r'\1', text)

    # \bm is already inside $...$ math mode — do NOT wrap it with extra $...$
    # Keep only the inner content (consistent with \textbf, \textit above).
    text = re.sub(r'\\bm\{([^}]+)\}', r'\1', text)

    # Replace \` and \'' with Chinese quotes
    LEFT_DQ = chr(0x201C)
    RIGHT_DQ = chr(0x201D)
    text = text.replace("\\`", LEFT_DQ)
    text = text.replace("''", RIGHT_DQ)

    # Replace \text{...} -> content
    text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

    # Replace \songti, \zihao{N}, \resizebox{...}{!}{...} -> keep inner content
    text = re.sub(r'\\songti', '', text)
    text = re.sub(r'\\zihao\{[^}]*\}', '', text)
    text = re.sub(r'\\resizebox\{[^}]*\}\{[^}]*\}\{', '', text)
    # Close extra braces from \resizebox — BUT preserve $...$ math segments
    MATH_PH = '\x01__MATH__\x01'
    _saved = []
    def _save_dollar(m):
        _saved.append(m.group(0))
        return f'{MATH_PH}{len(_saved)-1}\x01'
    text = re.sub(r'\$[^$]*\$', _save_dollar, text)
    text = re.sub(r'\}', '', text)
    for i, seg in enumerate(_saved):
        text = text.replace(f'{MATH_PH}{i}\x01', seg)

    # Remove formatting commands
    for cmd in ['tolerance', 'emergencystretch', 'cleardoublepage', 'phantomsection',
                'addcontentsline', 'mainmatter', 'pagestyle', 'tableofcontents',
                'maketitle', 'graphicspath', 'centering', 'captionsetup']:
        text = re.sub(r'\\' + cmd + r'(\{[^}]*\})?', '', text)

    # Remove \begin{...} and \end{...} for non-content environments
    text = re.sub(r'\\begin\{figure\*?\}.*?\\end\{figure\*?\}', '', text, flags=re.DOTALL)
    text = re.sub(r'\\begin\{table\*?\}.*?\\end\{table\*?\}', '', text, flags=re.DOTALL)
    text = re.sub(r'\\begin\{equation\*?\}.*?\\end\{equation\*?\}', '', text, flags=re.DOTALL)

    # Remove standalone braces that are just grouping
    # But be careful not to remove braces inside math mode

    # Collapse multiple whitespace
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)

    return text.strip()


def clean_inline_text(text):
    """Light cleaning for inline text (section titles, captions, etc.)."""
    text = re.sub(r'\\upcite\{([^}]+)\}', lambda m: '[REF:' + m.group(1) + ']', text)
    text = re.sub(r'\\ref\{([^}]+)\}', r'[REF:{\1}]', text)
    text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\textsf\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\bm\{([^}]+)\}', r'\1', text)
    LEFT_DQ = chr(0x201C)
    RIGHT_DQ = chr(0x201D)
    text = text.replace("\\`", LEFT_DQ)
    text = text.replace("''", RIGHT_DQ)
    text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\songti', '', text)
    text = re.sub(r'\\zihao\{[^}]*\}', '', text)
    return text.strip()


# ============================================================
# EQUATION PARSING
# ============================================================
def parse_equations(text, chapter_number):
    """Extract equations from chapter text."""
    equations = []
    eq_num = 0
    pattern = r'\\begin\{equation\*?\}(.*?)\\label\{([^}]*)\}(.*?)\\end\{equation\*?\}'
    for m in re.finditer(pattern, text, re.DOTALL):
        eq_num += 1
        eq_content = (m.group(1) + m.group(3)).strip()
        eq_label = m.group(2).strip()
        equations.append({
            "chapter": chapter_number,
            "number": f"({chapter_number}.{eq_num})",
            "label": eq_label,
            "content": eq_content
        })
    return equations, eq_num


def _extract_eq_ref(match, equations):
    """Replace equation environment with [REF:eq:label] marker for positional tracking.
    If equation has no \\label{}, just return empty string."""
    body = match.group(0)
    # Try to find label in the equation body
    label_m = re.search(r'\\label\{([^}]+)\}', body)
    if label_m:
        label = label_m.group(1).strip()
        # Only insert REF if this equation is in our equations list
        for eq in equations:
            if eq['label'] == label:
                return f' [REF:eq:{label}] '
    return ''


# ============================================================
# FIGURE PARSING
# ============================================================
def parse_figures(text, chapter_number):
    """Extract figures from chapter text."""
    figures = []
    fig_pattern = r'\\begin\{figure\*?\}(.*?)\\end\{figure\*?\}'
    for m in re.finditer(fig_pattern, text, re.DOTALL):
        fig_block = m.group(1)

        # Extract includegraphics
        graphics = re.findall(r'\\includegraphics(?:\[[^\]]*\])?\{(.+?)\}', fig_block)

        # Extract caption
        cap_m = re.search(r'\\caption\{(.+?)\}', fig_block)
        caption = cap_m.group(1).strip() if cap_m else ""

        # Extract label
        lab_m = re.search(r'\\label\{([^}]+)\}', fig_block)
        label = lab_m.group(1).strip() if lab_m else ""

        # Extract subfigures
        subfigs = []
        for sf in re.finditer(r'\\subfigure(?:\[[^\]]*\])?\{(.*?)\}', fig_block, re.DOTALL):
            sf_content = sf.group(1)
            sf_graphics = re.findall(r'\\includegraphics(?:\[[^\]]*\])?\{(.+?)\}', sf_content)
            sf_cap_m = re.search(r'\\caption\{(.+?)\}', sf_content)
            subfigs.append({
                "graphics": sf_graphics,
                "caption": sf_cap_m.group(1).strip() if sf_cap_m else ""
            })

        figures.append({
            "chapter": chapter_number,
            "graphics": graphics,
            "caption": clean_inline_text(caption),
            "label": label,
            "subfigures": subfigs
        })
    return figures


# ============================================================
# TABLE PARSING
# ============================================================
def parse_tables(text, chapter_number):
    """Extract tables from chapter text."""
    tables = []
    tab_pattern = r'\\begin\{table\*?\}(.*?)\\end\{table\*?\}'
    for m in re.finditer(tab_pattern, text, re.DOTALL):
        tab_block = m.group(1)

        # Extract caption
        cap_m = re.search(r'\\caption\{(.+?)\}', tab_block)
        caption = cap_m.group(1).strip() if cap_m else ""

        # Extract label
        lab_m = re.search(r'\\label\{([^}]+)\}', tab_block)
        label = lab_m.group(1).strip() if lab_m else ""

        # Extract tabular content
        tabular_m = re.search(r'\\begin\{tabular\}\{([^}]*)\}(.*?)\\end\{tabular\}', tab_block, re.DOTALL)
        headers = []
        rows = []
        if tabular_m:
            tab_cols = tabular_m.group(1).strip()
            tab_body = tabular_m.group(2).strip()

            # Split by \\ but not inside braces
            raw_rows = re.split(r'\\\\\s*', tab_body)

            data_rows = []
            for rr in raw_rows:
                rr = rr.strip()
                if not rr:
                    continue
                # Remove \toprule, \midrule, \bottomrule, \hline
                rr = re.sub(r'\\toprule|\\midrule|\\bottomrule|\\hline', '', rr).strip()
                if not rr:
                    continue
                # Split by & 
                cells = [c.strip() for c in rr.split('&')]
                data_rows.append(cells)

            # First non-empty row is header if it contains text
            if data_rows:
                headers = data_rows[0]
                rows = data_rows[1:]

        tables.append({
            "chapter": chapter_number,
            "caption": clean_inline_text(caption),
            "label": label,
            "columns": tab_cols if 'tab_cols' in dir() else "",
            "headers": [clean_inline_text(h) for h in headers],
            "rows": [[clean_inline_text(c) for c in row] for row in rows]
        })
    return tables


# ============================================================
# REFERENCE PARSING
# ============================================================
def parse_references(text):
    """Extract references from thebibliography environment."""
    refs = []
    bib_pattern = r'\\bibitem\{([^}]+)\}(.*?)(?=\\bibitem\{|\\end\{thebibliography\})'
    for m in re.finditer(bib_pattern, text, re.DOTALL):
        key = m.group(1).strip()
        raw = m.group(2).strip()

        # Clean LaTeX from reference text
        # \url{...} -> keep URL
        raw = re.sub(r'\\url\{([^}]+)\}', r'\1', raw)
        # \newblock -> space
        raw = re.sub(r'\\newblock', ' ', raw)
        # \S -> keep
        raw = re.sub(r'\\S', '§', raw)
        # Remove other commands
        raw = re.sub(r'\\[a-zA-Z]+(\{[^}]*\})?', '', raw)
        # Collapse whitespace
        raw = re.sub(r'\s+', ' ', raw).strip()

        refs.append({"key": key, "text": raw})
    return refs


# ============================================================
# CHAPTER PARSING
# ============================================================
def parse_chapter_text(text):
    """
    Parse a chapter's LaTeX content into sections with paragraphs.
    Returns list of sections.
    """
    sections = []
    section_levels = []  # stack for tracking section hierarchy

    # Remove figure/table/equation environments first (they're parsed separately)
    # But we need to keep their placeholders for paragraph splitting
    # Strategy: extract them, replace with markers, then parse paragraphs

    # Split content into blocks at section headers
    # Pattern matches \section{...}, \subsection{...}, \subsubsection*{...}
    header_pattern = r'(\\section\*?\{[^}]*\}|\\subsection\*?\{[^}]*\}|\\subsubsection\*?\{[^}]*\})'

    parts = re.split(header_pattern, text)

    current_section = None

    for i, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue

        # Check if this part is a section header
        header_match = re.match(r'\\(sub)?(sub)?section\*?\{(.+?)\}', part)
        if header_match:
            # Fix: use group(2) to detect \subsubsection (double "sub")
            # group(1)=None → \section, group(1)="sub",group(2)=None → \subsection
            # group(1)="sub",group(2)="sub" → \subsubsection
            level = 1
            if header_match.group(1):  # \subsection or \subsubsection
                level += 1
            if header_match.group(2):  # \subsubsection (second "sub" group)
                level += 1
            title = header_match.group(3).strip()

            # Save previous section
            if current_section is not None:
                sections.append(current_section)

            current_section = {
                "level": level,
                "title": clean_inline_text(title),
                "content": []
            }
            continue

        # This part is content for the current section
        if current_section is None:
            # Content before any section header (chapter intro)
            current_section = {
                "level": 0,
                "title": "",
                "content": []
            }

        # Process paragraph and environment blocks
        # First, extract environments from the part text so they don't get split
        env_blocks = []
        
        # Find all \begin{...}...\end{...} blocks
        env_pattern = r'\\(begin)\{(figure|table|equation)\*?\}(.*?)\\(end)\{\2\*?\}'
        for m in re.finditer(env_pattern, part, re.DOTALL):
            env_type = m.group(2)
            env_content = m.group(3)
            
            if env_type == 'figure':
                fig_data = parse_figures(m.group(0), chapter_number=0)
                for fig in fig_data:
                    env_blocks.append(("figure", fig))
            elif env_type == 'table':
                tab_data = parse_tables(m.group(0), chapter_number=0)
                for tab in tab_data:
                    env_blocks.append(("table", tab))
            elif env_type == 'equation':
                label_m = re.search(r'\\label\{([^}]+)\}', env_content)
                label = label_m.group(1).strip() if label_m else ""
                eq_text = clean_inline_text(env_content)
                env_blocks.append(("equation", {"latex": eq_text, "label": label}))
        
        # Add environment blocks to section content
        for env_type, data in env_blocks:
            if env_type == 'figure':
                current_section["content"].append({
                    "type": "figure",
                    "images": [{"path": g} for g in data.get("graphics", [])],
                    "caption": data.get("caption", ""),
                    "label": data.get("label", "")
                })
            elif env_type == 'table':
                current_section["content"].append({
                    "type": "table",
                    "caption": data.get("caption", ""),
                    "label": data.get("label", ""),
                    "headers": data.get("headers", []),
                    "rows": data.get("rows", [])
                })
            elif env_type == 'equation':
                current_section["content"].append({
                    "type": "equation",
                    "latex": data.get("latex", ""),
                    "label": data.get("label", "")
                })
        
        # Now extract paragraph text (skip environment blocks)
        # Remove environments from part for paragraph parsing
        clean_part = re.sub(env_pattern, '', part, flags=re.DOTALL)
        
        paragraphs = re.split(r'\n\s*\n', clean_part)
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue

            cleaned = clean_latex_text(para)
            if cleaned:
                current_section["content"].append({
                    "type": "paragraph",
                    "text": cleaned,
                    "indent": True
                })

    # Don't forget the last section
    if current_section is not None:
        sections.append(current_section)

    return sections


def parse_chapter(text, chapter_number, chapter_title):
    """
    Parse a full chapter including equations, figures, tables, and sections.
    """
    # Extract equations
    equations, eq_count = parse_equations(text, chapter_number)

    # Extract figures
    figures = parse_figures(text, chapter_number)

    # Extract tables
    tables = parse_tables(text, chapter_number)

    # Remove equation/figure/table environments for section parsing
    # Insert [REF:eq:xxx] markers where equations were removed (for gen_round1.py interleaving)
    text_for_sections = text
    def _replace_eq_with_ref(m):
        """Replace equation env with [REF:label] marker if label exists.
        The label from LaTeX \label{eq:xxx} already includes the 'eq:' prefix,
        so the REF marker becomes [REF:eq:xxx] — matching float_map keys.
        
        \n\n wrapping preserves paragraph boundaries: text before/after
        the equation becomes separate paragraphs, and the REF marker
        paragraph is cleaned to empty in gen_round1.py.
        """
        eq_text = m.group(0)
        label_m = re.search(r'\\label\{([^}]+)\}', eq_text)
        if label_m:
            return f'\n\n [REF:{label_m.group(1).strip()}] \n\n'
        return '\n\n'
    text_for_sections = re.sub(
        r'\\begin\{equation\*?\}.*?\\end\{equation\*?\}',
        _replace_eq_with_ref,
        text_for_sections, flags=re.DOTALL)
    text_for_sections = re.sub(r'\\begin\{figure\*?\}.*?\\end\{figure\*?\}', '', text_for_sections, flags=re.DOTALL)
    text_for_sections = re.sub(r'\\begin\{table\*?\}.*?\\end\{table\*?\}', '', text_for_sections, flags=re.DOTALL)

    # Parse sections
    sections = parse_chapter_text(text_for_sections)

    # Assign section numbers
    section_counters = [0, 0, 0]
    for sec in sections:
        lvl = sec["level"]
        if lvl == 0:
            sec["number"] = ""
        elif lvl == 1:
            section_counters[0] += 1
            section_counters[1] = 0
            section_counters[2] = 0
            sec["number"] = f"{chapter_number}.{section_counters[0]}"
        elif lvl == 2:
            section_counters[1] += 1
            section_counters[2] = 0
            sec["number"] = f"{chapter_number}.{section_counters[0]}.{section_counters[1]}"
        elif lvl == 3:
            section_counters[2] += 1
            sec["number"] = f"{chapter_number}.{section_counters[0]}.{section_counters[1]}.{section_counters[2]}"

    chapter_data = {
        "number": chapter_number,
        "title": chapter_title,
        "sections": sections,
        "equations": equations,
        "figures": figures,
        "tables": tables
    }

    return chapter_data, eq_count, len(figures), len(tables)


# ============================================================
# MAIN PARSER
# ============================================================
def parse_thesis():
    """Main parsing function."""
    result = {
        "metadata": {},
        "abstract_cn": "",
        "abstract_cn_keywords": "",
        "abstract_en": "",
        "abstract_en_keywords": "",
        "chapters": [],
        "conclusion": [],
        "acknowledgement": [],
        "references": [],
        "appendix": {},
        "figure_paths": {}
    }

    # ---- Read all files ----
    print("Reading LaTeX files...")
    report_text = read_file(os.path.join(BASE_DIR, "report.tex"))
    com_info_text = read_file(os.path.join(BASE_DIR, "data", "com_info.tex"))
    bachelor_info_text = read_file(os.path.join(BASE_DIR, "data", "bachelor", "bachelor_info.tex"))
    abstract_text = read_file(os.path.join(BASE_DIR, "data", "abstract.tex"))
    ch1_text = read_file(os.path.join(BASE_DIR, "data", "chapter1-introduction.tex"))
    ch2_text = read_file(os.path.join(BASE_DIR, "data", "chapter2-failure.tex"))
    ch3_text = read_file(os.path.join(BASE_DIR, "data", "chapter3-seizure.tex"))
    ch4_text = read_file(os.path.join(BASE_DIR, "data", "chapter4-accuracy.tex"))
    conclusion_text = read_file(os.path.join(BASE_DIR, "data", "conclusion.tex"))
    acknowledgement_text = read_file(os.path.join(BASE_DIR, "data", "bachelor", "acknowledgement.tex"))
    references_text = read_file(os.path.join(BASE_DIR, "data", "references.tex"))
    # R74: Appendix files
    appendix_text = read_file(os.path.join(BASE_DIR, "data", "appendix.tex"))

    # ---- Metadata ----
    print("Parsing metadata...")
    result["metadata"] = parse_metadata(com_info_text, bachelor_info_text)

    # ---- Abstract ----
    print("Parsing abstract...")
    # Combine com_info (has keywords) with abstract text
    combined_abstract = abstract_text + "\n" + com_info_text
    abstract_data = parse_abstract(combined_abstract)
    result["abstract_cn"] = abstract_data["abstract_cn"]
    result["abstract_cn_keywords"] = abstract_data["abstract_cn_keywords"]
    result["abstract_en"] = abstract_data["abstract_en"]
    result["abstract_en_keywords"] = abstract_data["abstract_en_keywords"]

    # ---- Chapters ----
    print("Parsing chapters...")
    chapters_info = [
        (1, "绪论", ch1_text),
        (2, "机翼折叠机构失效模式分析", ch2_text),
        (3, "机构卡滞可靠性分析与基础代理模型", ch3_text),
        (4, "机构运动精度可靠性分析与主动克里金策略", ch4_text),
    ]

    total_eq = 0
    total_fig = 0
    total_tab = 0
    total_sec = 0

    for ch_num, ch_title, ch_text in chapters_info:
        ch_data, eq_count, fig_count, tab_count = parse_chapter(ch_text, ch_num, ch_title)
        result["chapters"].append(ch_data)
        total_eq += eq_count
        total_fig += fig_count
        total_tab += tab_count
        for sec in ch_data["sections"]:
            if sec["level"] >= 1:
                total_sec += 1

    # ---- Appendix (R74) ----
    # Parse appendix.tex into chapter-like structure.
    # gen_round1.py expects: appendix = { "sections": [...], "tables": [...] }
    # Each top-level \section becomes "附录A/B/C" heading, tables from \begin{table}
    # are parsed and stored at chapter level (parse_tables) then distributed by gen_round1.
    print("Parsing appendix...")
    if appendix_text and appendix_text.strip():
        # Remove \chapter{附录} header and \appendix command
        app_text = re.sub(r'\\chapter\*?\{[^}]*\}', '', appendix_text)
        app_text = app_text.replace('\\appendix', '')
        # Parse appendix as a pseudo-chapter. Use chapter_number=0 so
        # equation/table numbers don't conflict with chapters 1-4.
        # The table labels (e.g. tab:apx_matern) are used directly in ref_map.
        app_ch_data, app_eq, app_fig, app_tab = parse_chapter(app_text, 0, "附录")
        result["appendix"] = {
            "sections": app_ch_data["sections"],
            "tables": app_ch_data["tables"]
        }
        total_eq += app_eq
        total_fig += app_fig
        total_tab += app_tab
        for sec in app_ch_data["sections"]:
            if sec["level"] >= 1:
                total_sec += 1

    # ---- Conclusion ----
    print("Parsing conclusion...")
    # Remove chapter header
    conclusion_body = re.sub(r'\\chapter\*?\{[^}]*\}', '', conclusion_text)
    conclusion_body = re.sub(r'\\addcontentsline\{[^}]*\}\{[^}]*\}\{[^}]*\}', '', conclusion_body)
    conclusion_body = re.sub(r'\\markboth\{[^}]*\}\{[^}]*\}', '', conclusion_body)
    conclusion_paras = re.split(r'\n\s*\n', conclusion_body.strip())
    for para in conclusion_paras:
        para = para.strip()
        if para:
            cleaned = clean_latex_text(para)
            if cleaned:
                result["conclusion"].append(cleaned)

    # ---- Acknowledgement ----
    print("Parsing acknowledgement...")
    ack_body = re.sub(r'\\chapter\*?\{[^}]*\}', '', acknowledgement_text)
    ack_body = re.sub(r'\\addcontentsline\{[^}]*\}\{[^}]*\}\{[^}]*\}', '', ack_body)
    ack_body = re.sub(r'\\cleardoublepage', '', ack_body)
    ack_paras = re.split(r'\n\s*\n', ack_body.strip())
    for para in ack_paras:
        para = para.strip()
        if para:
            cleaned = clean_latex_text(para)
            if cleaned:
                result["acknowledgement"].append(cleaned)

    # ---- References ----
    print("Parsing references...")
    result["references"] = parse_references(references_text)

    # ---- Figure paths ----
    print("Building figure paths...")
    if os.path.isdir(FIGURE_DIR):
        for fname in os.listdir(FIGURE_DIR):
            fpath = os.path.join(FIGURE_DIR, fname)
            if os.path.isfile(fpath) and not fname.endswith('.pdf'):
                result["figure_paths"][fname] = fpath

    # ---- Write output ----
    print(f"\nWriting output to {OUTPUT_FILE}...")
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # ---- Summary ----
    print(f"\n{'='*60}")
    apx_info = ""
    if result.get("appendix") and result["appendix"].get("sections"):
        apx_info = f", {len(result['appendix']['sections'])} appendix sections"
    print(f"Parsed: 4 chapters, {total_sec} sections, {total_eq} equations, "
          f"{total_fig} figures, {total_tab} tables, {len(result['references'])} references{apx_info}")
    print(f"Output: {OUTPUT_FILE}")
    print(f"{'='*60}")

    return result


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    parse_thesis()
