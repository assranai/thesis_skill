# ⚠️ 新团队接手前必须先读: ../HANDOFF_遗嘱_必读.md (遗嘱 — 已完成的成果、关键教训、待处理任务、绝对不要做的事)
"""Round 1: Fix content ordering + format corrections from template record.

Key changes from v3:
1. Interleave figures/tables/equations based on [REF:label] markers in paragraph text
2. Format corrections: H1 line=480, H2 line=420, H2 before/after=160pt, H3 before/after=160pt
3. Better LaTeX cleanup (remove \chapter{}, \section{} commands in body text)
4. Body text indent: 0.85cm (matches template style 18)
"""
import json, os, re, sys, zipfile, base64
from lxml import etree

# Ensure stdout supports UTF-8 (fixes UnicodeEncodeError for ℓ, Greek, etc.
# when printing to Windows console with GBK codepage).
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass  # Some environments (older Python, restricted consoles) may not support reconfigure

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ooxml_generator import (
    make_heading, make_body_paragraph, make_body_paragraph_with_math,
    make_body_paragraph_with_refs, make_caption_with_bookmark,
    make_image_tag, get_image_dimensions, make_table,
    make_equation_omml_paragraph, make_paragraph, make_center_text,
    make_equation_block, _CITATION_RE, escape_xml
)

PROJ_DIR = r'E:\thesis\skill_test'
TEMPLATE_PATH = os.path.join(PROJ_DIR, '北航本科生论文模板-教务部2026年发.docx')
JSON_PATH = os.path.join(PROJ_DIR, 'skills', 'thesis-converter', 'scripts', 'thesis_content.json')
FIGURE_DIR = os.path.join(PROJ_DIR, 'latex模板', 'figure')
OUTPUT_PATH = os.path.join(PROJ_DIR, 'thesis_output.docx')
# OUTPUT_PATH = os.path.join(PROJ_DIR, 'output1', 'thesis_r21.docx')

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ============================================================
# Header/Footer XML with proper page number fields
# ============================================================
# footer5.xml (rId17) is used by ALL body sections (template sectPr #6 + _SECT_BREAK_XML).
# The template's footer5.xml is EMPTY — no page numbers at all.
# This replacement adds a centered PAGE field.
_FOOTER5_PAGE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>'''

# header8.xml is modified at runtime by modifying the template's original header XML:
# the static Roman numeral I (U+2160) is replaced with a PAGE field while preserving
# all template formatting (text box, logo image, fonts, borders, spacing).

# Section break XML: used between chapters so each chapter starts a new section.
# Uses the same page dimensions, header/footer references as the template body.
_SECT_BREAK_XML = (
    '<w:p><w:pPr><w:sectPr>'
    '<w:headerReference r:id="rId33" w:type="default"/>'
    '<w:footerReference r:id="rId17" w:type="default"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1701" w:right="1417" w:bottom="1417" w:left="1417" '
    'w:header="567" w:footer="851" w:gutter="0"/>'
    '<w:pgNumType w:fmt="decimal"/>'
    '<w:cols w:space="720" w:num="1"/>'
    '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"/>'
    '</w:sectPr></w:pPr>'
    '<w:r><w:br w:type="page"/></w:r></w:p>'
)

# Body-start section break XML: used BEFORE the first chapter (绪论) to create the
# body section boundary.  Unlike _SECT_BREAK_XML, this goes between the TOC paragraph
# and the first body heading.  The sectPr here defines the first body section with
# decimal numbering starting at 1.  No explicit page break — the sectPr boundary itself
# creates a clean section transition without inserting a blank page.
# R31: Body-start section break merged into first body paragraph to avoid blank page.
# The _BODY_START_XML sectPr is now a standalone <w:sectPr> snippet (no wrapping <w:p>)
# that gets injected into the pPr of the first body paragraph.
_BODY_START_SECTPR = (
    '<w:sectPr>'
    '<w:type w:val="nextPage"/>'
    '<w:headerReference r:id="rId33" w:type="default"/>'
    '<w:footerReference r:id="rId17" w:type="default"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1701" w:right="1417" w:bottom="1417" w:left="1417" '
    'w:header="567" w:footer="851" w:gutter="0"/>'
    '<w:pgNumType w:fmt="decimal" w:start="1"/>'
    '<w:cols w:space="720" w:num="1"/>'
    '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"/>'
    '</w:sectPr>'
)
# _BODY_START_XML kept for backward compat but body sectPr is now injected into
# the first body paragraph's pPr after body_parts are assembled.

# ============================================================
# LaTeX Cleanup
# ============================================================
def build_cite_map(references):
    """Build citation key → number map from references array."""
    cite_map = {}
    for i, ref in enumerate(references):
        key = ref.get('key', '')
        if key:
            cite_map[key] = i + 1
    return cite_map

def resolve_citations(text, cite_map):
    """Replace [REF:key] citation markers with [N] using cite_map.
    Handles multi-key: [REF:k1,k2] → [N,M]. Skips fig:/tab:/eq: float labels.
    Also consumes preceding ~ (LaTeX non-breaking space).
    Warns about undefined citation keys."""
    def replacer(m):
        keys_str = m.group(1).strip()
        # Skip float labels (fig:/tab:/eq:)
        if re.match(r'(fig:|tab:|eq:)', keys_str):
            return m.group(0)  # Leave float REFs for float interleaving
        # Split multiple keys
        keys = [k.strip() for k in keys_str.split(',')]
        nums = []
        for k in keys:
            if k in cite_map:
                nums.append(str(cite_map[k]))
            else:
                print(f'  [WARN] Undefined citation key: "{k}" — rendering as [{k}]')
                nums.append(k)  # Keep unknown keys as-is
        return '[' + ','.join(nums) + ']'
    return re.sub(r'~?\[REF:\s*\{?([a-zA-Z0-9_,:]+)\}?\]', replacer, text)

def clean_inline_latex(text):
    # Fix JSON-parsed LaTeX commands where the JSON escape ate the first letter.
    # The JSON has \bm (single backslash) instead of \\bm, so \b → backspace (0x08).
    # Fix known cases first, then blanket-restore remaining backslash-equivalent control chars.
    text = text.replace('\x08m', '\\bm')    # \bm (most common corrupted command)
    text = text.replace('\x08', '\\')        # Remaining \b → backslash
    text = text.replace('\x0c', '\\')        # \f → backslash (e.g. \frac)
    text = re.sub(r'\r([a-zA-Z])', r'\\\1', text)  # \r in LaTeX cmds (e.g. \mathrm)
    text = re.sub(r'\t([a-zA-Z])', r'\\\1', text)  # \t in LaTeX cmds (e.g. \textbf)
    text = text.replace('\r', '')            # Remove remaining standalone \r (line endings)
    text = text.replace('\t', '')            # Remove remaining standalone \t (rare)
    # Handle double-escaped JSON quotes (\`\` → left double quote, \'\' → right double quote)
    text = text.replace('\\`\\`', '\u201c').replace("\\'\\'", '\u201d')
    text = text.replace('``', '\u201c').replace("''", '\u201d')
    # Convert straight ASCII double quotes " → Chinese curly quotes (alternating left/right).
    # The conclusion chapter stores Chinese quotation marks as plain " in JSON (escaped as \").
    # Other chapters use LaTeX ``/'' which are already handled above.
    parts = text.split('"')
    if len(parts) >= 3:
        result = [parts[0]]
        for i in range(1, len(parts)):
            result.append('\u201c' if i % 2 == 1 else '\u201d')
            result.append(parts[i])
        text = ''.join(result)
    text = re.sub(r'(?<!\\)%.*', '', text)
    text = re.sub(r'~?\\upcite\{[^}]*\}', '', text)
    text = re.sub(r'\\cite\{[^}]*\}', '', text)
    text = re.sub(r'\\(chapter|section|subsection|subsubsection)\s*\{[^}]*\}?', '', text)
    # R65: Remove LaTeX environment remnants aggressively
    # Handle \begin{word}, \begin{word\n} (partial with missing }), \end{word}, \item
    text = re.sub(r'\\begin\{[^}]*\}?', '', text)    # \begin{...}
    text = re.sub(r'\\begin\s*\{', '', text)          # partial \begin{ without }
    text = re.sub(r'\\end\{[^}]*\}?', '', text)       # \end{...}
    text = re.sub(r'\\end\s*\{', '', text)             # partial \end{ without }
    text = re.sub(r'\\item\b', '', text)               # \item
    # R66: Convert \ref{...} → [REF:...] for cross-reference resolution
    text = re.sub(r'\\ref\{([^}]+)\}', r'[REF:\1]', text)
    text = re.sub(r'% !Mode::.*?\n', '', text)
    text = re.sub(r'% !Version::.*?\n', '', text)
    text = re.sub(r'% !LastMod::.*?\n', '', text)
    text = re.sub(r'% !Status::.*?\n', '', text)
    
    # Strip \% (escaped percent) to bare % — the backslash is just LaTeX escaping,
    # not a letter command, so the [a-zA-Z]+ regex below can't catch it.
    # Must happen after comment-line removal (above) but before $...$ math
    # preservation (below), so that \% inside math regions is also corrected
    # (the OMML generator handles bare % correctly in both primary and fallback paths).
    text = text.replace(r'\%', '%')
    
    # R41: Strip manual numbering prefixes from heading-like text.
    # Remove ".", "。" from delimiter set — they eat decimal points in table data
    # (e.g. "380.000" → "380" matches, leaving just "000").
    text_stripped = text.strip()
    if re.match(r'^\d+[)）、]\s*', text_stripped):
        text = re.sub(r'^\d+[)）、]\s*', '', text_stripped)
    
    # Preserve $...$ and $$...$$ math segments so destructive cleanup
    # (brace stripping, backslash stripping) does NOT mangle math content.
    # These are restored after all cleanup operations complete.
    MATH_PH = '\x00__MATH__\x00'
    math_blocks = []
    def _save_math(m):
        math_blocks.append(m.group(0))
        return f'{MATH_PH}{len(math_blocks)-1}\x00'
    text = re.sub(r'\$\$.*?\$\$|\$.*?\$', _save_math, text, flags=re.DOTALL)
    
    greek_map = {
        'alpha': '\u03b1','beta': '\u03b2','gamma': '\u03b3','delta': '\u03b4',
        'epsilon': '\u03b5','zeta': '\u03b6','eta': '\u03b7','theta': '\u03b8',
        'iota': '\u03b9','kappa': '\u03ba','lambda': '\u03bb','mu': '\u03bc',
        'nu': '\u03bd','xi': '\u03be','pi': '\u03c0','rho': '\u03c1',
        'sigma': '\u03c3','tau': '\u03c4','upsilon': '\u03c5','phi': '\u03c6',
        'chi': '\u03c7','psi': '\u03c8','omega': '\u03c9',
        'Gamma': '\u0393','Delta': '\u0394','Theta': '\u0398','Lambda': '\u039b',
        'Xi': '\u039e','Pi': '\u03a0','Sigma': '\u03a3','Phi': '\u03a6',
        'Psi': '\u03a8','Omega': '\u03a9',
    }
    op_map = {
        'times': '\u00d7','cdot': '\u00b7','pm': '\u00b1',
        'leq': '\u2264','geq': '\u2265','approx': '\u2248','neq': '\u2260',
        'infty': '\u221e','forall': '\u2200','exists': '\u2203',
        'rightarrow': '\u2192','Rightarrow': '\u21d2','leftarrow': '\u2190',
        'leftrightarrow': '\u2194','Leftrightarrow': '\u21d4','Leftarrow': '\u21d0',
        'mapsto': '\u21a6','longrightarrow': '\u27f6','longleftarrow': '\u27f5',
        'to': '\u2192','partial': '\u2202','nabla': '\u2207',
        'cdots': '\u22ef','ldots': '\u2026','prime': '\u2032',
        'langle': '\u27e8','rangle': '\u27e9','in': '\u2208','notin': '\u2209',
        'subset': '\u2282','subseteq': '\u2286','cup': '\u222a','cap': '\u2229',
        'sim': '\u223c','equiv': '\u2261','propto': '\u221d','parallel': '\u2225',
        'backslash': '\\',
        'circ': '\u2218','triangle': '\u25b3','angle': '\u2220',
        'Vert': '\u2016',  # \| (norm): double vertical bar
    }
    all_cmds = {}
    all_cmds.update(greek_map)
    all_cmds.update(op_map)
    
    text = text.replace('{', '').replace('}', '')
    for _ in range(5):
        # Pre-process: replace \| (norm) with Unicode double vertical bar
        text = text.replace(r'\|', '\u2016')
        for cmd, sym in sorted(all_cmds.items(), key=lambda x: -len(x[0])):
            text = text.replace(f'\\{cmd}', sym)
        text = re.sub(r'\\(textbf|textit|textsf|mathrm|mathbf|mathit|mathbb|mathcal|mathfrak|mathsf|mathtt)\{([^}]*)\}', r'\2', text)
        text = re.sub(r'\\bm\{([^{}]*)\}', r'\1', text)
        text = re.sub(r'\\([a-zA-Z]+)', r'\1', text)
    
    # Restore preserved math segments ($...$ and $$...$$)
    for i, seg in enumerate(math_blocks):
        text = text.replace(f'{MATH_PH}{i}\x00', seg)
    
    # Collapse LaTeX line breaks — newlines are for source readability in .tex files,
    # not paragraph breaks. In PDF output they merge seamlessly (treated as spaces),
    # but in Word they produce visible extra spaces between lines.
    # Must happen after math restoration to leave $...$ content untouched.
    # Step 1: Gemini fix — remove newlines between Chinese characters (no space, matches PDF)
    text = re.sub(r'([\u4e00-\u9fa5])\s*\n\s*([\u4e00-\u9fa5])', r'\1\2', text)
    # Step 2: Remaining newlines (non-CJK contexts) → single space
    text = re.sub(r'\s*\n\s*', ' ', text)
    text = re.sub(r'  +', ' ', text)
    
    # --- R18: LaTeX spacing cleanup (after math restoration, safe for $...$) ---
    # Collapse any remaining multiple spaces (belt-and-suspenders)
    text = re.sub(r' {2,}', ' ', text)
    # Remove LaTeX thin-space commands — invisible in PDF but literal chars in Word
    text = text.replace(r'\,', '').replace(r'\;', '').replace(r'\:', '').replace(r'\!', '')
    # Convert LaTeX spacing commands to ordinary spaces
    text = text.replace(r'\enspace', ' ').replace(r'\quad', '  ')
    # Convert LaTeX non-breaking space ~ to ordinary space
    text = text.replace('~', ' ')
    
    # Decode HTML entities (order matters: &amp; first to handle &amp;quot; etc.)
    text = text.replace('&amp;', '&').replace('&quot;', '"')
    text = text.replace('&lt;', '<').replace('&gt;', '>')
    return text.strip()

_KNOWN_MATH_VARS_RE = None

def _wrap_math_variables(text):
    """Detect unmarked math variables in cleaned text and wrap in $...$.
    
    Heuristic detection for patterns like:
      - Greek letter with subscript: σ_u, σ_y, τ_u
      - Standalone Greek letters: σ, ρ, μ, ν  
      - Superscript patterns: E², G², σ²
      - Known single-letter math symbols: E, G (uppercase)
    
    Uses ASCII [A-Za-z0-9_] for word-boundary detection instead of
    Python-3 \w (which includes CJK ideographs as word characters).
    This ensures variables after Chinese text (e.g. "模量E", "密度ρ")
    are properly detected and wrapped in math mode.
    
    These would otherwise render as plain text instead of italic math.
    """
    import re as _re
    global _KNOWN_MATH_VARS_RE
    
    if _KNOWN_MATH_VARS_RE is None:
        # Unicode Greek letter ranges
        greek = '[α-ωΑ-Ωϵϑϰϱϕ]'
        # Known uppercase math symbols
        known_upper = '[EGVR]'
        # ASCII word chars only — NOT Python-3 \w which includes CJK.
        # Plain characters without brackets, used inside existing [...]
        # character classes to avoid nested-set FutureWarning.
        _W = 'A-Za-z0-9_'
        
        # Build compound pattern: tries subscript pattern first, then standalone
        _KNOWN_MATH_VARS_RE = _re.compile(
            # Greek + subscript with braces: σ_{max}  τ_{ij}
            '(' + greek + r')_\{\s*([^}]+)\s*\}'
            # Greek + single-char subscript: σ_u  σ_y  τ_u
            r'|(' + greek + r')_([A-Za-z0-9])'
            # Standalone Greek: σ  ρ  μ  ν  (CJK-safe boundaries)
            r'|(?<![' + _W + r'$])(' + greek + r')(?![' + _W + r'$])'
            # Greek + superscript: σ²  ρ³
            r'|(' + greek + r')([²³¹⁰⁴⁵⁶⁷⁸⁹]+)'
            # Known uppercase var: E  G  (CJK-safe boundaries)
            r'|(?<![' + _W + r'$])(' + known_upper + r')(?![' + _W + r'$])',
            _re.UNICODE
        )
    
    def _replacer(m):
        """Map match groups → $...$ wrapped math."""
        if m.group(1):   # Greek + subscript with braces: σ_{max}
            return '$' + m.group(1) + '_{' + m.group(2) + '}$'
        if m.group(3):   # Greek + single-char subscript: σ_u
            return '$' + m.group(3) + '_{' + m.group(4) + '}$'
        if m.group(5):   # Standalone Greek: σ, ρ
            return '$' + m.group(5) + '$'
        if m.group(6):   # Greek + superscript: σ²
            return '$' + m.group(6) + '^{' + m.group(7) + '}$'
        if m.group(8):   # Known uppercase: E, G
            return '$' + m.group(8) + '$'
        return m.group(0)
    
    return _KNOWN_MATH_VARS_RE.sub(_replacer, text)


def clean_cell_text(text):
    """Clean table cell text, preserving $...$ math markers for OMML rendering.
    
    Non-math parts are cleaned via clean_inline_latex (LaTeX commands → Unicode).
    Math parts ($...$) are preserved as raw LaTeX for the OMML generator.
    After cleaning, unmarked math variables are auto-wrapped in $...$
    via _wrap_math_variables().
    """
    import re as _re
    if '$' not in text:
        cleaned = clean_inline_latex(text)
        return _wrap_math_variables(cleaned)
    
    # Split by $...$, process non-math parts only
    parts = []
    last_end = 0
    for m in _re.finditer(r'\$(.*?)\$', text):
        if m.start() > last_end:
            prefix = clean_inline_latex(text[last_end:m.start()])
            prefix = _wrap_math_variables(prefix)
            parts.append(prefix)
        parts.append('$' + m.group(1) + '$')
        last_end = m.end()
    if last_end < len(text):
        suffix = clean_inline_latex(text[last_end:])
        suffix = _wrap_math_variables(suffix)
        parts.append(suffix)
    
    return ''.join(parts)

# ============================================================
# REF marker detection for content ordering
# ============================================================
REF_MARKER_RE = re.compile(r'\[REF:\s*\{?([a-zA-Z_:0-9]+)\]?\s*\]', re.IGNORECASE)
# Only labels starting with fig:, tab:, eq: are floats; others are citation refs
# Matches: [REF:{fig:mech_diagram], [REF:fig:mech_diagram], [REF:{tab:xxx], etc.

def find_ref_markers(text):
    """Find all [REF:label] markers in text. Returns list of (label, start, end)."""
    markers = []
    for m in REF_MARKER_RE.finditer(text):
        label = m.group(1).strip()
        markers.append((label, m.start(), m.end()))
    return markers

def remove_float_ref_markers(text):
    """Remove only float-related [REF:fig:xxx], [REF:tab:xxx], [REF:eq:xxx] markers."""
    # Remove only markers with fig:/tab:/eq: prefix
    return re.sub(r'\[REF:\s*\{?(fig:|tab:|eq:)[a-zA-Z_:0-9]+\]?\s*\]', '', text, flags=re.IGNORECASE).strip()

def build_float_map(chapter):
    """Build a map from label (fig:xxx, tab:xxx, eq:xxx) to content."""
    float_map = {}
    for obj in chapter.get('figures', []):
        label = obj.get('label', '')
        if label:
            float_map[label] = ('figure', obj)
    for obj in chapter.get('tables', []):
        label = obj.get('label', '')
        if label:
            float_map[label] = ('table', obj)
    for obj in chapter.get('equations', []):
        label = obj.get('label', '')
        if label:
            float_map[label] = ('equation', obj)
    return float_map

# ============================================================
# Format-corrected heading and body generators
# ============================================================
def make_heading_r1(level, title, back_matter=False):
    """Heading wrapper — delegates to ooxml_generator.make_heading().
    
    ooxml_generator now produces correct formats directly:
      H1 (body):     sz=32 (三号), 黑体, UNBOLD, centered, line=480, before=500, after=500
      H1 (back):     sz=32 (三号), 黑体, UNBOLD, centered, line=360, before=240, after=240
      H2:            sz=24 (小四), 黑体, UNBOLD, left, line=420
      H3:            sz=24 (小四), 黑体, UNBOLD, left
      H4:            sz=24 (小四), 黑体, UNBOLD, left
    No post-patching needed.
    """
    if back_matter:
        return make_heading(level, title, back_matter=True)
    return make_heading(level, title)

def generate_equation(latex_content, eq_number):
    """Generate equation paragraph."""
    try:
        return make_equation_omml_paragraph(latex_content, eq_number=eq_number)
    except Exception as e:
        print(f'  [WARN] Eq OMML failed: {e}')
        return make_equation_block(latex_content, eq_number)

def generate_paragraph_r1(text, indent=True, ref_map=None, cite_map=None):
    """Generate paragraph with LaTeX cleanup, citations, math, refs."""
    # Gemini fix: remove newlines between Chinese characters (LaTeX source readability)
    # Must happen BEFORE strip() and any other processing to collapse CJK text across lines
    text = re.sub(r'([\u4e00-\u9fa5])\s*\n\s*([\u4e00-\u9fa5])', r'\1\2', text)
    text = text.strip()
    if not text:
        return ''
    
    text = text.replace('\x08', '\\')
    text = text.replace('\x0c', '\\')
    text = text.replace('\r', '')
    # Resolve any remaining citation REFs (belt-and-suspenders)
    if cite_map:
        text = resolve_citations(text, cite_map)
    # Only convert float-type REF markers to \ref{} (citations already resolved)
    text = re.sub(r'\[REF:\s*\{?(fig:|tab:|eq:)([a-zA-Z_:0-9]+)\}?\]', r'\\ref{\1\2}', text)
    
    has_math = '$' in text
    has_ref = '\\ref{' in text
    ref_replaced_text = text
    
    if has_ref and ref_map:
        # R72: Replace \ref{...} with static number text (e.g., "表3.1")
        # Avoids Word REF field formatting inheritance issues (sz=21, bold from captions)
        def _replace_ref(m):
            label = m.group(1)
            if label in ref_map:
                return ref_map[label]['num']
            return m.group(0)
        ref_replaced_text = re.sub(r'\\ref\{([^}]+)\}', _replace_ref, text)
    
    if has_math:
        return make_body_paragraph_with_math(ref_replaced_text, indent=indent, ref_map=ref_map)
    elif has_ref and ref_map:
        return make_body_paragraph(ref_replaced_text, indent=indent)
    else:
        return make_body_paragraph(ref_replaced_text, indent=indent)

# ============================================================
# Main generation with interleaved content ordering
# ============================================================
print("=" * 60)
print("BUAA Thesis Converter — Round 1 (Content Ordering)")
print("=" * 60)

print("\n[1] Loading thesis content...")
with open(JSON_PATH, 'r', encoding='utf-8') as f:
    thesis = json.load(f)
print(f"  {len(thesis['chapters'])} chapters, "
      f"{sum(len(c.get('figures',[])) for c in thesis['chapters'])} figures, "
      f"{sum(len(c.get('tables',[])) for c in thesis['chapters'])} tables, "
      f"{sum(len(c.get('equations',[])) for c in thesis['chapters'])} equations")

print("\n[1.5] Building citation map...")
cite_map = build_cite_map(thesis.get('references', []))
print(f"  {len(cite_map)} citation keys mapped")

print("\n[2] Building body content with interleaved ordering...")
# Build reference map for cross-references
ref_map = {}
for ch in thesis['chapters']:
    cn = ch['number']
    for fi, fig in enumerate(ch.get('figures', []), 1):
        if fig.get('label'):
            ref_map[fig['label']] = {'bm': f'_Fig_{cn}_{fi}', 'num': f'{cn}.{fi}'}
    for ti, tbl in enumerate(ch.get('tables', []), 1):
        if tbl.get('label'):
            ref_map[tbl['label']] = {'bm': f'_Tbl_{cn}_{ti}', 'num': f'{cn}.{ti}'}

# R67: Add appendix items to ref_map for \ref resolution
appendix = thesis.get('appendix', {})
if appendix:
    app_label = ord('A')
    for section in appendix.get('sections', []):
        for ti, tbl in enumerate(section.get('tables', []), 1):
            if tbl.get('label'):
                ref_map[tbl['label']] = {'bm': f'_Tab_{chr(app_label)}_{ti}', 'num': f'{chr(app_label)}.{ti}'}
        app_label += 1
    for ti, tbl in enumerate(appendix.get('tables', []), 1):
        if tbl.get('label'):
            ref_map[tbl['label']] = {'bm': f'_Tab_App_{ti}', 'num': f'\u9644.{ti}'}

body_parts = []
img_id_counter = [1]
bm_id_counter = [1000]
fig_counter = {}
tbl_counter = {}
eq_counter = {}
used_floats = set()  # Track which floats have been placed

for chapter in thesis['chapters']:
    ch_num = chapter['number']
    ch_title = clean_inline_latex(chapter['title'])
    
    fig_counter.setdefault(ch_num, 0)
    tbl_counter.setdefault(ch_num, 0)
    eq_counter.setdefault(ch_num, 0)
    
    # Build float map for this chapter
    float_map = build_float_map(chapter)
    _pending_floats = []  # R69: defer float insertion until AFTER paragraph
    
    # Page break + section break before chapter.
    # First chapter gets a body-start section break (decimal, start=1, page break)
    # that creates the body section boundary WITHOUT putting a sectPr on the heading
    # itself (which would push body text to the next page).
    # Subsequent chapters get the standard section break.
    if ch_num == 1:
        # Inject body sectPr into first body paragraph's pPr instead of as separate paragraph
        # This eliminates the blank page between TOC and body
        first_elem = body_parts[0] if body_parts else ''
        if first_elem and '<w:pPr>' in first_elem:
            body_parts[0] = first_elem.replace('</w:pPr>', _BODY_START_SECTPR + '</w:pPr>', 1)
            print("  Body sectPr merged into first paragraph (no blank page)")
        else:
            body_parts.insert(0, _BODY_START_SECTPR)  # fallback
    elif ch_num > 1:
        body_parts.append(_SECT_BREAK_XML)
    
    body_parts.append(make_heading_r1(1, ch_title))
    
    # Process sections with interleaved floats
    for section in chapter['sections']:
        sec_level = section.get('level', 0)
        sec_title = clean_inline_latex(section.get('title', ''))
        
        if sec_level >= 1:
            h_level = min(sec_level + 1, 4)
            body_parts.append(make_heading_r1(h_level, sec_title))
        
        for block in section.get('content', []):
            if block['type'] != 'paragraph':
                continue
            
            text = clean_inline_latex(block.get('text', ''))
            # Gemini fix: collapse newlines between Chinese characters
            text = re.sub(r'([\u4e00-\u9fa5])\s*\n\s*([\u4e00-\u9fa5])', r'\1\2', text)
            indent = block.get('indent', True)
            
            if not text.strip():
                continue
            
            # Resolve citation REF markers to [N] BEFORE float REF detection
            # (cite_map skips fig:/tab:/eq: labels, so float REFs stay intact)
            text = resolve_citations(text, cite_map)
            
            # Check for REF markers in this paragraph
            markers = find_ref_markers(text)
            for label, _, _ in markers:
                # Only treat fig:/tab:/eq: prefixed labels as image/table/equation floats
                is_float = label.startswith(('fig:', 'tab:', 'eq:'))
                if is_float and label in float_map and label not in used_floats:
                    ftype, fobj = float_map[label]
                    used_floats.add(label)
                    
                    # R72: Equations go BEFORE paragraph (formula definition),
                    # tables/figures go AFTER paragraph (natural reading order)
                    if ftype == 'equation':
                        eq_counter[ch_num] += 1
                        eq_content = fobj.get('content', fobj.get('latex', ''))
                        eq_number = f'({ch_num}-{eq_counter[ch_num]})'
                        if eq_content.strip():
                            body_parts.append(generate_equation(eq_content, eq_number))
                    else:
                        _pending_floats.append((ftype, fobj))
            
            # Convert float REF markers to \ref{...} for cross-referencing
            # fig:/tab: → \ref{} for Word REF field
            clean_text = re.sub(r'\[REF:\s*\{?(fig:|tab:)([a-zA-Z_:0-9]+)\}?\]', r'\\ref{\1\2}', text)
            # eq: markers removed from display text (equations placed inline, no ref needed)
            clean_text = re.sub(r'\[REF:\s*\{?eq:[a-zA-Z_:0-9]+\}?\]', ' ', clean_text, flags=re.IGNORECASE)
            clean_text = re.sub(r'  +', ' ', clean_text).strip()
            if clean_text.strip():
                body_parts.append(generate_paragraph_r1(clean_text, indent=indent, ref_map=ref_map, cite_map=cite_map))
            # R72: Insert pending tables/figures AFTER this paragraph (equations already placed BEFORE)
            for ftype, fobj in _pending_floats:
                if ftype == 'figure':
                    fig_counter[ch_num] += 1
                    fi = fig_counter[ch_num]
                    graphics = fobj.get('graphics', [])
                    caption = clean_inline_latex(fobj.get('caption', ''))
                    if graphics:
                        for gfx in graphics:
                            if gfx:
                                img_id_counter[0] += 1
                                body_parts.append(make_image_tag(f'rId_img_{gfx}', img_path=os.path.join(FIGURE_DIR, gfx), img_id=img_id_counter[0]))
                        bm_id_counter[0] += 1
                        body_parts.append(make_caption_with_bookmark(f'\u56fe{ch_num}.{fi} {caption}', f'_Fig_{ch_num}_{fi}', bm_id_counter[0]))
                elif ftype == 'table':
                    tbl_counter[ch_num] += 1
                    ti = tbl_counter[ch_num]
                    tbl_caption = clean_inline_latex(fobj.get('caption', ''))
                    tbl_headers = [clean_cell_text(h) for h in fobj.get('headers', [])]
                    tbl_rows = [[clean_cell_text(c) for c in row] for row in fobj.get('rows', [])]
                    bm_id_counter[0] += 1
                    body_parts.append(make_table(tbl_headers, tbl_rows, caption_text=f'\u8868{ch_num}.{ti} {tbl_caption}',
                                                 bookmark_name=f'_Tbl_{ch_num}_{ti}', bookmark_id=bm_id_counter[0]))
                elif ftype == 'equation':
                    eq_counter[ch_num] += 1
                    eq_content = fobj.get('content', fobj.get('latex', ''))
                    eq_number = f'({ch_num}-{eq_counter[ch_num]})'
                    body_parts.append(generate_equation(eq_content, eq_number))
            _pending_floats.clear()
    
    # Place any remaining unmatched floats at the end of the chapter
    # R68: Preserve LaTeX source order (figures + tables interleaved, not grouped by type)
    all_floats = (
        [(obj.get('label',''), 'figure', obj) for obj in chapter.get('figures', [])] +
        [(obj.get('label',''), 'table', obj) for obj in chapter.get('tables', [])] +
        [(obj.get('label',''), 'equation', obj) for obj in chapter.get('equations', [])]
    )
    # Sort by position in JSON (preserves source order)
    for label, ftype, fobj in all_floats:
        if not label or label in used_floats:
            continue
        used_floats.add(label)
        
        if ftype == 'figure':
            fig_counter[ch_num] += 1
            fi = fig_counter[ch_num]
            graphics = fobj.get('graphics', [])
            caption = clean_inline_latex(fobj.get('caption', ''))
            if graphics:
                for gfx in graphics:
                    if gfx:
                        img_id_counter[0] += 1
                        body_parts.append(make_image_tag(f'rId_img_{gfx}', img_path=os.path.join(FIGURE_DIR, gfx), img_id=img_id_counter[0]))
                bm_id_counter[0] += 1
                body_parts.append(make_caption_with_bookmark(f'\u56fe{ch_num}.{fi} {caption}', f'_Fig_{ch_num}_{fi}', bm_id_counter[0]))
        
        elif ftype == 'table':
            tbl_counter[ch_num] += 1
            ti = tbl_counter[ch_num]
            tbl_caption = clean_inline_latex(fobj.get('caption', ''))
            tbl_headers = [clean_cell_text(h) for h in fobj.get('headers', [])]
            tbl_rows = [[clean_cell_text(c) for c in row] for row in fobj.get('rows', [])]
            bm_id_counter[0] += 1
            body_parts.append(make_table(tbl_headers, tbl_rows, caption_text=f'\u8868{ch_num}.{ti} {tbl_caption}',
                                         bookmark_name=f'_Tbl_{ch_num}_{ti}', bookmark_id=bm_id_counter[0]))
        
        elif ftype == 'equation':
            eq_counter[ch_num] += 1
            eq_content = fobj.get('content', fobj.get('latex', ''))
            eq_number = f'({ch_num}-{eq_counter[ch_num]})'
            if eq_content.strip():
                body_parts.append(generate_equation(eq_content, eq_number))

# Conclusion (new section)
body_parts.append(_SECT_BREAK_XML)
body_parts.append(make_heading_r1(1, '\u7ed3\u8bba', back_matter=True))
for item in thesis.get('conclusion', []):
    text = clean_inline_latex(item if isinstance(item, str) else item.get('text', ''))
    if text.strip():
        body_parts.append(generate_paragraph_r1(text, indent=True, cite_map=cite_map))

body_parts.append(make_heading_r1(1, '\u81f4\u8c22', back_matter=True))
for item in thesis.get('acknowledgement', []):
    text = clean_inline_latex(item if isinstance(item, str) else item.get('text', ''))
    if text.strip():
        body_parts.append(generate_paragraph_r1(text, indent=True, cite_map=cite_map))

# R37: Page break before references (separates from acknowledgement).
# Simple <w:p> wrapper may be dropped by lxml — safer to embed <w:br> in first run
# R37: Page break before references + R41: chapter-level heading format
body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
body_parts.append(make_heading_r1(1, '参考文献', back_matter=True))
for idx, ref in enumerate(thesis.get('references', []), 1):
    ref_text = clean_inline_latex(ref.get('text', ref if isinstance(ref, str) else ''))
    if ref_text:
        # R40: Manual [N] prefix — avoids ALL numbering conflicts with heading styles.
        # Auto-numbering (numId) shares Word's flat numbering space and inevitably
        # interferes with chapter/section numbering defined via heading styles.
        body_parts.append(make_paragraph('17', f'[{idx}] {ref_text}', font_size=24, detect_citations=False))

# R66: Appendix content — 附录A/附录B as chapter-level headings
appendix = thesis.get('appendix', {})
if appendix and appendix.get('sections'):
    # R71: Distribute chapter-level tables into sections (avoid all in last appendix)
    _ch_tables = appendix.get('tables', [])
    _ch_ti = 0
    app_label = ord('A')  # R74: reset counter (shared with ref_map code above)
    for section in appendix.get('sections', []):
        title = section.get('title', '').strip()
        if not title:
            continue
        title_clean = clean_inline_latex(title)
        body_parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        body_parts.append(make_heading_r1(1, f'\u9644\u5f55{chr(app_label)}  {title_clean}', back_matter=True))
        
        for block in section.get('content', []):
            if block['type'] == 'paragraph':
                text = clean_inline_latex(block.get('text', ''))
                text = re.sub(r'([\u4e00-\u9fa5])\s*\n\s*([\u4e00-\u9fa5])', r'\1\2', text)
                if text.strip():
                    body_parts.append(generate_paragraph_r1(text, indent=block.get('indent', True), cite_map=cite_map, ref_map=ref_map))
        # Section-level tables
        for ti, tbl in enumerate(section.get('tables', []), 1):
            tbl_headers = [clean_cell_text(h) for h in tbl.get('headers', [])]
            tbl_rows = [[clean_cell_text(c) for c in row] for row in tbl.get('rows', [])]
            tbl_caption = tbl.get('caption', '')
            body_parts.append(make_table(tbl_headers, tbl_rows, caption_text=f'\u8868{chr(app_label)}.{ti} {tbl_caption}', bookmark_name=f'_Tab_{chr(app_label)}_{ti}', bookmark_id=bm_id_counter[0]))
        # R71: Assign next chapter-level table to this section (distribute evenly)
        if _ch_ti < len(_ch_tables):
            tbl = _ch_tables[_ch_ti]
            _ch_ti += 1
            tbl_headers = [clean_cell_text(h) for h in tbl.get('headers', [])]
            tbl_rows = [[clean_cell_text(c) for c in row] for row in tbl.get('rows', [])]
            tbl_caption = tbl.get('caption', '')
            body_parts.append(make_table(tbl_headers, tbl_rows, caption_text=f'\u8868{chr(app_label)}.{len(section.get("tables",[]))+1} {tbl_caption}', bookmark_name=f'_Tab_{chr(app_label)}_{len(section.get("tables",[]))+1}', bookmark_id=bm_id_counter[0]))
        app_label += 1
    # Any remaining chapter-level tables
    while _ch_ti < len(_ch_tables):
        tbl = _ch_tables[_ch_ti]
        _ch_ti += 1
        tbl_headers = [clean_cell_text(h) for h in tbl.get('headers', [])]
        tbl_rows = [[clean_cell_text(c) for c in row] for row in tbl.get('rows', [])]
        tbl_caption = tbl.get('caption', '')
        last_lbl = chr(ord('A') + len(appendix['sections']) - 1)
        body_parts.append(make_table(tbl_headers, tbl_rows, caption_text=f'\u8868\u9644.{_ch_ti} {tbl_caption}'))
    print("  Appendix: 附录A/B H1 format + tables")

body_xml = '\n'.join(body_parts)
print(f"  Generated {len(body_parts)} body elements")

# ============================================================
# Inject into template + images + numbering (same as v3)
# ============================================================
print("\n[3] Injecting body into template...")

with zipfile.ZipFile(TEMPLATE_PATH, 'r') as zf:
    template_files = {n: zf.read(n) for n in zf.namelist()}

# R34: Page number fix.
# The template's header8.xml (text box) already contains a PAGE field — this is the
# ONLY place page numbers should appear. The footer must remain empty.
# Previous rounds injected a redundant PAGE field into footer5.xml, creating
# duplicated page numbers at the bottom of each page. This injection is removed.
# The original empty footer5.xml (one empty paragraph, style 23) is preserved.
# Page number FORMAT is controlled by sectPr pgNumType:
#   Front matter → upperRoman (I, II, III...) 
#   Body → decimal start=1 (1, 2, 3...)
# The header's PAGE field automatically picks up the correct format per section.
# Modify header8 at runtime: replace static Roman numeral I (U+2160) with PAGE field,
# preserving ALL template formatting (text box, logo drawing, fonts, borders, spacing).
header8_xml = template_files['word/header8.xml'].decode('utf-8')
roman_char = chr(0x2160)  # Roman numeral I
_h8_rpr = '<w:rPr><w:rFonts w:hint="eastAsia" w:eastAsia="\u5b8b\u4f53"/><w:sz w:val="21"/></w:rPr>'
_h8_old = '<w:r>' + _h8_rpr + '<w:t xml:space="preserve">\u7b2c  ' + roman_char + ' </w:t></w:r>'
_h8_page_field = (
    '<w:r>' + _h8_rpr + '<w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r>' + _h8_rpr + '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r>' + _h8_rpr + '<w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r>' + _h8_rpr + '<w:t xml:space="preserve">1</w:t></w:r>'
    '<w:r>' + _h8_rpr + '<w:fldChar w:fldCharType="end"/></w:r>'
)
_h8_new = '<w:r>' + _h8_rpr + '<w:t xml:space="preserve">\u7b2c  </w:t></w:r>' + _h8_page_field
header8_xml = header8_xml.replace(_h8_old, _h8_new)
# Fix spacing: template had \u7b2c  I  (2sp before I, 1sp after) with \u9875 in next run (1sp+\u9875)
# = 2 spaces between I and \u9875. After replacing I with PAGE field, we lose the trailing space
# from the I run, so add one extra leading space before \u9875 via regex (handles any attr variations).
import re
header8_xml = re.sub(
    r'(<w:t\s+xml:space="preserve">) (\u9875</w:t>)',
    r'\1  \2',
    header8_xml
)
# R33: Header border line fix (correct direction)
# ──────────────────────────────────────────────────────────────
# ORIGINAL TEMPLATE STRUCTURE (analyzed from header8.xml):
#   Para #0-#3: HEADER empty paragraphs, bottom=none (spacers)
#   Para #4:    TEXT BOX (wps:txbx) — bottom=single, sz=6 ← REAL header line
#   Para #5:    HEADER mirror paragraph — bottom=single ← DUPLICATE (creates double line)
#   Para #6:    HEADER whitespace — bottom=none
#
# The text box is a WPS floating shape that positions "北京航空航天大学" +
# page number in the header. Its internal paragraph's bottom=single IS the
# correct header line. WPS leaves a mirror paragraph OUTSIDE the text box
# (for non-WPS compatibility) that also has bottom=single — this creates
# the double line when both render.
#
# R31-R32 errors:
#   R31: regex with re.DOTALL crossed </wps:txbx> boundary, deleted header
#        paragraph's border instead of text box's (direction reversed)
#   R32: neutralized text box's pBdr (deleted the REAL header line), but
#        left mirror paragraph's border intact
#
# R33 FIXES:
#   A) DO NOT add bottom=single to Para #0 — the text box already has the line
#   B) KEEP the text box's bottom=single (original template design)
#   C) REMOVE the mirror paragraph's bottom=single (the duplicate)
# ──────────────────────────────────────────────────────────────

# A) SKIP Step 1 altogether — the text box border IS the header line.
#    Previous rounds added bottom=single to the first empty paragraph, which
#    created an extra line ABOVE the text box. That was the "top line that
#    should be deleted" mentioned in R31 feedback.

# B) The text box's bottom=single is PRESERVED as-is (no modification needed).

# C) Remove border from the mirror paragraph immediately after </wps:txbx>.
#    This paragraph has the SAME text content as the text box but its
#    bottom=single creates a duplicate horizontal line.
_txbx_end_match = re.search(r'</wps:txbx>', header8_xml)
if _txbx_end_match:
    _after_txbx = header8_xml[_txbx_end_match.end():]
    # Find the first <w:p> after the text box (the mirror paragraph)
    _mirror_para_match = re.search(r'<w:p[\s>]', _after_txbx)
    if _mirror_para_match:
        _mirror_start = _txbx_end_match.end() + _mirror_para_match.start()
        # Find the mirror paragraph's pBdr
        _mirror_pbdr_match = re.search(
            r'<w:pBdr>.*?</w:pBdr>',
            header8_xml[_mirror_start:],
            flags=re.DOTALL
        )
        if _mirror_pbdr_match and 'w:bottom w:val="single"' in _mirror_pbdr_match.group():
            _pbdr_abs_start = _mirror_start + _mirror_pbdr_match.start()
            _pbdr_abs_end = _mirror_start + _mirror_pbdr_match.end()
            _new_pbdr = ('<w:pBdr>'
                         '<w:top w:val="none" w:color="auto" w:sz="0" w:space="0"/>'
                         '<w:bottom w:val="none" w:color="auto" w:sz="0" w:space="0"/>'
                         '</w:pBdr>')
            header8_xml = (header8_xml[:_pbdr_abs_start]
                           + _new_pbdr
                           + header8_xml[_pbdr_abs_end:])
            print("  Header8 R33: neutralized mirror paragraph border (preserved text box line)")
        else:
            print("  Header8 R33: mirror paragraph border already none - no change needed")
    else:
        print("  Header8 R33: [WARN] no mirror paragraph found after txbx")
else:
    print("  Header8 R33: [WARN] could not find </wps:txbx>")

# R55: Save header8.xml as-is for front matter (PAGE field + mirror border already fixed above).
# Body pages use header14.xml (rId26) which has native 18.9pt VML + 240030 cy.
# header14 already has PAGE field, school name, NO Roman I — no modifications needed.
template_files['word/header8.xml'] = header8_xml.encode('utf-8')

# R59: Modify header20.xml (rId33) — template body header with VML 18.9pt (0.88cm in Word).
# Inject dynamic PAGE field, then use rId33 for ALL body sections.
if 'word/header20.xml' in template_files:
    h20_xml = template_files['word/header20.xml'].decode('utf-8')
    _rpr20 = '<w:rPr><w:rFonts w:hint="eastAsia" w:eastAsia="\u5b8b\u4f53"/><w:sz w:val="21"/></w:rPr>'
    # header20 has "第 48 " in one run: <w:r><w:rPr>...</w:rPr><w:t xml:space="preserve">第 48 </w:t></w:r>
    _old20 = '<w:r>' + _rpr20 + '<w:t xml:space="preserve">\u7b2c 48 </w:t></w:r>'
    if _old20 not in h20_xml:
        _old20 = '<w:r>' + _rpr20 + '<w:t>\u7b2c 48 </w:t></w:r>'
    # Replace with: 第  (text) + PAGE field + SPACE run + 页 (original run stays)
    _new20 = (
        '<w:r>' + _rpr20 + '<w:t xml:space="preserve">\u7b2c  </w:t></w:r>'
        '<w:r>' + _rpr20 + '<w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r>' + _rpr20 + '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r>' + _rpr20 + '<w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r>' + _rpr20 + '<w:t xml:space="preserve">1</w:t></w:r>'
        '<w:r>' + _rpr20 + '<w:fldChar w:fldCharType="end"/></w:r>'
    )
    h20_xml = h20_xml.replace(_old20, _new20)
    # R59: Add 2 spaces before "页" — template has natural spacing after "48 "
    # The space run in _new20 may get eaten by lxml. Fix "页" run directly.
    h20_xml = h20_xml.replace(
        '<w:t>' + chr(0x9875) + '</w:t>',
        '<w:t xml:space="preserve">  ' + chr(0x9875) + '</w:t>'
    )
    template_files['word/header20.xml'] = h20_xml.encode('utf-8')
    print("  Header20: injected dynamic PAGE field (rId33, VML 18.9pt)")
else:
    print("  Header20: [WARN] not found in template")

print(f"  Header8: saved for front matter, header20 (rId33) for body pages")

# Also need to add a final section break after 参考文献 so the last section
# (containing 结论, 致谢, 参考文献) inherits proper header/footer references.
# The template's body-level sectPr has NO header/footer references.
# R50: Add template-matching final section with ftr=964 to preserve crop mark
# alignment. Template has 16 sections, final one with footer=964 (appendix/refs).
# Our section injection was missing this, causing crop mark position mismatch.
body_parts.append(
    '<w:p><w:pPr><w:sectPr>'
    '<w:headerReference r:id="rId33" w:type="default"/>'
    '<w:footerReference r:id="rId17" w:type="default"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1701" w:right="1417" w:bottom="1417" w:left="1417" '
    'w:header="567" w:footer="964" w:gutter="0"/>'
    '<w:pgNumType w:fmt="decimal"/>'
    '<w:cols w:space="720" w:num="1"/>'
    '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"/>'
    '</w:sectPr></w:pPr></w:p>'
)
print("  Appended final section break with rId16/rId17 page references (continues body numbering)")

# Regenerate body_xml to include the final section break
body_xml = '\n'.join(body_parts)

template_doc = etree.fromstring(template_files['word/document.xml'])
template_body = template_doc.find(f'{{{NS_W}}}body')
template_children = list(template_body)

# ============================================================
# Front matter: fill in cover page, abstract, keywords from thesis content
# ============================================================
print("\n[3a] Filling front matter (cover, abstract, keywords)...")
meta = thesis['metadata']

def set_para_1st_t(para, new_text):
    """Replace text in first <w:t> of a paragraph, clear rest (preserves formatting)."""
    t_elements = list(para.iter(f'{{{NS_W}}}t'))
    if t_elements:
        t_elements[0].text = new_text
        for t in t_elements[1:]:
            t.text = ''
    else:
        # No t elements — add a new run with text
        r = etree.SubElement(para, f'{{{NS_W}}}r')
        t = etree.SubElement(r, f'{{{NS_W}}}t')
        t.text = new_text

def set_para_jc(para, val):
    """Set or change paragraph justification (jc)."""
    ppr = para.find(f'{{{NS_W}}}pPr')
    if ppr is None:
        ppr = etree.SubElement(para, f'{{{NS_W}}}pPr')
    jc = ppr.find(f'{{{NS_W}}}jc')
    if jc is None:
        jc = etree.SubElement(ppr, f'{{{NS_W}}}jc')
    jc.set(f'{{{NS_W}}}val', val)

def set_keywords_para(para, label, content):
    """Set keywords paragraph: bold label + normal content as separate runs."""
    runs = list(para.iter(f'{{{NS_W}}}r'))
    if runs:
        # First run -> bold label
        r0 = runs[0]
        r0_rpr = r0.find(f'{{{NS_W}}}rPr')
        if r0_rpr is None:
            r0_rpr = etree.SubElement(r0, f'{{{NS_W}}}rPr')
        if r0_rpr.find(f'{{{NS_W}}}b') is None:
            etree.SubElement(r0_rpr, f'{{{NS_W}}}b')
        if r0_rpr.find(f'{{{NS_W}}}bCs') is None:
            etree.SubElement(r0_rpr, f'{{{NS_W}}}bCs')
        t0 = r0.find(f'{{{NS_W}}}t')
        if t0 is None:
            t0 = etree.SubElement(r0, f'{{{NS_W}}}t')
        t0.text = label
        # Preserve leading/trailing spaces
        if label and (label[0] == ' ' or label[-1] == ' ' or content and content[0] == ' '):
            t0.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        # Second run -> normal content
        if len(runs) > 1:
            r1 = runs[1]
            r1_rpr = r1.find(f'{{{NS_W}}}rPr')
            if r1_rpr is not None:
                for tag in (f'{{{NS_W}}}b', f'{{{NS_W}}}bCs'):
                    el = r1_rpr.find(tag)
                    if el is not None:
                        r1_rpr.remove(el)
            t1 = r1.find(f'{{{NS_W}}}t')
            if t1 is None:
                t1 = etree.SubElement(r1, f'{{{NS_W}}}t')
            t1.text = content
            # Clear subsequent runs
            for r in runs[2:]:
                for t in r.iter(f'{{{NS_W}}}t'):
                    t.text = ''
        else:
            r1 = etree.SubElement(para, f'{{{NS_W}}}r')
            t1 = etree.SubElement(r1, f'{{{NS_W}}}t')
            t1.text = content
            if content and (content[0] == ' ' or content[-1] == ' '):
                t1.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    else:
        r0 = etree.SubElement(para, f'{{{NS_W}}}r')
        r0_rpr = etree.SubElement(r0, f'{{{NS_W}}}rPr')
        etree.SubElement(r0_rpr, f'{{{NS_W}}}b')
        etree.SubElement(r0_rpr, f'{{{NS_W}}}bCs')
        t0 = etree.SubElement(r0, f'{{{NS_W}}}t')
        t0.text = label
        r1 = etree.SubElement(para, f'{{{NS_W}}}r')
        t1 = etree.SubElement(r1, f'{{{NS_W}}}t')
        t1.text = content

# Build index-to-text map for key paragraphs
# Template structure (279 children total):
#   [1]  学 号38020326            → physical cover student ID
#   [16]  TABLE (cover info: school, major, student, advisor in cells)
#   [102] 作者：王小亮            → declaration page
#   [114] Chinese title           → abstract section
#   [115] Student CN
#   [116] Advisor CN
#   [117] "摘    要" (keep)
#   [118] Chinese abstract
#   [120] Chinese keywords
#   [122] English title
#   [123] Author EN
#   [124] Tutor EN
#   [125] "Abstract" (keep)
#   [126] English abstract part 1
#   [127] English abstract part 2
#   [128] English keywords
front_matter_fixes = {}

# [102] 作者：王小亮 → 作者：奚浩文 (declaration page)
front_matter_fixes[102] = f'作者：{meta["author_cn"]}'

# [114] Chinese title
front_matter_fixes[114] = meta['thesis_title_cn']

# [115] Author CN
front_matter_fixes[115] = f'学    生：{meta["author_cn"]}'

# [116] Advisor CN
front_matter_fixes[116] = f'指导教师：{meta["teacher_cn"]}'

# [118] Chinese abstract — collapse newlines to spaces
abstract_cn_flat = thesis['abstract_cn'].replace('\n', ' ').replace('\r', ' ')
abstract_cn_flat = re.sub(r'  +', ' ', abstract_cn_flat).strip()
front_matter_fixes[118] = abstract_cn_flat

# [120] Chinese keywords — handled separately (bold label + normal content)
cn_keywords_label = '关键词：'
cn_keywords_content = thesis['abstract_cn_keywords']

# [122] English title
front_matter_fixes[122] = meta['thesis_title_en']

# [123] Author EN
front_matter_fixes[123] = f'Author : {meta["author_en"]}'

# [124] Tutor EN  
front_matter_fixes[124] = f'Tutor : {meta["teacher_en"]}'

# [126] English abstract part 1 (first half of sentences)
abstract_en_text = thesis['abstract_en']
abstract_en_text = abstract_en_text.replace('\n', ' ').replace('\r', ' ')
abstract_en_text = re.sub(r'  +', ' ', abstract_en_text).strip()
# Split into ~2 halves by sentence count
sentences_en = re.split(r'(?<=[.!])\s+', abstract_en_text)
mid = max(1, len(sentences_en) // 2)
en_part1 = ' '.join(sentences_en[:mid])
en_part2 = ' '.join(sentences_en[mid:])
front_matter_fixes[126] = en_part1

# [127] English abstract part 2 (second half)
front_matter_fixes[127] = en_part2

# [128] English keywords — handled separately (bold label + normal content)
en_keywords_label = 'Key words: '
en_keywords_content = thesis['abstract_en_keywords']

# [1] Student ID on physical cover — special handling: replace only the numeric run,
# preserving the template's multi-run spacing structure (学 + spacing runs + 号 + spacing + ID + trailing spaces)
if 1 < len(template_children):
    child1 = template_children[1]
    if child1.tag == f'{{{NS_W}}}p':
        for r in child1.iter(f'{{{NS_W}}}r'):
            t = r.find(f'{{{NS_W}}}t')
            if t is not None and t.text and t.text.strip().isdigit():
                old_id = t.text
                t.text = meta['student_id']
                # Ensure the run has run properties (font/size) — fallback if template is missing rPr
                rpr = r.find(f'{{{NS_W}}}rPr')
                if rpr is None:
                    rpr = etree.SubElement(r, f'{{{NS_W}}}rPr')
                    rf = etree.SubElement(rpr, f'{{{NS_W}}}rFonts')
                    rf.set(f'{{{NS_W}}}eastAsia', '宋体')
                    rf.set(f'{{{NS_W}}}ascii', '宋体')
                    rf.set(f'{{{NS_W}}}hAnsi', '宋体')
                    _sz = etree.SubElement(rpr, f'{{{NS_W}}}sz')
                    _sz.set(f'{{{NS_W}}}val', '21')
                    _szCs = etree.SubElement(rpr, f'{{{NS_W}}}szCs')
                    _szCs.set(f'{{{NS_W}}}val', '21')
                print(f'  Front matter [1] Student ID: {old_id} -> {meta["student_id"]}')
                break

# R48: Replace template placeholder "TN953" with actual category value.
# Template cover ALREADY has "分类号" text (split across runs: 分/类/号)
# followed by "TN953" placeholder. We just replace the value, no insertion.
category_val = meta.get('category', '')
if category_val:
    found = False
    for child in template_children:
        for t in child.iter(f'{{{NS_W}}}t'):
            if t.text and 'TN953' in t.text:
                old = t.text
                t.text = t.text.replace('TN953', category_val)
                print(f'  R48: Replaced TN953 -> {category_val} in cover')
                found = True
                break
        if found:
            break
    if not found:
        print(f'  R48: [WARN] TN953 placeholder not found on cover')

# Apply all fixes (paragraph-based)
fixed_count = 0
for idx, new_text in front_matter_fixes.items():
    if idx < len(template_children):
        child = template_children[idx]
        if child.tag == f'{{{NS_W}}}p':
            set_para_1st_t(child, new_text)
            fixed_count += 1
            # Truncate for print
            display = new_text[:60] + ('...' if len(new_text) > 60 else '')
            print(f'  Front matter [{idx}]: {display}')
        else:
            print(f'  WARN: index {idx} is not a paragraph')
    else:
        print(f'  WARN: index {idx} beyond template_children ({len(template_children)})')
print(f'  {fixed_count} front matter paragraphs updated')

# Fix cover page table cells (template [16] is a table with cover info)
for i, child in enumerate(template_children):
    if child.tag == f'{{{NS_W}}}tbl':
        cells = child.findall(f'.//{{{NS_W}}}tc')
        cell_texts = []
        for c in cells:
            ts = [t.text or '' for t in c.iter(f'{{{NS_W}}}t')]
            cell_texts.append(''.join(ts).strip())
        # Detect cover info table by checking cell contents
        if len(cells) >= 8 and '学院名称' in cell_texts[0] and '学生姓名' in cell_texts[4]:
            print(f'  Found cover table at index {i}, updating cells...')
            # cell[1] = school, cell[3] = major, cell[5] = student, cell[7] = advisor
            cover_updates = {1: meta['school_cn'], 3: meta['major_cn'],
                             5: meta['author_cn'], 7: meta['teacher_cn']}
            for cell_idx, new_val in cover_updates.items():
                if cell_idx < len(cells):
                    ts = list(cells[cell_idx].iter(f'{{{NS_W}}}t'))
                    if ts:
                        ts[0].text = new_val
                        for t in ts[1:]:
                            t.text = ''
                        print(f'    Cover table cell [{cell_idx}]: {new_val}')
            break

# Handle keywords separately: bold label + normal content in distinct runs
for kw_idx, kw_label, kw_content in [
    (120, cn_keywords_label, cn_keywords_content),
    (128, en_keywords_label, en_keywords_content),
]:
    if kw_idx < len(template_children):
        child = template_children[kw_idx]
        if child.tag == f'{{{NS_W}}}p':
            set_keywords_para(child, kw_label, kw_content)
            display = f'{kw_label}{kw_content}'[:60]
            print(f'  Keywords [{kw_idx}]: {display}')

# R49: Fix paragraph justification for cover / abstract lines
# Template 学号 paragraph (idx 1) uses firstLine indent (6247) for positioning,
# NOT jc=right. Adding jc=right breaks the template's native alignment.
for jc_idx in [115, 116, 123, 124]:
    if jc_idx < len(template_children):
        child = template_children[jc_idx]
        if child.tag == f'{{{NS_W}}}p':
            set_para_jc(child, 'right')
            texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
            joined = ''.join(texts).strip()[:40]
            print(f'  Set jc=right for para [{jc_idx}]: {joined}')

body_doc = etree.fromstring(
    f'<root xmlns:w="{NS_W}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    f'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    f'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{body_xml}</root>'.encode('utf-8')
)
body_children = list(body_doc)

# Find body zone (reverse search for last 参考文献)
body_zone_start = None
body_zone_end = None
for i, child in enumerate(template_children):
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if body_zone_start is None and i > 145 and len(joined) > 2:
        body_zone_start = i
for i in range(len(template_children)-1, 0, -1):
    child = template_children[i]
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if '参考文献' in joined and i > body_zone_start + 20:
        body_zone_end = i
        break
if body_zone_start is None: body_zone_start = 170
if not body_zone_end or body_zone_end <= body_zone_start+10:
    body_zone_end = len(template_children) - 50

# ---- TOC injection: replace static TOC entries with dynamic TOC field ----
print(f"  Replace zone: {body_zone_start} to {body_zone_end}")
print("  Injecting dynamic TOC field...")
toc_title_idx = None
for i in range(100, body_zone_start):
    child = template_children[i]
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if '目' in joined and '录' in joined:
        toc_title_idx = i
        break

if toc_title_idx is not None:
    # Remove static TOC entries (after 目录 title, before body_zone_start)
    toc_remove_start = toc_title_idx + 1
    toc_remove_end = body_zone_start
    num_removed = toc_remove_end - toc_remove_start
    for i in range(toc_remove_end - 1, toc_remove_start - 1, -1):
        template_body.remove(template_children[i])
    
    # Build TOC field code paragraph
    toc_b = chr(92)
    toc_para = etree.Element(f'{{{NS_W}}}p')
    toc_ppr = etree.SubElement(toc_para, f'{{{NS_W}}}pPr')
    toc_jc = etree.SubElement(toc_ppr, f'{{{NS_W}}}jc')
    toc_jc.set(f'{{{NS_W}}}val', 'center')
    toc_ppr_rpr = etree.SubElement(toc_ppr, f'{{{NS_W}}}rPr')
    toc_sz = etree.SubElement(toc_ppr_rpr, f'{{{NS_W}}}sz')
    toc_sz.set(f'{{{NS_W}}}val', '24')
    
    r1 = etree.SubElement(toc_para, f'{{{NS_W}}}r')
    etree.SubElement(r1, f'{{{NS_W}}}fldChar', {f'{{{NS_W}}}fldCharType': 'begin'})
    r2 = etree.SubElement(toc_para, f'{{{NS_W}}}r')
    instr = etree.SubElement(r2, f'{{{NS_W}}}instrText', {f'{{{NS_W}}}space': 'preserve'})
    instr.text = f' TOC {toc_b}o "1-3" {toc_b}h {toc_b}z {toc_b}u '
    r3 = etree.SubElement(toc_para, f'{{{NS_W}}}r')
    etree.SubElement(r3, f'{{{NS_W}}}fldChar', {f'{{{NS_W}}}fldCharType': 'separate'})
    r4 = etree.SubElement(toc_para, f'{{{NS_W}}}r')
    r4rpr = etree.SubElement(r4, f'{{{NS_W}}}rPr')
    etree.SubElement(r4rpr, f'{{{NS_W}}}sz', {f'{{{NS_W}}}val': '24'})
    etree.SubElement(r4, f'{{{NS_W}}}t').text = '\u8bf7\u5728Word\u4e2d\u53f3\u952e\u66f4\u65b0\u57df'
    r5 = etree.SubElement(toc_para, f'{{{NS_W}}}r')
    etree.SubElement(r5, f'{{{NS_W}}}fldChar', {f'{{{NS_W}}}fldCharType': 'end'})
    
    # Insert blank paragraph after 目录 heading
    blank_para = etree.Element(f'{{{NS_W}}}p')
    blank_ppr = etree.SubElement(blank_para, f'{{{NS_W}}}pPr')
    blank_spacing = etree.SubElement(blank_ppr, f'{{{NS_W}}}spacing')
    blank_spacing.set(f'{{{NS_W}}}line', '360')
    blank_spacing.set(f'{{{NS_W}}}lineRule', 'auto')
    template_body.insert(toc_remove_start, blank_para)
    
    # Insert TOC field paragraph after the blank
    template_body.insert(toc_remove_start + 1, toc_para)
    
    # Adjust indices: removed num_removed elements, inserted 2
    body_zone_start = body_zone_start - num_removed + 2
    body_zone_end = body_zone_end - num_removed + 2
    print(f"  TOC injected with blank line, adjusted body_zone_start={body_zone_start}, body_zone_end={body_zone_end}")
else:
    print("  [WARN] TOC title not found, skipping TOC injection")

# Refresh template_children after TOC modification
template_children = list(template_body)

# ---- Fix page numbering: Roman numerals for front matter, Arabic from 1 for body ----
# 1. Add a section break to the paragraph right before body_zone_start (the TOC field
#    paragraph, injected above) with upperRoman pgNumType. This defines the front matter
#    section (abstract + TOC) to use Roman numerals.
toc_para = template_children[body_zone_start - 1]
if toc_para.tag == f'{{{NS_W}}}p':
    ppr = toc_para.find(f'{{{NS_W}}}pPr')
    if ppr is not None:
        # Check if this paragraph already has a sectPr (should not, but be safe)
        existing_sect = ppr.find(f'{{{NS_W}}}sectPr')
        if existing_sect is None:
            front_sect = etree.SubElement(ppr, f'{{{NS_W}}}sectPr')
        else:
            front_sect = existing_sect
        # Set page size (A4)
        pg_sz = front_sect.find(f'{{{NS_W}}}pgSz')
        if pg_sz is None:
            pg_sz = etree.SubElement(front_sect, f'{{{NS_W}}}pgSz')
        pg_sz.set(f'{{{NS_W}}}w', '11906')
        pg_sz.set(f'{{{NS_W}}}h', '16838')
        # Set page margins with explicit header/footer spacing
        pg_mar = front_sect.find(f'{{{NS_W}}}pgMar')
        if pg_mar is None:
            pg_mar = etree.SubElement(front_sect, f'{{{NS_W}}}pgMar')
        pg_mar.set(f'{{{NS_W}}}top', '1701')
        pg_mar.set(f'{{{NS_W}}}right', '1417')
        pg_mar.set(f'{{{NS_W}}}bottom', '1417')
        pg_mar.set(f'{{{NS_W}}}left', '1417')
        pg_mar.set(f'{{{NS_W}}}header', '567')
        pg_mar.set(f'{{{NS_W}}}footer', '851')
        pg_mar.set(f'{{{NS_W}}}gutter', '0')
        # Columns
        cols = front_sect.find(f'{{{NS_W}}}cols')
        if cols is None:
            cols = etree.SubElement(front_sect, f'{{{NS_W}}}cols')
        cols.set(f'{{{NS_W}}}space', '720')
        cols.set(f'{{{NS_W}}}num', '1')
        # Doc grid
        doc_grid = front_sect.find(f'{{{NS_W}}}docGrid')
        if doc_grid is None:
            doc_grid = etree.SubElement(front_sect, f'{{{NS_W}}}docGrid')
        doc_grid.set(f'{{{NS_W}}}type', 'linesAndChars')
        doc_grid.set(f'{{{NS_W}}}linePitch', '312')
        doc_grid.set(f'{{{NS_W}}}charSpace', '0')
        # Header/footer references (so TOC page uses the same header/footer as body)
        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        hdr_refs = front_sect.findall(f'{{{NS_W}}}headerReference')
        if not hdr_refs:
            hdr_ref = etree.SubElement(front_sect, f'{{{NS_W}}}headerReference')
            hdr_ref.set(f'{{{NS_W}}}type', 'default')
            hdr_ref.set(f'{{{ns_r}}}id', 'rId16')  # Front matter header
        ftr_refs = front_sect.findall(f'{{{NS_W}}}footerReference')
        if not ftr_refs:
            ftr_ref = etree.SubElement(front_sect, f'{{{NS_W}}}footerReference')
            ftr_ref.set(f'{{{NS_W}}}type', 'default')
            ftr_ref.set(f'{{{ns_r}}}id', 'rId17')
        # Set page number type (Roman numerals for front matter)
        pg_num_type = front_sect.find(f'{{{NS_W}}}pgNumType')
        if pg_num_type is None:
            pg_num_type = etree.SubElement(front_sect, f'{{{NS_W}}}pgNumType')
        pg_num_type.set(f'{{{NS_W}}}fmt', 'upperRoman')
        # Remove any existing start attribute (Roman starts from I by default)
        if f'{{{NS_W}}}start' in pg_num_type.attrib:
            del pg_num_type.attrib[f'{{{NS_W}}}start']
        print(f"  Front matter sectPr: page margins + header/footer refs + upperRoman set")

    # ---- Patch existing front-matter para-level sectPrs to use upperRoman ----
    # R34 bug: ALL front-matter sectPrs were blindly set to upperRoman (including
    # cover pages with template's w:start="9"). Cover pages ate Roman numerals
    # I-V, causing abstract to start at VI instead of I.
    # R35 fix: Skip cover sectPrs that already have a pgNumType with start attr.
    # Only convert sectPrs WITHOUT start (these govern the abstract/TOC area,
    # where numbering should begin fresh at I).
    patched_fm_count = 0
    patched_skip_count = 0
    for i in range(body_zone_start):
        child = template_children[i]
        if child.tag != f'{{{NS_W}}}p':
            continue
        ppr = child.find(f'{{{NS_W}}}pPr')
        if ppr is None:
            continue
        sp = ppr.find(f'{{{NS_W}}}sectPr')
        if sp is None:
            continue
        # Check if this sectPr already has a pgNumType with start attr (cover page)
        pg_num = sp.find(f'{{{NS_W}}}pgNumType')
        if pg_num is not None and f'{{{NS_W}}}start' in pg_num.attrib:
            patched_skip_count += 1
            continue  # Skip cover sections — keep template's original numbering
        if pg_num is None:
            pg_num = etree.SubElement(sp, f'{{{NS_W}}}pgNumType')
        pg_num.set(f'{{{NS_W}}}fmt', 'upperRoman')
        # Do NOT add or remove start attr — upperRoman defaults to start at I,
        # which is what we want for the abstract/TOC section.
        # Also ensure pgNumType has NO start attr (Roman starts from i by default)
        if f'{{{NS_W}}}start' in pg_num.attrib:
            del pg_num.attrib[f'{{{NS_W}}}start']
        # R49: Ensure footer reference rId17 so the PAGE field shows Roman numerals
        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        ftr_refs = sp.findall(f'{{{NS_W}}}footerReference')
        has_rId17 = any(fr.get(f'{{{ns_r}}}id') == 'rId17' for fr in ftr_refs)
        if not has_rId17:
            fr = etree.SubElement(sp, f'{{{NS_W}}}footerReference')
            fr.set(f'{{{NS_W}}}type', 'default')
            fr.set(f'{{{ns_r}}}id', 'rId17')
        patched_fm_count += 1
    print(f"  Patched {patched_fm_count} front-matter sectPrs to upperRoman (skipped {patched_skip_count} cover sectPrs)")

# 2. Body-start sectPr is merged into first body paragraph via _BODY_START_SECTPR
#    before the first chapter heading).  No additional sectPr is added to body_children.
print(f"  Body start sectPr merged into first body paragraph (decimal pgNumType, start=1, no blank page)")

# ---- Replace body content ----
for i in range(body_zone_end-1, body_zone_start-1, -1):
    template_body.remove(template_children[i])
insert_idx = body_zone_start
for child in body_children:
    template_body.insert(insert_idx, child)
    insert_idx += 1

# Fix section breaks: remove body-level sectPr, then ensure ALL paragraph-level sectPrs
# in the body zone have footerReference rId17 (for page numbers). Template appendix
# section breaks lack footer references, causing missing page numbers in later sections.
# 1. Remove body-level <w:sectPr> (overrides our section properties)
for child in list(template_body):
    if child.tag == f'{{{NS_W}}}sectPr':
        template_body.remove(child)

# 2. Ensure all paragraph-level sectPrs AFTER the first rId17 sectPr have footerReference.
#    This fixes template appendix section breaks that lack page number footers.
ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
seen_body_sect = False  # True once we've encountered a sectPr with footerRef rId17
for child in template_body:
    if child.tag == f'{{{NS_W}}}p':
        ppr = child.find(f'{{{NS_W}}}pPr')
        if ppr is not None:
            sect = ppr.find(f'{{{NS_W}}}sectPr')
            if sect is not None:
                frefs = sect.findall(f'{{{NS_W}}}footerReference')
                has_rId17 = any(fr.get(f'{{{ns_r}}}id') == 'rId17' for fr in frefs)
                if has_rId17:
                    seen_body_sect = True
                elif seen_body_sect and not frefs:
                    # Section break in body zone without footerReference — add one
                    fr = etree.SubElement(sect, f'{{{NS_W}}}footerReference')
                    fr.set(f'{{{NS_W}}}type', 'default')
                    fr.set(f'{{{ns_r}}}id', 'rId17')
print("  Section breaks patched: body-level sectPr removed, appendix footer refs added")

# ---- Remove "本人声明" paragraph (should not appear in TOC) ----
for child in list(template_body):
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if '本人声明' in joined:
        template_body.remove(child)
        print("  Removed '本人声明' paragraph from document")
        break

# ---- R75: Remove template placeholder appendix paragraphs ----
# Template has 附录A-C placeholders ("1/f频谱图", "C的取值表", "一维1/f波动数据")
# outside the body replacement zone. Remove them so our generated appendix is clean.
_TEMPLATE_APX_MARKERS = ['1/f频谱图', 'C的取值表', 'C取值表', '一维1/f波动数据']
_removed_apx = 0
for child in list(template_body):
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if any(m in joined for m in _TEMPLATE_APX_MARKERS):
        template_body.remove(child)
        _removed_apx += 1
if _removed_apx:
    print(f"  Removed {_removed_apx} template appendix placeholder paragraphs")


# ---- R75: Remove template placeholder appendix paragraphs ----
# Template has built-in 附录A-C placeholder paragraphs ("1/f频谱图", "C的取值表",
# "一维1/f波动数据") that are OUTSIDE the body replacement zone (paras ~307-331).
# Our generated appendix inserts 附录A-F with real thesis content.
# These template placeholders must be removed to avoid duplicate/conflicting labels.
template_apx_markers = ['1/f频谱图', 'C的取值表', 'C取值表', '一维1/f波动数据',
                         '1/f波动数据', '一维1/f噪声']
removed_apx = 0
for child in list(template_body):
    if child.tag != f'{{{NS_W}}}p': continue
    texts = [t.text or '' for t in child.iter(f'{{{NS_W}}}t')]
    joined = ''.join(texts).strip()
    if any(m in joined for m in template_apx_markers):
        template_body.remove(child)
        removed_apx += 1
print(f"  Removed {removed_apx} template appendix placeholder paragraphs")

# Image embedding
print("\n[4] Embedding images...")
new_doc_xml_str = etree.tostring(template_doc, encoding='unicode')

# R36 fix: Body page numbering must start at 1, not continue from upperRoman count.
# The body sectPr inherits the page count from preceding sections. After the abstract
# and TOC consume Roman numerals I-V, the body starts at 6 without explicit start=1.
# Fix: find the first <w:pgNumType w:fmt="decimal"/> without a start attr and add w:start="1".
_body_pgNum = re.search(r'<w:pgNumType w:fmt="decimal"[^>]*/>', new_doc_xml_str)
if _body_pgNum and 'w:start="1"' not in _body_pgNum.group():
    _new_pg = _body_pgNum.group().replace('/>', ' w:start="1"/>')
    new_doc_xml_str = (new_doc_xml_str[:_body_pgNum.start()]
                       + _new_pg
                       + new_doc_xml_str[_body_pgNum.end():])
    print("  R36: body pgNumType decimal start=1 injected")

img_placeholders = re.findall(r'rId_img_([^\s"\'>]+)', new_doc_xml_str)
unique_images = sorted(set(img_placeholders))
print(f"  Found {len(unique_images)} images")

img_files = {}
for fname in unique_images:
    for cand in [os.path.join(FIGURE_DIR, fname),
                 os.path.join(FIGURE_DIR, fname.replace('.png','_v1.png').replace('.jpg','_v1.jpg'))]:
        if os.path.isfile(cand):
            with open(cand, 'rb') as f:
                img_files[fname] = (f.read(), 'image/png' if cand.endswith('.png') else 'image/jpeg', cand)
            break

rels_xml = template_files.get('word/_rels/document.xml.rels', b'')
rels_root = etree.fromstring(rels_xml) if rels_xml else None
existing_rids = []
if rels_root is not None:
    for rel in rels_root:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try: existing_rids.append(int(rid[3:]))
            except: pass
max_rid = max(existing_rids) if existing_rids else 100
rId_map = {}
for i, (fname, (data, ct, path)) in enumerate(sorted(img_files.items())):
    new_rid_num = max_rid + i + 1
    rId = f'rId{new_rid_num}'
    rId_map[f'rId_img_{fname}'] = rId
    ext = path.rsplit('.',1)[-1].lower()
    template_files[f'word/media/image_{new_rid_num}.{ext}'] = data
    if rels_root is not None:
        rel = etree.SubElement(rels_root, 'Relationship')
        rel.set('Id', rId)
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        rel.set('Target', f'media/image_{new_rid_num}.{ext}')

if rels_root is not None:
    template_files['word/_rels/document.xml.rels'] = etree.tostring(rels_root, encoding='UTF-8', xml_declaration=True)

for placeholder, actual in rId_map.items():
    new_doc_xml_str = new_doc_xml_str.replace(f'rId_img_{placeholder.replace("rId_img_","")}', actual)
    new_doc_xml_str = new_doc_xml_str.replace(placeholder, actual)

# Fix Content_Types for new image extensions
ct_xml = template_files.get('[Content_Types].xml', b'')
if ct_xml:
    ct_root = etree.fromstring(ct_xml)
    existing_exts = set()
    for child in ct_root:
        ext = child.get('Extension', '')
        if ext: existing_exts.add(ext)
    for fname in img_files:
        ext = fname.rsplit('.', 1)[-1].lower()
        if ext not in existing_exts:
            existing_exts.add(ext)
            ct = 'image/png' if ext == 'png' else 'image/jpeg'
            ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
            default = etree.SubElement(ct_root, f'{{{ns_ct}}}Default')
            default.set('Extension', ext)
            default.set('ContentType', ct)
            print(f'    Added Content_Type: {ext}')
    template_files['[Content_Types].xml'] = etree.tostring(ct_root, encoding='UTF-8', xml_declaration=True)

# Numbering
print("\n[5] Injecting numbering...")
num_xml = template_files.get('word/numbering.xml', b'')
if num_xml:
    num_root = etree.fromstring(num_xml)
    num_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    abs_num = etree.SubElement(num_root, f'{{{num_ns}}}abstractNum')
    abs_num.set(f'{{{num_ns}}}abstractNumId', '100')
    for idx, (ilvl, fmt, txt) in enumerate([(0,'decimal','%1'),(1,'decimal','%1.%2'),(2,'decimal','%1.%2.%3'),(3,'decimal','%1.%2.%3.%4')]):
        lvl = etree.SubElement(abs_num, f'{{{num_ns}}}lvl')
        lvl.set(f'{{{num_ns}}}ilvl', str(ilvl))
        for tag, val in [('start','1'),('numFmt',fmt),('lvlText',txt),('lvlJc','left')]:
            e = etree.SubElement(lvl, f'{{{num_ns}}}{tag}')
            e.set(f'{{{num_ns}}}val', val)
        ppr = etree.SubElement(lvl, f'{{{num_ns}}}pPr')
        ind = etree.SubElement(ppr, f'{{{num_ns}}}ind')
        ind.set(f'{{{num_ns}}}left', '480')
        ind.set(f'{{{num_ns}}}hanging', '420')
    num = etree.SubElement(num_root, f'{{{num_ns}}}num')
    num.set(f'{{{num_ns}}}numId', '100')
    ailvl = etree.SubElement(num, f'{{{num_ns}}}abstractNumId')
    ailvl.set(f'{{{num_ns}}}val', '100')
    template_files['word/numbering.xml'] = etree.tostring(num_root, encoding='UTF-8', xml_declaration=True)

# updateFields
settings_xml = template_files.get('word/settings.xml', b'')
if settings_xml:
    set_root = etree.fromstring(settings_xml)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    if not set_root.findall(f'{{{ns}}}updateFields'):
        uf = etree.SubElement(set_root, f'{{{ns}}}updateFields')
        uf.set(f'{{{ns}}}val', 'true')
        template_files['word/settings.xml'] = etree.tostring(set_root, encoding='UTF-8', xml_declaration=True)

# ---- Inject TOC1/TOC2/TOC3 styles into styles.xml ----
print("\n[6] Injecting TOC styles into styles.xml...")
styles_xml = template_files.get('word/styles.xml', b'')
if styles_xml:
    sty_root = etree.fromstring(styles_xml)
    sty_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # Helper to build a TOC style element
    def make_toc_style(style_id, toc_name, based_on, eastasia_font, ascii_font,
                       bold, sz_val, left_indent, tab_pos):
        style_el = etree.Element(f'{{{sty_ns}}}style')
        style_el.set(f'{{{sty_ns}}}type', 'paragraph')
        style_el.set(f'{{{sty_ns}}}styleId', style_id)
        # name
        name_el = etree.SubElement(style_el, f'{{{sty_ns}}}name')
        name_el.set(f'{{{sty_ns}}}val', toc_name)
        # basedOn
        bo = etree.SubElement(style_el, f'{{{sty_ns}}}basedOn')
        bo.set(f'{{{sty_ns}}}val', based_on)
        # next
        nx = etree.SubElement(style_el, f'{{{sty_ns}}}next')
        nx.set(f'{{{sty_ns}}}val', based_on)
        # uiPriority
        uip = etree.SubElement(style_el, f'{{{sty_ns}}}uiPriority')
        uip.set(f'{{{sty_ns}}}val', '39')
        # pPr
        ppr = etree.SubElement(style_el, f'{{{sty_ns}}}pPr')
        jc = etree.SubElement(ppr, f'{{{sty_ns}}}jc')
        jc.set(f'{{{sty_ns}}}val', 'left')
        spacing = etree.SubElement(ppr, f'{{{sty_ns}}}spacing')
        spacing.set(f'{{{sty_ns}}}line', '360')
        spacing.set(f'{{{sty_ns}}}lineRule', 'auto')
        if left_indent > 0:
            ind = etree.SubElement(ppr, f'{{{sty_ns}}}ind')
            ind.set(f'{{{sty_ns}}}left', str(left_indent))
        tabs = etree.SubElement(ppr, f'{{{sty_ns}}}tabs')
        tab = etree.SubElement(tabs, f'{{{sty_ns}}}tab')
        tab.set(f'{{{sty_ns}}}val', 'right')
        tab.set(f'{{{sty_ns}}}leader', 'dot')
        tab.set(f'{{{sty_ns}}}pos', str(tab_pos))
        # rPr
        rpr = etree.SubElement(style_el, f'{{{sty_ns}}}rPr')
        rfonts = etree.SubElement(rpr, f'{{{sty_ns}}}rFonts')
        rfonts.set(f'{{{sty_ns}}}eastAsia', eastasia_font)
        rfonts.set(f'{{{sty_ns}}}ascii', ascii_font)
        rfonts.set(f'{{{sty_ns}}}hAnsi', ascii_font)
        # Explicit unbold for all TOC entries (TOC1/TOC2/TOC3)
        b_el = etree.SubElement(rpr, f'{{{sty_ns}}}b')
        b_el.set(f'{{{sty_ns}}}val', '0')
        bCs_el = etree.SubElement(rpr, f'{{{sty_ns}}}bCs')
        bCs_el.set(f'{{{sty_ns}}}val', '0')
        sz_el = etree.SubElement(rpr, f'{{{sty_ns}}}sz')
        sz_el.set(f'{{{sty_ns}}}val', str(sz_val))
        szCs_el = etree.SubElement(rpr, f'{{{sty_ns}}}szCs')
        szCs_el.set(f'{{{sty_ns}}}val', str(sz_val))
        return style_el
    
    PG_W = 11906  # page width twips
    LM = 1417     # left margin twips
    RM = 1417     # right margin twips
    TEXT_W = PG_W - LM - RM  # 9072 twips
    
    toc_styles = [
        # TOC1: 黑体 bold 小四(sz=24), left-aligned, no indent, tab at right margin
        ('TOC1', 'toc 1', '1', '黑体', 'Times New Roman', True, 24, 0, TEXT_W),
        # TOC2: 宋体 小四(sz=24), left-indent 480 (~2 chars), tab at right margin
        ('TOC2', 'toc 2', '1', '宋体', 'Times New Roman', False, 24, 480, TEXT_W),
        # TOC3: 宋体 小四(sz=24), left-indent 960 (~4 chars), tab at right margin
        ('TOC3', 'toc 3', '1', '宋体', 'Times New Roman', False, 24, 960, TEXT_W),
    ]
    for args in toc_styles:
        sty_root.append(make_toc_style(*args))
    
    # ---- Inject BUAA Thesis Table table style with conditional formatting ----
    def inject_table_style(root):
        """Create and append BUAA Thesis Table style element.
        
        Custom table style with tblStylePr conditional formatting:
        - firstRow: 黑体 bold, sz=21, centered (header row)
        - wholeTable: 宋体, sz=21, centered (data cells)
        - lastRow/firstCol/lastCol: empty (no special formatting)
        - Cell margins: top=0, left=108, bottom=0, right=108
        """
        s = etree.SubElement(root, f'{{{sty_ns}}}style')
        s.set(f'{{{sty_ns}}}type', 'table')
        s.set(f'{{{sty_ns}}}styleId', 'BUAA_Thesis_Table')
        
        n = etree.SubElement(s, f'{{{sty_ns}}}name')
        n.set(f'{{{sty_ns}}}val', 'BUAA Thesis Table')
        
        bo = etree.SubElement(s, f'{{{sty_ns}}}basedOn')
        bo.set(f'{{{sty_ns}}}val', '31')
        
        # Base rPr: 宋体 Times New Roman sz=21
        rpr = etree.SubElement(s, f'{{{sty_ns}}}rPr')
        rf = etree.SubElement(rpr, f'{{{sty_ns}}}rFonts')
        rf.set(f'{{{sty_ns}}}eastAsia', '宋体')
        rf.set(f'{{{sty_ns}}}ascii', 'Times New Roman')
        rf.set(f'{{{sty_ns}}}hAnsi', 'Times New Roman')
        sz = etree.SubElement(rpr, f'{{{sty_ns}}}sz'); sz.set(f'{{{sty_ns}}}val', '21')
        szCs = etree.SubElement(rpr, f'{{{sty_ns}}}szCs'); szCs.set(f'{{{sty_ns}}}val', '21')
        
        # tblPr: cell margins
        tp = etree.SubElement(s, f'{{{sty_ns}}}tblPr')
        tcm = etree.SubElement(tp, f'{{{sty_ns}}}tblCellMar')
        for side, val in [('top', 0), ('left', 108), ('bottom', 0), ('right', 108)]:
            m = etree.SubElement(tcm, f'{{{sty_ns}}}{side}')
            m.set(f'{{{sty_ns}}}w', str(val))
            m.set(f'{{{sty_ns}}}type', 'dxa')
        
        # tblStylePr firstRow: 黑体, centered, sz=21 (no bold per user req)
        fr = etree.SubElement(s, f'{{{sty_ns}}}tblStylePr')
        fr.set(f'{{{sty_ns}}}type', 'firstRow')
        fr_ppr = etree.SubElement(fr, f'{{{sty_ns}}}pPr')
        fr_jc = etree.SubElement(fr_ppr, f'{{{sty_ns}}}jc')
        fr_jc.set(f'{{{sty_ns}}}val', 'center')
        fr_rpr = etree.SubElement(fr, f'{{{sty_ns}}}rPr')
        fr_rf = etree.SubElement(fr_rpr, f'{{{sty_ns}}}rFonts')
        fr_rf.set(f'{{{sty_ns}}}eastAsia', '黑体')
        fr_rf.set(f'{{{sty_ns}}}ascii', '黑体')
        fr_rf.set(f'{{{sty_ns}}}hAnsi', '黑体')
        etree.SubElement(fr_rpr, f'{{{sty_ns}}}b', {f'{{{sty_ns}}}val': '0'})
        etree.SubElement(fr_rpr, f'{{{sty_ns}}}bCs', {f'{{{sty_ns}}}val': '0'})
        fr_sz = etree.SubElement(fr_rpr, f'{{{sty_ns}}}sz'); fr_sz.set(f'{{{sty_ns}}}val', '21')
        fr_szCs = etree.SubElement(fr_rpr, f'{{{sty_ns}}}szCs'); fr_szCs.set(f'{{{sty_ns}}}val', '21')
        # Vertical centering for header cells
        fr_tcpr = etree.SubElement(fr, f'{{{sty_ns}}}tcPr')
        fr_valign = etree.SubElement(fr_tcpr, f'{{{sty_ns}}}vAlign')
        fr_valign.set(f'{{{sty_ns}}}val', 'center')
        
        # tblStylePr wholeTable: 宋体, centered, sz=21
        wt = etree.SubElement(s, f'{{{sty_ns}}}tblStylePr')
        wt.set(f'{{{sty_ns}}}type', 'wholeTable')
        wt_ppr = etree.SubElement(wt, f'{{{sty_ns}}}pPr')
        wt_jc = etree.SubElement(wt_ppr, f'{{{sty_ns}}}jc')
        wt_jc.set(f'{{{sty_ns}}}val', 'center')
        wt_rpr = etree.SubElement(wt, f'{{{sty_ns}}}rPr')
        wt_rf = etree.SubElement(wt_rpr, f'{{{sty_ns}}}rFonts')
        wt_rf.set(f'{{{sty_ns}}}eastAsia', '宋体')
        wt_rf.set(f'{{{sty_ns}}}ascii', '宋体')
        wt_rf.set(f'{{{sty_ns}}}hAnsi', '宋体')
        wt_sz = etree.SubElement(wt_rpr, f'{{{sty_ns}}}sz'); wt_sz.set(f'{{{sty_ns}}}val', '21')
        wt_szCs = etree.SubElement(wt_rpr, f'{{{sty_ns}}}szCs'); wt_szCs.set(f'{{{sty_ns}}}val', '21')
        
        # Empty tblStylePr entries (no special formatting)
        for empty_type in ('lastRow', 'firstCol', 'lastCol'):
            e = etree.SubElement(s, f'{{{sty_ns}}}tblStylePr')
            e.set(f'{{{sty_ns}}}type', empty_type)
    
    inject_table_style(sty_root)
    print("  BUAA Thesis Table style injected (firstRow=黑体, dataRow=宋体)")
    
    # ---- Inject 黑体 into heading style definitions (4, 6) so headings inherit font from style,
    #      not from direct run formatting. This prevents Word from copying heading font to TOC entries.
    for sid in ('4', '6'):
        for style in sty_root:
            s = style.get(f'{{{sty_ns}}}styleId')
            if s == sid:
                rpr = None
                for child in style:
                    if child.tag == f'{{{sty_ns}}}rPr':
                        rpr = child
                        break
                if rpr is not None:
                    rf = None
                    for child in rpr:
                        if child.tag == f'{{{sty_ns}}}rFonts':
                            rf = child
                            break
                    if rf is None:
                        rf = etree.SubElement(rpr, f'{{{sty_ns}}}rFonts')
                    rf.set(f'{{{sty_ns}}}eastAsia', '黑体')
                    print(f"  Heading style {sid}: eastAsia=黑体")
                break
    
    # ---- Fix body text styles 17/18: ensure correct defaults ----
    # Template style 17 (Body Text) and 18 (Body Text Indent) lack line spacing
    # and east-asia font settings. Fix them so inherited formatting is correct.
    for sid, fix_firstline in [('17', None), ('18', '480')]:
        for style in sty_root:
            if style.tag != f'{{{sty_ns}}}style':
                continue
            if style.get(f'{{{sty_ns}}}styleId') != sid:
                continue
            # rPr: sz=24, eastAsia=宋体
            rpr = style.find(f'{{{sty_ns}}}rPr')
            if rpr is None:
                rpr = etree.SubElement(style, f'{{{sty_ns}}}rPr')
            rf = rpr.find(f'{{{sty_ns}}}rFonts')
            if rf is None:
                rf = etree.SubElement(rpr, f'{{{sty_ns}}}rFonts')
            rf.set(f'{{{sty_ns}}}eastAsia', '宋体')
            rf.set(f'{{{sty_ns}}}ascii', 'Times New Roman')
            rf.set(f'{{{sty_ns}}}hAnsi', 'Times New Roman')
            sz_el = rpr.find(f'{{{sty_ns}}}sz')
            if sz_el is None:
                sz_el = etree.SubElement(rpr, f'{{{sty_ns}}}sz')
            sz_el.set(f'{{{sty_ns}}}val', '24')
            szcs_el = rpr.find(f'{{{sty_ns}}}szCs')
            if szcs_el is None:
                szcs_el = etree.SubElement(rpr, f'{{{sty_ns}}}szCs')
            szcs_el.set(f'{{{sty_ns}}}val', '24')
            # pPr: line spacing 360 (1.5x)
            ppr = style.find(f'{{{sty_ns}}}pPr')
            if ppr is None:
                ppr = etree.SubElement(style, f'{{{sty_ns}}}pPr')
            spacing = ppr.find(f'{{{sty_ns}}}spacing')
            if spacing is None:
                spacing = etree.SubElement(ppr, f'{{{sty_ns}}}spacing')
            spacing.set(f'{{{sty_ns}}}line', '360')
            spacing.set(f'{{{sty_ns}}}lineRule', 'auto')
            # firstLine indent for style 18
            if fix_firstline:
                ind = ppr.find(f'{{{sty_ns}}}ind')
                if ind is None:
                    ind = etree.SubElement(ppr, f'{{{sty_ns}}}ind')
                ind.set(f'{{{sty_ns}}}firstLine', fix_firstline)
            print(f'  Style {sid} ({sid}: {"Body Text Indent" if fix_firstline else "Body Text"}): patched sz=24, eastAsia=宋体, line=360, firstLine={fix_firstline or "none"}')
    
    template_files['word/styles.xml'] = etree.tostring(sty_root, encoding='UTF-8', xml_declaration=True)
    print("  TOC1/TOC2/TOC3 style definitions injected.")

# Write
print("\n[7] Writing DOCX...")
template_files['word/document.xml'] = new_doc_xml_str.encode('utf-8')
with zipfile.ZipFile(OUTPUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in sorted(template_files.items()):
        zout.writestr(name, data)

print(f"Output: {OUTPUT_PATH} ({os.path.getsize(OUTPUT_PATH)/1024:.0f} KB)")

with zipfile.ZipFile(OUTPUT_PATH, 'r') as zf:
    try:
        etree.fromstring(zf.read('word/document.xml'))
        print("document.xml: VALID")
    except Exception as e: print(f"document.xml: {e}")
    try:
        etree.fromstring(zf.read('word/_rels/document.xml.rels'))
        print("rels: VALID")
    except Exception as e: print(f"rels: {e}")
print("Done!")
