# ⚠️ 新团队接手前必须先读: ../HANDOFF_遗嘱_必读.md
"""
OOXML (Office Open XML) fragment generator for building Word document content
in Flat OPC format. Produces valid XML strings for paragraphs, tables, images,
and other Word document elements.

Used to convert LaTeX thesis content into Word XML paragraphs.
"""

import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
import re
import os

# OOXML namespaces
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

# Citation pattern for superscript detection: [1], [2,3], [4-7], [3,6-9]
_CITATION_RE = re.compile(r'(\[[0-9]+(?:[-,][0-9]+)*\])')

# Register namespace prefixes so ElementTree serializes as w:, r:, etc.
ET.register_namespace('w', NS_W)
ET.register_namespace('r', NS_R)
ET.register_namespace('wp', NS_WP)
ET.register_namespace('a', NS_A)
ET.register_namespace('pic', NS_PIC)
ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')

# Shortcut for w:val attribute (most common OOXML attribute)
_W_VAL = f'{{{NS_W}}}val'
_W_W = f'{{{NS_W}}}w'
_W_TYPE = f'{{{NS_W}}}type'

# Registry of namespace prefixes for ElementTree serialization
_NS_MAP = {
    'w': NS_W,
    'r': NS_R,
    'wp': NS_WP,
    'a': NS_A,
    'pic': NS_PIC,
}

# Alignment value mapping
_ALIGN_MAP = {
    'center': 'center',
    'left': 'left',
    'right': 'right',
    'both': 'both',
    'justify': 'both',
}


def _make_element(tag, attrib=None, ns=NS_W):
    """Create an Element with the given namespace prefix."""
    if attrib is None:
        attrib = {}
    # Use Clark notation: {namespace}localname
    full_tag = f'{{{ns}}}{tag}'
    return ET.Element(full_tag, attrib)


def _make_sub_element(parent, tag, attrib=None, ns=NS_W, text=None):
    """Create and append a sub-element."""
    elem = _make_element(tag, attrib, ns)
    parent.append(elem)
    if text is not None:
        elem.text = escape_xml(text) if ns == NS_W else text
    return elem


def _serialize(element):
    """Serialize an ElementTree element to a string without XML declaration."""
    return ET.tostring(element, encoding='unicode', short_empty_elements=True)


def escape_xml(text):
    """Escape XML special characters: & < > \" """
    return escape(text, {'"': '&quot;'})


def make_run(text, bold=False, italic=False, subscript=False, superscript=False):
    """Generate a single <w:r> element as a string.

    Args:
        text: Run text content.
        bold: Apply bold formatting.
        italic: Apply italic formatting.
        subscript: Apply subscript formatting (vertAlign subscript).
        superscript: Apply superscript formatting (vertAlign superscript).

    Returns:
        XML string for a <w:r> element.
    """
    r = _make_element('r')
    has_rpr = bold or italic or subscript or superscript
    if has_rpr:
        rpr = _make_sub_element(r, 'rPr')
        if bold:
            _make_sub_element(rpr, 'b')
        if italic:
            _make_sub_element(rpr, 'i')
        if subscript:
            _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'subscript'})
        if superscript:
            _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
    _make_sub_element(r, 't', text=text)
    return _serialize(r)


def make_paragraph(style_id, text, bold=False, italic=False,
                   font_size=None, alignment=None,
                   detect_citations=True,
                   num_id=None, ilvl=0):
    """Generate a <w:p> element as a string.

    Automatically detects citation patterns like [1], [2,3], [4-7]
    and renders them as superscript with smaller font (9pt).

    Args:
        style_id: Style ID string (e.g. '1', 'a9'), or None to omit pStyle.
        text: Paragraph text content.
        bold: Apply bold to the entire paragraph text.
        italic: Apply italic to the entire paragraph text.
        font_size: Font size in half-points (e.g. 44 for 22pt). None = omit.
        alignment: 'center', 'left', 'right', 'both' (justified).
        num_id: Template numbering ID for auto-numbered lists (e.g. 1 for [N] refs).
        ilvl: Indentation level for the numbering (default 0).

    Returns:
        XML string for a <w:p> element.
    """
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')

    if style_id is not None:
        _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: style_id},
                          ns=NS_W)

    if num_id is not None:
        numPr = _make_sub_element(ppr, 'numPr', ns=NS_W)
        _make_sub_element(numPr, 'ilvl', attrib={_W_VAL: str(ilvl)},
                          ns=NS_W)
        _make_sub_element(numPr, 'numId', attrib={_W_VAL: str(num_id)},
                          ns=NS_W)

    if alignment is not None:
        align_val = _ALIGN_MAP.get(alignment, alignment)
        _make_sub_element(ppr, 'jc', attrib={_W_VAL: align_val},
                          ns=NS_W)

    # Detect citations and split text into segments
    parts = _CITATION_RE.split(text) if detect_citations else [text]
    has_multiple_parts = detect_citations and len(parts) > 1

    for part in parts:
        if not part:
            continue
        is_citation = has_multiple_parts and _CITATION_RE.fullmatch(part) is not None

        r = _make_sub_element(p, 'r')
        rpr_needed = bold or italic or font_size is not None or is_citation
        if rpr_needed:
            rpr = _make_sub_element(r, 'rPr')
            if bold:
                _make_sub_element(rpr, 'b')
            if italic:
                _make_sub_element(rpr, 'i')
            if font_size is not None and not is_citation:
                # Paragraph-level font size (not applied to superscript citations)
                _make_sub_element(rpr, 'sz', attrib={_W_VAL: str(font_size)},
                                  ns=NS_W)
                _make_sub_element(rpr, 'szCs', attrib={_W_VAL: str(font_size)},
                                  ns=NS_W)
            if is_citation:
                _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
                _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'},
                                  ns=NS_W)
                _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'},
                                  ns=NS_W)
        _make_sub_element(r, 't', text=part)

    return _serialize(p)


def make_heading(level, title, back_matter=False):
    """Generate a heading paragraph with proper run-level formatting.

    Based on template analysis:
    - level=1 body: unbold, 32pt, HeiTi, centered, line=480, before=500, after=500
    - level=1 back: unbold, 32pt, HeiTi, centered, line=360, before=240, after=240
    - level=2 (section): unbold, 24pt, justified, HeiTi, line=360
    - level=3 (subsection): NOT bold, 24pt, left, HeiTi
    - level=4 (subsubsection): NOT bold, 24pt, left, HeiTi

    Args:
        level: 1 (chapter), 2 (section), 3 (subsection), 4 (subsubsection).
        title: Heading text (should already include manual numbering, e.g. "1 绪论").
        back_matter: True for 结论/致谢/参考文献 (tighter spacing than body chapters).

    Returns:
        XML string for a heading <w:p> element.
    """
    # Common fonts
    HEITI = '黑体'

    if level == 1:
        # Chapter heading (H1): 三号(16pt, sz=32), 黑体, UNBOLD, centered
        # Body chapters: line=480, before=500, after=500
        # Back matter: line=360, before=240, after=240 (tighter, matches template)
        p = _make_element('p')
        ppr = _make_sub_element(p, 'pPr')
        _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '2'}, ns=NS_W)
        # Auto-numbering: numId=100, ilvl=0 (H1) — body only
        if not back_matter:
            numPr = _make_sub_element(ppr, 'numPr', ns=NS_W)
            _make_sub_element(numPr, 'ilvl', attrib={_W_VAL: '0'}, ns=NS_W)
            _make_sub_element(numPr, 'numId', attrib={_W_VAL: '100'}, ns=NS_W)
        # Explicit spacing
        spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
        if back_matter:
            spacing.set(f'{{{NS_W}}}line', '480')
            spacing.set(f'{{{NS_W}}}lineRule', 'auto')
            spacing.set(f'{{{NS_W}}}before', '500')
            spacing.set(f'{{{NS_W}}}after', '500')
        else:
            spacing.set(f'{{{NS_W}}}line', '480')
            spacing.set(f'{{{NS_W}}}lineRule', 'auto')
            spacing.set(f'{{{NS_W}}}before', '500')
            spacing.set(f'{{{NS_W}}}after', '500')
        # Centered alignment
        _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'center'}, ns=NS_W)
        # Run-level override in pPr: 32pt (16pt, 三号), NO bold (override style 2 inheritance)
        ppr_rpr = _make_sub_element(ppr, 'rPr')
        _make_sub_element(ppr_rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '32'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '32'}, ns=NS_W)
        # Run with formatting
        r = _make_sub_element(p, 'r')
        rpr = _make_sub_element(r, 'rPr')
        rf = _make_sub_element(rpr, 'rFonts', ns=NS_W)
        rf.set(f'{{{NS_W}}}eastAsia', HEITI)
        _make_sub_element(rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(rpr, 'sz', attrib={_W_VAL: '32'}, ns=NS_W)
        _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '32'}, ns=NS_W)
        _make_sub_element(r, 't', text=title)
        return _serialize(p)

    elif level == 2:
        # H2: 黑体小四(12pt, sz=24), UNBOLD, left, line=420, before/after=160
        p = _make_element('p')
        ppr = _make_sub_element(p, 'pPr')
        _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '4'}, ns=NS_W)
        # Auto-numbering: numId=100, ilvl=1 (H2)
        numPr = _make_sub_element(ppr, 'numPr', ns=NS_W)
        _make_sub_element(numPr, 'ilvl', attrib={_W_VAL: '1'}, ns=NS_W)
        _make_sub_element(numPr, 'numId', attrib={_W_VAL: '100'}, ns=NS_W)
        # Line spacing: 420 (1.75x), before=160, after=160
        spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
        spacing.set(f'{{{NS_W}}}line', '420')
        spacing.set(f'{{{NS_W}}}lineRule', 'auto')
        spacing.set(f'{{{NS_W}}}before', '160')
        spacing.set(f'{{{NS_W}}}after', '160')
        # Left-aligned (aligns with body text margin)
        _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'left'}, ns=NS_W)
        # pPr/rPr override: 24pt (12pt, 小四), 黑体, NO bold (font moved from run to pPr/rPr so TOC inherits style font)
        ppr_rpr = _make_sub_element(ppr, 'rPr')
        rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
        rf_ppr.set(f'{{{NS_W}}}eastAsia', HEITI)
        _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 4
        _make_sub_element(ppr_rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        r = _make_sub_element(p, 'r')
        rpr = _make_sub_element(r, 'rPr')
        _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 4
        _make_sub_element(rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(r, 't', text=title)
        return _serialize(p)

    elif level == 3:
        # H3: 黑体小四(12pt, sz=24), UNBOLD, left, before/after=160 (template Style 6)
        p = _make_element('p')
        ppr = _make_sub_element(p, 'pPr')
        _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '6'}, ns=NS_W)
        # Auto-numbering: numId=100, ilvl=2 (H3)
        numPr = _make_sub_element(ppr, 'numPr', ns=NS_W)
        _make_sub_element(numPr, 'ilvl', attrib={_W_VAL: '2'}, ns=NS_W)
        _make_sub_element(numPr, 'numId', attrib={_W_VAL: '100'}, ns=NS_W)
        spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
        spacing.set(f'{{{NS_W}}}line', '360')
        spacing.set(f'{{{NS_W}}}lineRule', 'auto')
        spacing.set(f'{{{NS_W}}}before', '160')
        spacing.set(f'{{{NS_W}}}after', '160')
        _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'left'}, ns=NS_W)
        # pPr/rPr override: 24pt (12pt, 小四), 黑体, NOT bold (font moved from run to pPr/rPr so TOC inherits style font)
        ppr_rpr = _make_sub_element(ppr, 'rPr')
        rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
        rf_ppr.set(f'{{{NS_W}}}eastAsia', HEITI)
        _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 6
        _make_sub_element(ppr_rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        r = _make_sub_element(p, 'r')
        rpr = _make_sub_element(r, 'rPr')
        _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 6
        _make_sub_element(rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(r, 't', text=title)
        return _serialize(p)

    elif level == 4:
        # H4 (subsubsection): 黑体小四(12pt, sz=24), UNBOLD, left
        p = _make_element('p')
        ppr = _make_sub_element(p, 'pPr')
        _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '6'}, ns=NS_W)
        # Auto-numbering: numId=100, ilvl=3 (H4)
        numPr = _make_sub_element(ppr, 'numPr', ns=NS_W)
        _make_sub_element(numPr, 'ilvl', attrib={_W_VAL: '3'}, ns=NS_W)
        _make_sub_element(numPr, 'numId', attrib={_W_VAL: '100'}, ns=NS_W)
        _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'left'}, ns=NS_W)
        # Explicit spacing: line=360 (1.5x), before=120, after=120
        spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
        spacing.set(f'{{{NS_W}}}line', '360')
        spacing.set(f'{{{NS_W}}}lineRule', 'auto')
        spacing.set(f'{{{NS_W}}}before', '120')
        spacing.set(f'{{{NS_W}}}after', '120')
        # pPr/rPr override: 24pt (12pt, 小四), 黑体, NOT bold (font moved from run to pPr/rPr so TOC inherits style font)
        ppr_rpr = _make_sub_element(ppr, 'rPr')
        rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
        rf_ppr.set(f'{{{NS_W}}}eastAsia', HEITI)
        _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 6
        _make_sub_element(ppr_rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(ppr_rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        r = _make_element('r')
        rpr = _make_sub_element(r, 'rPr')
        _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        # Explicit OFF override for bold inherited from template pStyle 6
        _make_sub_element(rpr, 'b', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(rpr, 'bCs', attrib={_W_VAL: '0'}, ns=NS_W)
        _make_sub_element(r, 't', text=title)
        # Wrap run in paragraph
        p.append(r)
        return _serialize(p)

    else:
        style_map = {2: '4', 3: '6'}
        style_id = style_map.get(level, '4')
        return make_paragraph(style_id, title)


def make_body_paragraph(text, indent=True):
    """Generate a body text paragraph with explicit formatting.

    Adds explicit firstLine indent (480 twips approx 2 chars) and font size (24=12pt)
    matching template body paragraph format. Detects citations for superscript.

    Args:
        text: Paragraph text.
        indent: True for style '18' (Body Text Indent), False for '17' (Body Text).

    Returns:
        XML string for a body text <w:p> element.
    """
    style_id = '18' if indent else '17'
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: style_id}, ns=NS_W)
    # Explicit line spacing: 360 (1.5x, per template)
    spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
    spacing.set(f'{{{NS_W}}}line', '360')
    spacing.set(f'{{{NS_W}}}lineRule', 'auto')
    # Explicit first-line indent (if style aa: 480 twips ≈ 2 Chinese chars)
    if indent:
        ind = _make_sub_element(ppr, 'ind', ns=NS_W)
        ind.set(f'{{{NS_W}}}firstLine', '480')
    # Explicit font size in pPr run properties, with east-asia font for Chinese characters
    ppr_rpr = _make_sub_element(ppr, 'rPr')
    rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
    rf_ppr.set(f'{{{NS_W}}}eastAsia', '宋体')
    _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
    # Detect citations and split text into segments (same as make_paragraph)
    parts = _CITATION_RE.split(text)
    has_multiple_parts = len(parts) > 1
    for part in parts:
        if not part:
            continue
        is_citation = has_multiple_parts and _CITATION_RE.fullmatch(part) is not None
        r = _make_sub_element(p, 'r')
        if is_citation:
            rpr = _make_sub_element(r, 'rPr')
            rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
            rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
            _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
            _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
            _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        else:
            rpr = _make_sub_element(r, 'rPr')
            rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
            rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
            _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
            _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
        _make_sub_element(r, 't', text=part)
    return _serialize(p)


def make_body_paragraph_with_refs(text, ref_map, indent=True):
    """Generate a body paragraph with REF field codes for cross-references.

    Splits text at \ref{...} patterns and generates proper run-level
    field codes for each cross-reference.

    Args:
        text: Paragraph text possibly containing \ref{...} patterns.
        ref_map: Dict mapping label -> {'bm': bookmark_name, 'num': number_str}.
        indent: True for first-line indent.

    Returns:
        XML string for a <w:p> element.
    """
    style_id = '18' if indent else '17'
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: style_id}, ns=NS_W)
    spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
    spacing.set(f'{{{NS_W}}}line', '360')
    spacing.set(f'{{{NS_W}}}lineRule', 'auto')
    if indent:
        ind = _make_sub_element(ppr, 'ind', ns=NS_W)
        ind.set(f'{{{NS_W}}}firstLine', '480')
    ppr_rpr = _make_sub_element(ppr, 'rPr')
    rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
    rf_ppr.set(f'{{{NS_W}}}eastAsia', '宋体')
    _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)

    # Wrap any bare fig:/tab: label references that lost their \ref{} wrapper
    # e.g., "如图fig:mech_diagram所示" → "如图\ref{fig:mech_diagram}所示"
    # IMPORTANT: skip labels already inside \ref{...} to avoid double-wrapping
    for label in ref_map:
        if label.startswith(('fig:', 'tab:')):
            ref_wrapped = f'\\ref{{{label}}}'
            if ref_wrapped not in text:
                text = text.replace(label, ref_wrapped)
    
    # Split by \ref{...} patterns
    ref_pat = re.compile(r'\\ref\{([^}]+)\}')
    parts = ref_pat.split(text)
    has_ref = len(parts) > 1

    for i, part in enumerate(parts):
        if not part:
            continue
        # Even indices: text, odd indices: ref labels (after split)
        if has_ref and i % 2 == 1:
            # This is a ref label — generate REF field
            label = part
            if label in ref_map:
                entry = ref_map[label]
                # Parse the field XML and add its children to p
                field_xml = _make_ref_field_internal(entry['bm'], entry['num'])
                # field_xml is a string of XML fragments, need to parse
                from xml.etree.ElementTree import fromstring
                try:
                    fragment = fromstring(f'<frag>{field_xml}</frag>')
                    for child in fragment:
                        p.append(child)
                except Exception:
                    # Fallback: just add text
                    r = _make_sub_element(p, 'r')
                    _make_sub_element(r, 't', text=entry['num'])
            else:
                print(f'  [WARN] Undefined cross-reference label: "\\ref{{{label}}}" — rendering as [{label}]')
                r = _make_sub_element(p, 'r')
                _make_sub_element(r, 't', text=f'[{label}]')
        else:
            # Regular text — handle citations too
            cite_parts = _CITATION_RE.split(part)
            has_cite = len(cite_parts) > 1
            for cp in cite_parts:
                if not cp:
                    continue
                is_cite = has_cite and _CITATION_RE.fullmatch(cp) is not None
                r = _make_sub_element(p, 'r')
                rpr = _make_sub_element(r, 'rPr')
                rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
                if is_cite:
                    _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
                    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                else:
                    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                _make_sub_element(r, 't', text=cp)
    return _serialize(p)


def _make_ref_field_internal(bookmark_name, fallback_text):
    """Generate REF field XML fragments (runs without paragraph wrapper).
    Uses proper namespace declaration for fromstring parsing.
    """
    ns_decl = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    return (f'<w:r {ns_decl}><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r {ns_decl}><w:instrText xml:space="preserve"> REF {bookmark_name} \\h </w:instrText></w:r>'
            f'<w:r {ns_decl}><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r {ns_decl}><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{fallback_text}</w:t></w:r>'
            f'<w:r {ns_decl}><w:fldChar w:fldCharType="end"/></w:r>')


def make_caption(text, is_table=False):
    """Generate a caption paragraph (style '50' = 图/表, centered).

    Per 北航2026模板:
      - 图题: 宋体五号不加粗 (is_table=False, default)
      - 表题: 黑体五号加粗 (is_table=True)
    """
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '50'}, ns=NS_W)
    _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'center'}, ns=NS_W)
    r = _make_sub_element(p, 'r')
    rpr = _make_sub_element(r, 'rPr')
    rf = _make_sub_element(rpr, 'rFonts', ns=NS_W)
    rf.set(f'{{{NS_W}}}eastAsia', '黑体' if is_table else '宋体')
    if is_table:
        _make_sub_element(rpr, 'b')
        _make_sub_element(rpr, 'bCs')
    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(r, 't', text=text)
    return _serialize(p)


def make_caption_with_bookmark(caption_text, bookmark_name, bookmark_id, is_table=False):
    """Generate a caption paragraph with a bookmark for cross-referencing.

    The bookmark wraps ONLY the number portion (e.g., "1.1" in "图1.1 说明")
    so that REF fields display just the number without the prefix/description text.

    Per 北航2026模板:
      - 图题: 宋体五号不加粗 (is_table=False, default)
      - 表题: 黑体五号加粗 (is_table=True)

    Args:
        caption_text: Caption text (e.g. "图1.1 技术路线图").
        bookmark_name: Name for the bookmark (e.g. "_Fig_1_1").
        bookmark_id: Unique integer ID for the bookmark.
        is_table: True for table captions (黑体加粗), False for figure captions (宋体).

    Returns:
        XML string for a caption <w:p> element with bookmarks.
    """
    import re as _re
    _cap_font = '黑体' if is_table else '宋体'
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: '50'}, ns=NS_W)
    jc = _make_sub_element(ppr, 'jc', ns=NS_W)
    jc.set(_W_VAL, 'center')

    # Parse caption to isolate the number portion for bookmarking
    # Expected: "图1.1 说明" or "表2.3 说明" → number="1.1"/"2.3"
    num_match = _re.match(r'^[图表](\d+\.\d+)(.*)', caption_text)
    if num_match:
        prefix = caption_text[0]  # "图" or "表"
        num_part = num_match.group(1)  # "1.1"
        rest = num_match.group(2)  # " 说明" or ""
        # Ensure space between number and title (e.g. "图1.1 标题" not "图1.1标题")
        if rest and not rest.startswith(' '):
            rest = ' ' + rest
    else:
        prefix = ''
        num_part = caption_text
        rest = ''

    # Prefix text run (e.g., "图" or "表")
    if prefix:
        rp = _make_sub_element(p, 'r')
        rprp = _make_sub_element(rp, 'rPr')
        rfp = _make_sub_element(rprp, 'rFonts', ns=NS_W)
        rfp.set(f'{{{NS_W}}}eastAsia', _cap_font)
        if is_table:
            _make_sub_element(rprp, 'b')
            _make_sub_element(rprp, 'bCs')
        _make_sub_element(rprp, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(rprp, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(rp, 't', text=prefix)

    # Bookmark start — wraps ONLY the number portion
    bm_start = _make_element('bookmarkStart', ns=NS_W)
    bm_start.set(f'{{{NS_W}}}id', str(bookmark_id))
    bm_start.set(f'{{{NS_W}}}name', bookmark_name)
    p.append(bm_start)

    # Number text (this is what REF fields display)
    rn = _make_sub_element(p, 'r')
    rprn = _make_sub_element(rn, 'rPr')
    rfn = _make_sub_element(rprn, 'rFonts', ns=NS_W)
    rfn.set(f'{{{NS_W}}}eastAsia', _cap_font)
    if is_table:
        _make_sub_element(rprn, 'b')
        _make_sub_element(rprn, 'bCs')
    _make_sub_element(rprn, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(rprn, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(rn, 't', text=num_part)

    # Bookmark end
    bm_end = _make_element('bookmarkEnd', ns=NS_W)
    bm_end.set(f'{{{NS_W}}}id', str(bookmark_id))
    p.append(bm_end)

    # Explicit space run after bookmark to ensure visible spacing in Word
    rs = _make_sub_element(p, 'r')
    rspr = _make_sub_element(rs, 'rPr')
    rsf = _make_sub_element(rspr, 'rFonts', ns=NS_W)
    rsf.set(f'{{{NS_W}}}eastAsia', _cap_font)
    if is_table:
        _make_sub_element(rspr, 'b')
        _make_sub_element(rspr, 'bCs')
    _make_sub_element(rspr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(rspr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
    _make_sub_element(rs, 't', text=' ', attrib={'{http://www.w3.org/XML/1998/namespace}space': 'preserve'})

    # Rest of caption text (e.g., "说明")
    if rest:
        rr = _make_sub_element(p, 'r')
        rprr = _make_sub_element(rr, 'rPr')
        rfr = _make_sub_element(rprr, 'rFonts', ns=NS_W)
        rfr.set(f'{{{NS_W}}}eastAsia', _cap_font)
        if is_table:
            _make_sub_element(rprr, 'b')
            _make_sub_element(rprr, 'bCs')
        _make_sub_element(rprr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(rprr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(rr, 't', text=rest.lstrip())

    return _serialize(p)


def make_ref_field(bookmark_name, fallback_text):
    """Generate a Word REF field code for cross-referencing.

    Args:
        bookmark_name: Target bookmark name (e.g. "_Fig_1_1").
        fallback_text: Text to display before Word updates the field.

    Returns:
        XML string for REF field runs.
    """
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> REF {bookmark_name} \\h </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{fallback_text}</w:t></w:r>'
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def make_center_text(text):
    """Generate a centered paragraph with no style specified.

    Args:
        text: Text content.

    Returns:
        XML string for a centered <w:p> element.
    """
    return make_paragraph(None, text, alignment='center')


def make_equation_block(equation_text, equation_number):
    """Generate an equation paragraph with Cambria Math font, centered equation
    and right-aligned number using a tab stop.

    Args:
        equation_text: The equation content (e.g. 'y = mx + b').
        equation_number: The equation number (e.g. '(1-1)').

    Returns:
        XML string for an equation <w:p> element.
    """
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    
    # Set right-aligned tab stop at 9072 twips (right margin for A4 with 1417 margins)
    # 11906 (page width) - 1417 (left) - 1417 (right) = 9072 text width
    tabs = _make_sub_element(ppr, 'tabs')
    _make_sub_element(tabs, 'tab', attrib={_W_VAL: 'right', f'{{{NS_W}}}pos': '9072'})
    
    # Paragraph spacing to visually separate equations
    spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
    spacing.set(f'{{{NS_W}}}line', '360')
    spacing.set(f'{{{NS_W}}}lineRule', 'auto')
    spacing.set(f'{{{NS_W}}}before', '120')
    spacing.set(f'{{{NS_W}}}after', '120')

    # Equation text run — Cambria Math font, italic
    r1 = _make_sub_element(p, 'r')
    rpr1 = _make_sub_element(r1, 'rPr')
    rf1 = _make_sub_element(rpr1, 'rFonts', ns=NS_W)
    rf1.set(f'{{{NS_W}}}ascii', 'Cambria Math')
    rf1.set(f'{{{NS_W}}}hAnsi', 'Cambria Math')
    rf1.set(f'{{{NS_W}}}eastAsia', 'Cambria Math')
    _make_sub_element(rpr1, 'i')
    _make_sub_element(rpr1, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(rpr1, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(r1, 't', text=equation_text)

    # Tab character run
    r_tab = _make_sub_element(p, 'r')
    _make_sub_element(r_tab, 'tab')
    
    # Equation number run — Cambria Math font, normal weight
    r2 = _make_sub_element(p, 'r')
    rpr2 = _make_sub_element(r2, 'rPr')
    rf2 = _make_sub_element(rpr2, 'rFonts', ns=NS_W)
    rf2.set(f'{{{NS_W}}}ascii', 'Cambria Math')
    rf2.set(f'{{{NS_W}}}hAnsi', 'Cambria Math')
    rf2.set(f'{{{NS_W}}}eastAsia', 'Cambria Math')
    _make_sub_element(rpr2, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(rpr2, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(r2, 't', text=equation_number)

    return _serialize(p)


def _make_cell_paragraph(cell_text, header=False):
    """Build a w:p element for a table cell, handling $...$ math segments.
    
    Font formatting MUST be explicit: Word's tblStylePr firstRow is NOT always
    honored for East Asian fonts. Header cells get 黑体 bold, data cells get 宋体.
    Math segments ($...$) are converted to OMML via latex_to_omath.
    
    Args:
        cell_text: Cell text with optional $...$ math segments.
        header: True for header cells (黑体 bold), False for data (宋体).
    
    Returns:
        ElementTree <w:p> element ready to append into a <w:tc>.
    """
    import re as _re
    font_name = '宋体'  # All table content 宋体五号不加粗 per user req
    
    # Create paragraph
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'center'}, ns=NS_W)
    
    def _add_text_run(parent, text):
        """Append a simple w:r element to parent."""
        r = _make_sub_element(parent, 'r')
        rpr = _make_sub_element(r, 'rPr')
        rf = _make_sub_element(rpr, 'rFonts', ns=NS_W)
        rf.set(f'{{{NS_W}}}eastAsia', font_name)
        # No bold in table cells (user req: 宋体五号不加粗)
        _make_sub_element(rpr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
        _make_sub_element(r, 't', text=text)
    
    if '$' not in cell_text:
        _add_text_run(p, cell_text)
        return p
    
    # Has math - split by $...$
    math_re = _re.compile(r'\$(.*?)\$', _re.DOTALL)
    parts = []
    last_end = 0
    for m in math_re.finditer(cell_text):
        if m.start() > last_end:
            parts.append(('text', cell_text[last_end:m.start()]))
        math_content = m.group(1)
        if math_content:
            parts.append(('math', math_content))
        last_end = m.end()
    if last_end < len(cell_text):
        parts.append(('text', cell_text[last_end:]))
    
    if not parts:
        _add_text_run(p, '')
        return p
    
    from omml_generator import latex_to_omath
    
    for ptype, pcontent in parts:
        if ptype == 'text':
            if not pcontent:
                continue
            _add_text_run(p, pcontent)
        elif ptype == 'math':
            try:
                math_omml = latex_to_omath(pcontent, display=False)
                if math_omml:
                    # Parse OMML fragment with proper namespace context
                    # OMML output contains <m:> AND <w:> elements (e.g. <w:rPr>, <w:rFonts>, <w:sz>)
                    wrapped = (
                        '<_root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
                        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                        + math_omml
                        + '</_root>'
                    )
                    root = ET.fromstring(wrapped)
                    for child in root:
                        p.append(child)
                else:
                    # OMML conversion returned empty — fallback italic
                    r = _make_sub_element(p, 'r')
                    rpr = _make_sub_element(r, 'rPr')
                    rf = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                    rf.set(f'{{{NS_W}}}eastAsia', font_name)
                    _make_sub_element(rpr, 'i')
                    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
                    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
                    _make_sub_element(r, 't', text=pcontent)
            except Exception:
                # Fallback: render as italic text
                r = _make_sub_element(p, 'r')
                rpr = _make_sub_element(r, 'rPr')
                rf = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                rf.set(f'{{{NS_W}}}eastAsia', font_name)
                _make_sub_element(rpr, 'i')
                _make_sub_element(rpr, 'sz', attrib={_W_VAL: '21'}, ns=NS_W)
                _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '21'}, ns=NS_W)
                _make_sub_element(r, 't', text=pcontent)
    
    return p


def make_table(headers, rows, caption_text=None, bookmark_name=None, bookmark_id=None):
    """Generate a complete <w:tbl> element as a string in 三线表 (three-line table) format.

    三线表 borders: thick top (1.5pt), thin header-bottom (0.5pt via insideH),
    thick bottom (1.5pt). No vertical lines, no left/right side borders.

    Args:
        headers: List of header cell text strings.
        rows: List of rows, each a list of cell text strings.
        caption_text: Optional caption text (prepended as a caption paragraph).
        bookmark_name: Optional bookmark name for cross-referencing.
        bookmark_id: Optional bookmark ID (required if bookmark_name is set).

    Returns:
        XML string for a <w:tbl> element (with optional caption).
    """
    parts = []

    # Spacer paragraph before table block (body text → table gap)
    before_p = _make_element('p')
    before_ppr = _make_sub_element(before_p, 'pPr')
    _make_sub_element(before_ppr, 'spacing', attrib={
        f'{{{NS_W}}}before': '120',
        f'{{{NS_W}}}after': '120',
    }, ns=NS_W)
    parts.append(_serialize(before_p))

    if caption_text:
        if bookmark_name and bookmark_id is not None:
            caption_xml = make_caption_with_bookmark(caption_text, bookmark_name, bookmark_id, is_table=True)
        else:
            caption_xml = make_caption(caption_text, is_table=True)
        # Add w:spacing before="120" after="120" (6pt each) to caption pPr for gaps
        # before caption (body text → table) and between caption and table
        caption_xml = caption_xml.replace('</w:pPr>', '<w:spacing w:before="120" w:after="120"/></w:pPr>', 1)
        parts.append(caption_xml)
    else:
        # No caption — add a before-table spacer paragraph for vertical gap from body text
        before_p = _make_element('p')
        before_ppr = _make_sub_element(before_p, 'pPr')
        _make_sub_element(before_ppr, 'spacing', attrib={
            f'{{{NS_W}}}before': '60',
            f'{{{NS_W}}}after': '60',
        }, ns=NS_W)
        parts.append(_serialize(before_p))

    tbl = _make_element('tbl')

    # Table properties
    tblPr = _make_sub_element(tbl, 'tblPr')

    # Table width: 9000 dxa (fills page width within A4 margins)
    _make_sub_element(tblPr, 'tblW',
                      attrib={_W_W: '9000', _W_TYPE: 'dxa'},
                      ns=NS_W)

    # Table style: BUAA Thesis Table — handles font/size/alignment via tblStylePr conditional formatting
    _make_sub_element(tblPr, 'tblStyle', attrib={_W_VAL: 'BUAA_Thesis_Table'}, ns=NS_W)

    # Table centering
    _make_sub_element(tblPr, 'jc', attrib={_W_VAL: 'center'}, ns=NS_W)

    # 三线表 borders: thick top/bottom (12=1.5pt), thin header-bottom (4=0.5pt via insideH)
    # NO insideV, left, or right —三线表 has only horizontal lines
    tbl_borders = _make_sub_element(tblPr, 'tblBorders')
    for border_name, sz in [('top', 12), ('bottom', 12)]:
        b = _make_sub_element(tbl_borders, border_name)
        b.set(_W_VAL, 'single')
        b.set(f'{{{NS_W}}}sz', str(sz))
        b.set(f'{{{NS_W}}}space', '0')
        b.set(f'{{{NS_W}}}color', '000000')
    # insideH = header-data separator (replaces per-cell tcBorders bottom)
    ib = _make_sub_element(tbl_borders, 'insideH')
    ib.set(_W_VAL, 'single')
    ib.set(f'{{{NS_W}}}sz', '4')
    ib.set(f'{{{NS_W}}}space', '0')
    ib.set(f'{{{NS_W}}}color', '000000')

    # Cell margins (from Normal Table style: left/right=108dxa, top/bottom=0)
    tbl_cell_mar = _make_sub_element(tblPr, 'tblCellMar')
    for side, val in [('top', 0), ('left', 108), ('bottom', 0), ('right', 108)]:
        m = _make_sub_element(tbl_cell_mar, side)
        m.set(f'{{{NS_W}}}w', str(val))
        m.set(f'{{{NS_W}}}type', 'dxa')

    # Header row (formatting handled by BT1 style tblStylePr firstRow)
    if headers:
        tr_header = _make_sub_element(tbl, 'tr')
        for header_text in headers:
            tc = _make_sub_element(tr_header, 'tc')
            tc_pr = _make_sub_element(tc, 'tcPr')
            _make_sub_element(tc_pr, 'vAlign', attrib={_W_VAL: 'center'}, ns=NS_W)
            # Header-data separator line is handled by table-level insideH border (above)
            # No per-cell tcBorders needed — avoids interference with page header line
            # Header cell paragraph with math support ($...$ → OMML)
            tc_p = _make_cell_paragraph(header_text, header=True)
            tc.append(tc_p)

    # Data rows — explicit font via properly namespaced elements
    for row in rows:
        tr = _make_sub_element(tbl, 'tr')
        for cell_text in row:
            tc = _make_sub_element(tr, 'tc')
            tc_p = _make_cell_paragraph(cell_text, header=False)
            tc.append(tc_p)

    parts.append(_serialize(tbl))

    # Spacer paragraph after table block (table → body text gap)
    after_p = _make_element('p')
    after_ppr = _make_sub_element(after_p, 'pPr')
    _make_sub_element(after_ppr, 'spacing', attrib={
        f'{{{NS_W}}}before': '120',
        f'{{{NS_W}}}after': '120',
    }, ns=NS_W)
    parts.append(_serialize(after_p))

    return '\n'.join(parts)


def get_image_dimensions(img_path):
    """Get image dimensions in EMU, scaled to fit page width.

    Uses PIL to read actual pixel dimensions, then scales to fit
    within max_width_emu (default: 5040000 ≈ 14cm, suitable for A4 margins).

    Args:
        img_path: Path to image file.

    Returns:
        (width_emu, height_emu) scaled to fit.
    """
    default_w, default_h = 5040000, 3600000
    max_w = 5040000  # Max width in EMU (~14cm)

    if not img_path or not os.path.isfile(img_path):
        return default_w, default_h

    try:
        from PIL import Image
        with Image.open(img_path) as img:
            w_px, h_px = img.size
    except Exception:
        return default_w, default_h

    # Convert pixels to EMU assuming 96 DPI
    factor = 914400 / 96  # 1 pixel at 96 DPI = 9525 EMU
    w_emu = int(w_px * factor)
    h_emu = int(h_px * factor)

    # Scale to fit max width, preserving aspect ratio
    if w_emu > max_w:
        scale = max_w / w_emu
        w_emu = int(w_emu * scale)
        h_emu = int(h_emu * scale)

    return w_emu, h_emu


def make_image_tag(rel_id, img_path=None, width_emu=None, height_emu=None, img_id=1):
    """Generate a <w:drawing> inside a <w:p> wrapper (paragraph containing image).

    Default size: 14cm x 10cm (5040000 x 3600000 EMU).

    Args:
        rel_id: Relationship ID for the image (e.g. 'rId5').
        width_emu: Image width in EMU.
        height_emu: Image height in EMU.
        img_id: Unique image ID for wp:docPr and pic:cNvPr (default=1, caller should increment).

    Returns:
        XML string for a <w:p> element containing the <w:r> with <w:drawing>.
    """
    # Auto-detect dimensions from image file if not explicitly provided
    if img_path is not None and (width_emu is None or height_emu is None):
        width_emu, height_emu = get_image_dimensions(img_path)
    if width_emu is None:
        width_emu = 5040000
    if height_emu is None:
        height_emu = 3600000

    # Create w:p wrapper with centered alignment
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'jc', attrib={_W_VAL: 'center'}, ns=NS_W)
    
    # Create w:r inside w:p
    r = _make_sub_element(p, 'r')
    drawing = _make_sub_element(r, 'drawing')

    inline = _make_sub_element(drawing, 'inline', ns=NS_WP)
    inline.set('distT', '0')
    inline.set('distB', '0')
    inline.set('distL', '0')
    inline.set('distR', '0')

    _make_sub_element(inline, 'extent', ns=NS_WP,
                      attrib={'cx': str(width_emu), 'cy': str(height_emu)})
    _make_sub_element(inline, 'effectExtent', ns=NS_WP,
                      attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
    _make_sub_element(inline, 'docPr', ns=NS_WP,
                      attrib={'id': str(img_id), 'name': 'Picture ' + str(img_id)})

    cNvGraphicFramePr = _make_sub_element(inline, 'cNvGraphicFramePr', ns=NS_WP)
    graphic_frame_locks = _make_sub_element(cNvGraphicFramePr, 'graphicFrameLocks', ns=NS_A)
    graphic_frame_locks.set('{http://schemas.openxmlformats.org/drawingml/2006/main}noChangeAspect', '1')

    graphic = _make_sub_element(inline, 'graphic', ns=NS_A)
    graphic_data = _make_sub_element(graphic, 'graphicData', ns=NS_A)
    graphic_data.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

    pic = _make_sub_element(graphic_data, 'pic', ns=NS_PIC)

    # nvPicPr
    nvPicPr = _make_sub_element(pic, 'nvPicPr', ns=NS_PIC)
    _make_sub_element(nvPicPr, 'cNvPr', ns=NS_PIC,
                      attrib={'id': str(img_id), 'name': 'Picture ' + str(img_id)})
    _make_sub_element(nvPicPr, 'cNvPicPr', ns=NS_PIC)

    # blipFill
    blip_fill = _make_sub_element(pic, 'blipFill', ns=NS_PIC)
    blip = _make_sub_element(blip_fill, 'blip', ns=NS_A)
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rel_id)
    stretch = _make_sub_element(blip_fill, 'stretch', ns=NS_A)
    _make_sub_element(stretch, 'fillRect', ns=NS_A)

    # spPr
    sp_pr = _make_sub_element(pic, 'spPr', ns=NS_PIC)
    xfrm = _make_sub_element(sp_pr, 'xfrm', ns=NS_A)
    _make_sub_element(xfrm, 'off', ns=NS_A, attrib={'x': '0', 'y': '0'})
    _make_sub_element(xfrm, 'ext', ns=NS_A,
                      attrib={'cx': str(width_emu), 'cy': str(height_emu)})
    _make_sub_element(sp_pr, 'prstGeom', ns=NS_A, attrib={'prst': 'rect'})

    return _serialize(p)


# Unicode superscript/subscript mappings for math fallback
_UNICODE_SUPER_FB = {
    '0': '\u2070', '1': '\u00b9', '2': '\u00b2', '3': '\u00b3', '4': '\u2074',
    '5': '\u2075', '6': '\u2076', '7': '\u2077', '8': '\u2078', '9': '\u2079',
    '+': '\u207a', '-': '\u207b', '=': '\u207c', '(': '\u207d', ')': '\u207e',
    'i': '\u2071', 'n': '\u207f', 'T': '\u1d40',
    'a': '\u1d43', 'b': '\u1d47', 'c': '\u1d9c', 'd': '\u1d48', 'e': '\u1d49',
    'f': '\u1da0', 'g': '\u1d4d', 'h': '\u02b0', 'j': '\u02b2', 'k': '\u1d4f',
    'l': '\u02e1', 'm': '\u1d50', 'o': '\u1d52', 'p': '\u1d56', 'r': '\u02b3',
    's': '\u02e2', 't': '\u1d57', 'u': '\u1d58', 'v': '\u1d5b', 'w': '\u02b7',
    'x': '\u02e3', 'y': '\u02b8', 'z': '\u1dbb',
}

_UNICODE_SUB_FB = {
    '0': '\u2080', '1': '\u2081', '2': '\u2082', '3': '\u2083', '4': '\u2084',
    '5': '\u2085', '6': '\u2086', '7': '\u2087', '8': '\u2088', '9': '\u2089',
    '+': '\u208a', '-': '\u208b', '=': '\u208c', '(': '\u208d', ')': '\u208e',
    'a': '\u2090', 'e': '\u2091', 'h': '\u2095', 'i': '\u1d62', 'j': '\u2c7c',
    'k': '\u2096', 'l': '\u2097', 'm': '\u2098', 'n': '\u2099', 'o': '\u2092',
    'p': '\u209a', 'r': '\u1d63', 's': '\u209b', 't': '\u209c', 'u': '\u1d64',
    'v': '\u1d65', 'x': '\u2093',
}


def _fallback_math_to_unicode(latex_text):
    """Convert simple inline LaTeX math to Unicode text when OMML fails.

    Handles Greek letters, common math operators, arrows, subscripts,
    superscripts, and fractions as readable Unicode text.
    """
    from omml_generator import _GREEK_LOWER, _GREEK_UPPER, _OPERATOR_SYMBOLS, _ARROWS
    import re as _fr

    result = latex_text

    for name, sym in _GREEK_LOWER.items():
        result = result.replace(f'\\{name}', sym)
    for name, sym in _GREEK_UPPER.items():
        result = result.replace(f'\\{name}', sym)
    for name, sym in _OPERATOR_SYMBOLS.items():
        result = result.replace(f'\\{name}', sym)
    for name, sym in _ARROWS.items():
        result = result.replace(f'\\{name}', sym)

    result = _fr.sub(r'\\frac\{([^}]*)\}\{([^}]*)\}', r'\\1/\\2', result)
    result = _fr.sub(r'\\sqrt\{([^}]*)\}', r'sqrt(\\1)', result)
    result = _fr.sub(r'\\text\{([^}]*)\}', r'\\1', result)
    result = _fr.sub(r'\\mathrm\{([^}]*)\}', r'\\1', result)
    result = _fr.sub(r'\\mathbf\{([^}]*)\}', r'\\1', result)
    result = _fr.sub(r'\\mathit\{([^}]*)\}', r'\\1', result)
    result = _fr.sub(r'\\mathbb\{([^}]*)\}', r'\\1', result)
    result = _fr.sub(r'\\mathcal\{([^}]*)\}', r'\\1', result)

    def _sub_to_unicode(m):
        inner = m.group(1).strip()
        mapped = ''.join(_UNICODE_SUB_FB.get(c, c) for c in inner)
        return mapped if mapped != inner else '_' + inner

    def _sup_to_unicode(m):
        inner = m.group(1).strip()
        mapped = ''.join(_UNICODE_SUPER_FB.get(c, c) for c in inner)
        return mapped if mapped != inner else '^' + inner

    result = _fr.sub(r'_\{\s*([^}]*?)\s*\}', _sub_to_unicode, result)
    result = _fr.sub(r'\^\{\s*([^}]*?)\s*\}', _sup_to_unicode, result)
    result = _fr.sub(r'_([0-9a-zA-Z])', lambda m: _UNICODE_SUB_FB.get(m.group(1), '_' + m.group(1)), result)
    result = _fr.sub(r'\^([0-9a-zA-Z])', lambda m: _UNICODE_SUPER_FB.get(m.group(1), '^' + m.group(1)), result)
    result = _fr.sub(r'\\sum', '\u03a3', result)
    result = _fr.sub(r'\\prod', '\u03a0', result)
    result = _fr.sub(r'\\int', '\u222b', result)
    result = _fr.sub(r'\\partial', '\u2202', result)
    result = _fr.sub(r'\\infty', '\u221e', result)
    result = _fr.sub(r'\\nabla', '\u2207', result)
    result = _fr.sub(r'\\left|\\right', '', result)
    result = _fr.sub(r'\\left\\.?|\\right\\.?', '', result)
    result = _fr.sub(r'\\([a-zA-Z]+)', r'\\1', result)
    result = result.replace('{', '').replace('}', '')
    return result.strip()


def make_body_paragraph_with_math(text, indent=True, ref_map=None):
    """Generate body paragraph with inline OMML math ($...$) embedded.

    Splits paragraph text at $...$ math delimiters. Non-math parts get
    normal w:t runs. Math parts get converted to OMML via omml_generator.

    Also handles citations [1] as superscripts and \\ref{...} as REF fields.

    Args:
        text: Paragraph text with possible $...$ math and \\ref{...} patterns.
        indent: True for first-line indent.
        ref_map: Optional dict for cross-reference labels.

    Returns:
        XML string for a w:p element.
    """
    import re as _re
    from omml_generator import latex_to_omath

    style_id = '18' if indent else '17'
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')
    _make_sub_element(ppr, 'pStyle', attrib={_W_VAL: style_id}, ns=NS_W)
    spacing = _make_sub_element(ppr, 'spacing', ns=NS_W)
    spacing.set(f'{{{NS_W}}}line', '360')
    spacing.set(f'{{{NS_W}}}lineRule', 'auto')
    if indent:
        ind = _make_sub_element(ppr, 'ind', ns=NS_W)
        ind.set(f'{{{NS_W}}}firstLine', '480')
    ppr_rpr = _make_sub_element(ppr, 'rPr')
    rf_ppr = _make_sub_element(ppr_rpr, 'rFonts', ns=NS_W)
    rf_ppr.set(f'{{{NS_W}}}eastAsia', '宋体')
    _make_sub_element(ppr_rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
    _make_sub_element(ppr_rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)

    if ref_map:
        for label in ref_map:
            if label.startswith(('fig:', 'tab:')):
                ref_wrapped = f'\\ref{{{label}}}'
                if ref_wrapped not in text:
                    text = text.replace(label, ref_wrapped)

    # Split by $...$ and $$...$$ displaying math
    math_re = _re.compile(r'\$\$(.*?)\$\$|\$(.*?)\$', _re.DOTALL)

    # Iterate through regex matches
    last_end = 0
    math_segments = []

    for m in math_re.finditer(text):
        # Text before this match
        prefix = text[last_end:m.start()]
        if prefix:
            math_segments.append(('text', prefix))
        # Math content
        is_display = m.group(1) is not None
        math_content = m.group(1) or m.group(2)
        if math_content:
            math_segments.append(('math', math_content.strip(), is_display))
        last_end = m.end()

    # Trailing text
    suffix = text[last_end:]
    if suffix:
        math_segments.append(('text', suffix))

    # If no math found, fall through to normal paragraph
    if all(s[0] == 'text' for s in math_segments):
        # Just use normal body paragraph
        return make_body_paragraph(text, indent=indent)

    # Build runs from segments
    for seg in math_segments:
        seg_type = seg[0]

        if seg_type == 'text':
            seg_text = seg[1]
            # Handle citations in text segments
            cite_parts = _CITATION_RE.split(seg_text)
            has_cite = len(cite_parts) > 1
            # Also handle \\ref{} in text
            ref_pat = _re.compile(r'\\ref\{([^}]+)\}')
            for cp in cite_parts:
                if not cp:
                    continue
                is_cite = has_cite and _CITATION_RE.fullmatch(cp) is not None

                if is_cite:
                    r = _make_sub_element(p, 'r')
                    rpr = _make_sub_element(r, 'rPr')
                    rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                    rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
                    _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
                    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(r, 't', text=cp)
                elif ref_map and '\\ref{' in cp:
                    # Handle \\ref{} in this text chunk
                    ref_parts = ref_pat.split(cp)
                    for ri, rp in enumerate(ref_parts):
                        if not rp:
                            continue
                        if ri % 2 == 1 and rp in ref_map:
                            entry = ref_map[rp]
                            bm = entry['bm']
                            num = entry['num']
                            field_xml = f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldChar w:fldCharType="begin"/></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:instrText xml:space="preserve"> REF {bm} \\h </w:instrText></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldChar w:fldCharType="separate"/></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{num}</w:t></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldChar w:fldCharType="end"/></w:r>'
                            try:
                                from xml.etree.ElementTree import fromstring
                                fragment = fromstring(f'<frag>{field_xml}</frag>')
                                for child in fragment:
                                    p.append(child)
                            except Exception:
                                r = _make_sub_element(p, 'r')
                                _make_sub_element(r, 't', text=rp)
                        else:
                            r = _make_sub_element(p, 'r')
                            rpr = _make_sub_element(r, 'rPr')
                            rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                            rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
                            _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                            _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                            _make_sub_element(r, 't', text=rp)
                else:
                    r = _make_sub_element(p, 'r')
                    rpr = _make_sub_element(r, 'rPr')
                    rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                    rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
                    _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                    _make_sub_element(r, 't', text=cp)

        elif seg_type == 'math':
            math_content = seg[1]
            is_display = seg[2]
            omml_ok = False
            try:
                omml = latex_to_omath(math_content, display=False)
                if omml:
                    # Parse OMML XML and append directly to paragraph.
                    # Must declare both m: and w: namespaces since OMML uses
                    # w:rPr / w:rFonts inside m:r and m:ctrlPr elements.
                    from xml.etree.ElementTree import fromstring, register_namespace
                    register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
                    register_namespace('w', NS_W)
                    ns_decl = ('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
                               f' xmlns:w="{NS_W}"')
                    fragment = fromstring(f'<frag {ns_decl}>{omml}</frag>')
                    for child in fragment:
                        p.append(child)
                    omml_ok = True
            except Exception as e:
                print(f'  [WARN] OMML inline math parse failed for "${math_content[:40]}...": {e}')

            if not omml_ok:
                # Fallback: convert LaTeX math to Unicode and insert as normal text run
                try:
                    fallback_text = _fallback_math_to_unicode(math_content)
                except Exception:
                    fallback_text = math_content
                r = _make_sub_element(p, 'r')
                rpr = _make_sub_element(r, 'rPr')
                rf_run = _make_sub_element(rpr, 'rFonts', ns=NS_W)
                rf_run.set(f'{{{NS_W}}}eastAsia', '宋体')
                _make_sub_element(rpr, 'sz', attrib={_W_VAL: '24'}, ns=NS_W)
                _make_sub_element(rpr, 'szCs', attrib={_W_VAL: '24'}, ns=NS_W)
                _make_sub_element(r, 't', text=fallback_text)

    return _serialize(p)


def make_equation_omml_paragraph(latex_text, eq_number=None):
    """Generate a centered paragraph with OMML display equation.

    Uses m:oMathPara for proper display equation rendering.
    Equation number is placed in a separate run with a right-aligned tab stop
    so it appears flush with the right margin.

    Args:
        latex_text: LaTeX equation content.
        eq_number: Optional equation number, e.g. '(2-1)'.

    Returns:
        XML string for a w:p element with embedded OMML.
    """
    from omml_generator import latex_to_omath

    # Build OMML with equation number embedded via # prefix in eqArr
    omml = latex_to_omath(latex_text, display=True, eq_number=eq_number)
    if not omml:
        return make_equation_block(latex_text, eq_number or '')

    # Build the paragraph as raw XML string (avoid ElementTree namespace issues)
    parts = ['<w:p>',
             '<w:pPr>',
             '<w:jc w:val="center"/>',
             '<w:tabs>',
             '<w:tab w:val="right" w:leader="none" w:pos="9072"/>',
             '</w:tabs>',
             '<w:spacing w:line="360" w:lineRule="auto" w:before="120" w:after="120"/>',
             '</w:pPr>',
             omml,
             '</w:p>']
    return ''.join(parts)


def make_rich_paragraph(style_id, runs_list, alignment=None):
    """Generate a paragraph with multiple runs for mixed formatting.

    Args:
        style_id: Style ID string, or None to omit pStyle.
        runs_list: List of dicts, each with keys:
            - 'text' (str, required)
            - 'bold' (bool, optional)
            - 'italic' (bool, optional)
            - 'subscript' (bool, optional)
            - 'superscript' (bool, optional)
        alignment: 'center', 'left', 'right', 'both'.

    Returns:
        XML string for a <w:p> element with multiple runs.
    """
    p = _make_element('p')
    ppr = _make_sub_element(p, 'pPr')

    if style_id is not None:
        _make_sub_element(ppr, 'pStyle',
                          attrib={_W_VAL: style_id},
                          ns=NS_W)

    if alignment is not None:
        align_val = _ALIGN_MAP.get(alignment, alignment)
        _make_sub_element(ppr, 'jc',
                          attrib={_W_VAL: align_val},
                          ns=NS_W)

    for run_info in runs_list:
        text = run_info.get('text', '')
        bold = run_info.get('bold', False)
        italic = run_info.get('italic', False)
        subscript = run_info.get('subscript', False)
        superscript = run_info.get('superscript', False)

        r = _make_sub_element(p, 'r')
        has_rpr = bold or italic or subscript or superscript
        if has_rpr:
            rpr = _make_sub_element(r, 'rPr')
            if bold:
                _make_sub_element(rpr, 'b')
            if italic:
                _make_sub_element(rpr, 'i')
            if subscript:
                _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'subscript'})
            if superscript:
                _make_sub_element(rpr, 'vertAlign', attrib={_W_VAL: 'superscript'})
        _make_sub_element(r, 't', text=text)

    return _serialize(p)
