"""
flatopc_to_docx.py - v2 with built-in XML repair.

Fixes template tag mismatches during conversion:
1. <w:t>text</w:r> -> </w:t>
2. Unclosed <w:r> before </w:p>
3. Missing </w:p> at end of body
4. lxml recovery for simple parts (headers)
"""

import sys, re, base64, zipfile
from lxml import etree

def repair_xml_text(text):
    """Fix common tag mismatches in Word XML text."""
    # Fix <w:t>text</w:r> -> </w:t>
    for m in re.finditer(r'<w:t[\s>]', text):
        t_end = m.end()
        nr = text.find('</w:r>', t_end)
        nt = text.find('</w:t>', t_end)
        if nr >= 0 and (nt < 0 or nr < nt):
            if '<' not in text[t_end:nr]:
                text = text[:nr] + '</w:t>' + text[nr+6:]
    
    # Fix unclosed <w:r> before </w:p> (per-paragraph)
    pos = 0
    while True:
        ps = text.find('<w:p', pos)
        if ps < 0:
            break
        pe = text.find('</w:p>', ps)
        if pe < 0:
            break
        para = text[ps:pe]
        ro = len(re.findall(r'<w:r[\s>]', para))
        rc = len(re.findall(r'</w:r>', para))
        if ro > rc:
            missing = ro - rc
            text = text[:pe] + '</w:r>' * missing + text[pe:]
            pe += 6 * missing
        pos = pe + 6
    
    # Fix missing </w:p> before </w:body>
    body_end = text.find('</w:body>')
    if body_end >= 0:
        before = text[:body_end]
        wp_o = len(re.findall(r'<w:p[\s>]', before))
        wp_c = len(re.findall(r'</w:p>', before))
        if wp_o > wp_c:
            text = text[:body_end] + '</w:p>' * (wp_o - wp_c) + text[body_end:]
    
    return text


def flatopc_to_docx(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        content_types = {}
        parts_data = {}

        part_re = re.compile(
            r'<pkg:part\s+'
            r'pkg:name="(?P<name>[^"]+)"\s+'
            r'(?:pkg:contentType="(?P<ctype>[^"]*)")?'
            r'[^>]*>'
            r'(?P<body>.*?)'
            r'</pkg:part>',
            re.DOTALL
        )

        for match in part_re.finditer(content):
            name = match.group('name')
            ctype = match.group('ctype') or ''
            body = match.group('body')

            archive_name = name.lstrip('/')
            if ctype:
                content_types[archive_name] = ctype

            xml_match = re.search(r'<pkg:xmlData[^>]*>(.*?)</pkg:xmlData>', body, re.DOTALL)
            if xml_match:
                xml_content = xml_match.group(1)
                
                # REPAIR XML before encoding
                # Fix tag mismatches
                xml_content = repair_xml_text(xml_content)
                
                # Try to validate with lxml; if fails, use recovery
                try:
                    etree.fromstring(xml_content.encode('utf-8'))
                except Exception:
                    # Check if this is a complex part (document.xml) - use text-level repair only
                    if '/word/document.xml' in name:
                        # text-level repair already done above
                        # Also try to fix remaining issues with regex
                        # Fix orphan </w:r> before first <w:r
                        first_r = xml_content.find('<w:r')
                        while True:
                            first_r_close = xml_content.find('</w:r>')
                            if first_r_close >= 0 and (first_r < 0 or first_r_close < first_r):
                                xml_content = xml_content[:first_r_close] + xml_content[first_r_close+6:]
                            else:
                                break
                    else:
                        # Simpler parts (headers) - use lxml recovery
                        try:
                            parser = etree.XMLParser(recover=True)
                            tree = etree.fromstring(xml_content.encode('utf-8'), parser)
                            xml_content = etree.tostring(tree, encoding='unicode')
                        except:
                            pass
                
                parts_data[archive_name] = (xml_content.encode('utf-8'), False)
            else:
                bin_match = re.search(r'<pkg:binaryData>(.*?)</pkg:binaryData>', body, re.DOTALL)
                if bin_match:
                    b64_data = bin_match.group(1)
                    b64_clean = re.sub(r'\s', '', b64_data)
                    try:
                        binary_data = base64.b64decode(b64_clean)
                        parts_data[archive_name] = (binary_data, True)
                    except Exception as e:
                        print(f"Warning: Could not decode {name}: {e}")

        # Generate [Content_Types].xml
        overrides = []
        defaults_added = set()
        ext_map = {
            'rels': 'application/vnd.openxmlformats-package.relationships+xml',
            'xml': 'application/xml',
            'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
            'gif': 'image/gif', 'bmp': 'image/bmp', 'tiff': 'image/tiff',
            'wmf': 'image/x-wmf', 'emf': 'image/x-emf',
            'bin': 'application/vnd.openxmlformats-officedocument.oleObject',
        }

        for archive_name, ctype in sorted(content_types.items()):
            ext = archive_name.rsplit('.', 1)[-1] if '.' in archive_name else ''
            if ext in ext_map and ext_map[ext] == ctype:
                if ext not in defaults_added:
                    defaults_added.add(ext)
                continue
            else:
                overrides.append(f'    <Override PartName="/{archive_name}" ContentType="{ctype}"/>')

        defaults_xml = ''.join(f'    <Default Extension="{ext}" ContentType="{ext_map[ext]}"/>\n'
                               for ext in sorted(defaults_added))
        ct_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
{defaults_xml}{''.join(overrides)}
</Types>'''

        zf.writestr('[Content_Types].xml', ct_xml.encode('utf-8'))
        for archive_name, (data, is_binary) in sorted(parts_data.items()):
            zf.writestr(archive_name, data)

    print(f"Created: {output_path} ({len(parts_data)} parts)")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python flatopc_to_docx.py input.xml output.docx")
        sys.exit(1)
    flatopc_to_docx(sys.argv[1], sys.argv[2])
