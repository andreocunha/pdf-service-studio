#!/usr/bin/env python3
"""Compare final download PDFs, not raw renders from different platforms.

Requires pypdf and Poppler's pdftotext. Usage:
  python3 scripts/check-pdf-regression.py reference.pdf candidate.pdf
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET
from pypdf import PdfReader
from pypdf.generic import IndirectObject

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('reference', type=Path)
parser.add_argument('candidate', type=Path)
parser.add_argument('--max-size-ratio', type=float, default=1.05)
parser.add_argument('--link-coordinate-tolerance', type=float, default=0,
                    help='Allowed PDF-point rounding for annotation rectangles only; destinations stay exact')
parser.add_argument('--fit-width-links', action='store_true',
                    help='Allow XYZ inherited zoom to become FitH at the exact same page/y coordinate')
args = parser.parse_args()


def audit(path):
    reader = PdfReader(path)
    page_refs = {p.indirect_reference.idnum: i for i, p in enumerate(reader.pages)}
    seen = set()
    fonts = []

    def visit(obj):
        if isinstance(obj, IndirectObject):
            if obj.idnum in seen:
                return
            seen.add(obj.idnum)
            obj = obj.get_object()
        if isinstance(obj, dict):
            # CFF/color glyphs can be embedded as Type 3 CharProcs instead of
            # FontFile streams. They are fonts too, and must not escape this
            # check (Carbona documents exercise this path).
            if obj.get('/Subtype') == '/Type3' and '/CharProcs' in obj:
                glyphs = obj['/CharProcs'].get_object()
                digest = hashlib.sha256()
                for name, stream in sorted(glyphs.items()):
                    digest.update(str(name).encode())
                    digest.update(stream.get_object().get_data())
                for key in ('/FontMatrix', '/FontBBox', '/FirstChar', '/LastChar', '/Widths'):
                    value = obj.get(key)
                    digest.update(str(value.get_object() if isinstance(value, IndirectObject) else value).encode())
                fonts.append(('Type3', digest.hexdigest()))
            for key, value in obj.items():
                if key in ('/FontFile', '/FontFile2', '/FontFile3'):
                    font_name = re.sub(r'^/[A-Z]{6}\+', '', str(obj.get('/FontName', '')))
                    fonts.append((font_name, hashlib.sha256(value.get_object().get_data()).hexdigest()))
                elif key not in ('/Parent', '/P'):
                    visit(value)
        elif isinstance(obj, list):
            for value in obj:
                visit(value)

    def destination(value):
        # Compare actual targets, allowing named destinations to become explicit
        # GoTo actions without hiding a changed page/position/zoom.
        if isinstance(value, IndirectObject):
            value = value.get_object()
        if isinstance(value, str) and value in reader.named_destinations:
            value = reader.named_destinations[value].dest_array
        if isinstance(value, dict):
            value = value.get('/D', value)
        if isinstance(value, list):
            return [page_refs.get(value[0].idnum) if isinstance(value[0], IndirectObject) else value[0],
                    *[float(v) if isinstance(v, (int, float)) else str(v) for v in value[1:]]]
        return str(value)

    links = []
    for index, page in enumerate(reader.pages):
        visit(page)
        for ref in page.get('/Annots', []):
            annotation = ref.get_object()
            if annotation.get('/Subtype') != '/Link':
                continue
            action = annotation.get('/A', {}).get_object() if '/A' in annotation else {}
            links.append({'page': index, 'rect': [float(x) for x in annotation['/Rect']],
                          'uri': str(action.get('/URI', '')),
                          'dest': destination(annotation.get('/Dest', action.get('/D', '')))})
    xml = subprocess.check_output(['pdftotext', '-bbox', str(path), '-'])
    root = ET.fromstring(xml)
    ns = '{http://www.w3.org/1999/xhtml}'
    words = []
    pages = root.findall(f'.//{ns}page')
    for index, page in enumerate(pages):
        for word in page.findall(f'.//{ns}word'):
            words.append((index, word.text, word.attrib))
    return {'pages': [(p.mediabox, p.cropbox, p.rotation) for p in reader.pages],
            'words': words, 'fonts': sorted(fonts), 'links': links, 'bytes': path.stat().st_size}


reference = audit(args.reference)
candidate = audit(args.candidate)
errors = []
for key in ('pages', 'words', 'fonts'):
    if reference[key] != candidate[key]:
        errors.append(f'{key} changed')
links_equal = len(reference['links']) == len(candidate['links'])
for before, after in zip(reference['links'], candidate['links']):
    links_equal &= all(before[k] == after[k] for k in ('page', 'uri'))
    dest_equal = before['dest'] == after['dest']
    old, new = before['dest'], after['dest']
    if (args.fit_width_links and isinstance(old, list) and len(old) == 5
            and old[1] == '/XYZ' and old[4] in (0, 'NullObject')
            and isinstance(old[3], (int, float))):
        dest_equal |= new == [old[0], '/FitH', old[3]]
    links_equal &= dest_equal
    links_equal &= all(abs(x - y) <= args.link_coordinate_tolerance
                       for x, y in zip(before['rect'], after['rect']))
if not links_equal:
    errors.append('links changed')
if candidate['bytes'] > reference['bytes'] * args.max_size_ratio:
    errors.append(f"size increased beyond {args.max_size_ratio:.0%}")
report = {'referenceBytes': reference['bytes'], 'candidateBytes': candidate['bytes'],
          'pages': len(candidate['pages']), 'words': len(candidate['words']),
          'fontFiles': len(candidate['fonts']), 'links': len(candidate['links']), 'errors': errors}
print(json.dumps(report, indent=2))
if errors:
    raise SystemExit(1)
print('PASS: exact text positions, embedded font bytes, page geometry and link destinations preserved. '
      f'Annotation rectangle tolerance: {args.link_coordinate_tolerance} pt.')
