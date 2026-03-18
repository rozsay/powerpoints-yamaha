#!/usr/bin/env python3
"""
Extract all text content, slide structure, and meaningful information
from a PowerPoint (.pptx) file using only Python standard library.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re
from collections import defaultdict

PPTX_PATH = "/home/user/powerpoints-yamaha/Yamaha.pptx"

# XML namespaces used in PPTX format
NS = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
}

def get_text_from_element(elem):
    """Recursively extract all text runs from a DrawingML element."""
    texts = []
    for t in elem.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def get_paragraphs_from_element(elem):
    """Extract paragraphs (a:p) preserving structure."""
    paragraphs = []
    for para in elem.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
        runs = []
        for r in para.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}r'):
            t = r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}t')
            if t is not None and t.text:
                runs.append(t.text)
        text = ''.join(runs).strip()
        if text:
            paragraphs.append(text)
    return paragraphs

def parse_slide(zf, slide_path, rels_path):
    """Parse a single slide XML and return structured content."""
    slide_data = {
        'title': None,
        'shapes': [],
        'notes': None,
        'hyperlinks': [],
        'images': [],
    }

    # Load relationships for this slide
    rels = {}
    try:
        with zf.open(rels_path) as f:
            rel_tree = ET.parse(f)
            for rel in rel_tree.iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rels[rel.get('Id')] = {
                    'type': rel.get('Type', '').split('/')[-1],
                    'target': rel.get('Target', ''),
                }
    except KeyError:
        pass

    with zf.open(slide_path) as f:
        tree = ET.parse(f)

    root = tree.getroot()

    # Find all shape trees
    sp_tree = root.findall(
        './/{http://schemas.openxmlformats.org/presentationml/2006/main}cSld'
        '/{http://schemas.openxmlformats.org/drawingml/2006/main}spTree'
        if False else
        './/{http://schemas.openxmlformats.org/drawingml/2006/main}spTree'
    )

    for tree_elem in sp_tree:
        # Process each shape (sp)
        for sp in tree_elem.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
            shape_info = {}

            # Get shape name
            nvSpPr = sp.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr')
            if nvSpPr is None:
                nvSpPr = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}nvSpPr')
            cNvPr = None
            if nvSpPr is not None:
                cNvPr = nvSpPr.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr')
                if cNvPr is None:
                    cNvPr = nvSpPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
            shape_name = cNvPr.get('name', '') if cNvPr is not None else ''

            # Check placeholder type
            ph = sp.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}ph')
            ph_type = ph.get('type', 'body') if ph is not None else None

            # Get text body
            txBody = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}txBody')
            if txBody is not None:
                paras = get_paragraphs_from_element(txBody)
                full_text = '\n'.join(paras)
                if full_text.strip():
                    shape_info['name'] = shape_name
                    shape_info['ph_type'] = ph_type
                    shape_info['text'] = full_text
                    shape_info['paragraphs'] = paras

                    # Identify title
                    if ph_type in ('title', 'ctrTitle') and slide_data['title'] is None:
                        slide_data['title'] = full_text.strip()

                    slide_data['shapes'].append(shape_info)

        # Process graphic frames (tables, charts, SmartArt)
        for graphicFrame in tree_elem.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame'):
            graphic = graphicFrame.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
            if graphic is not None:
                graphicData = graphic.find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
                if graphicData is not None:
                    uri = graphicData.get('uri', '')
                    frame_info = {'type': uri.split('/')[-1], 'content': []}

                    # Tables
                    tbl = graphicData.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}tbl')
                    if tbl is not None:
                        frame_info['type'] = 'table'
                        rows = []
                        for tr in tbl.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr'):
                            row_cells = []
                            for tc in tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc'):
                                cell_text = get_text_from_element(tc).strip()
                                row_cells.append(cell_text)
                            if any(row_cells):
                                rows.append(row_cells)
                        frame_info['content'] = rows

                    # All other text in graphic frames
                    all_text = get_paragraphs_from_element(graphicData)
                    if all_text and frame_info['type'] != 'table':
                        frame_info['content'] = all_text

                    if frame_info['content']:
                        slide_data['shapes'].append({'graphic_frame': frame_info})

        # Process pictures (images)
        for pic in tree_elem.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}pic'):
            nvPicPr = pic.find('{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr')
            blipFill = pic.find('{http://schemas.openxmlformats.org/presentationml/2006/main}blipFill')
            img_info = {}
            if nvPicPr is not None:
                cNvPr = nvPicPr.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr')
                if cNvPr is not None:
                    img_info['name'] = cNvPr.get('name', '')
                    img_info['descr'] = cNvPr.get('descr', '')
            if blipFill is not None:
                blip = blipFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if blip is not None:
                    embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
                    if embed and embed in rels:
                        img_info['file'] = rels[embed]['target']
            slide_data['images'].append(img_info)

    return slide_data

def get_slide_notes(zf, notes_path):
    """Extract notes text from a slide notes file."""
    try:
        with zf.open(notes_path) as f:
            tree = ET.parse(f)
        root = tree.getroot()
        paras = get_paragraphs_from_element(root)
        # Filter out slide number placeholder text (usually just a number)
        notes = [p for p in paras if not p.strip().isdigit()]
        return '\n'.join(notes).strip() if notes else None
    except KeyError:
        return None

def get_presentation_metadata(zf):
    """Extract core metadata from the presentation."""
    metadata = {}
    try:
        with zf.open('docProps/core.xml') as f:
            tree = ET.parse(f)
        root = tree.getroot()
        dc_ns = 'http://purl.org/dc/elements/1.1/'
        cp_ns = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        dcterms_ns = 'http://purl.org/dc/terms/'
        for tag, key in [
            (f'{{{dc_ns}}}title', 'title'),
            (f'{{{dc_ns}}}creator', 'author'),
            (f'{{{dc_ns}}}subject', 'subject'),
            (f'{{{dc_ns}}}description', 'description'),
            (f'{{{cp_ns}}}lastModifiedBy', 'last_modified_by'),
            (f'{{{dcterms_ns}}}created', 'created'),
            (f'{{{dcterms_ns}}}modified', 'modified'),
        ]:
            el = root.find(f'.//{tag}')
            if el is not None and el.text:
                metadata[key] = el.text
    except KeyError:
        pass
    return metadata

def get_slide_order(zf):
    """Get the ordered list of slide rIds from presentation.xml."""
    with zf.open('ppt/presentation.xml') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    sldIdLst = root.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
    rids = []
    if sldIdLst is not None:
        for sldId in sldIdLst.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}sldId'):
            rid = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rid:
                rids.append(rid)
    return rids

def get_presentation_rels(zf):
    """Map rIds to slide paths from ppt/_rels/presentation.xml.rels."""
    rels = {}
    try:
        with zf.open('ppt/_rels/presentation.xml.rels') as f:
            tree = ET.parse(f)
        for rel in tree.iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rels[rel.get('Id')] = rel.get('Target', '')
    except KeyError:
        pass
    return rels

def main():
    print("=" * 70)
    print("YAMAHA POWERPOINT EXTRACTION REPORT")
    print("=" * 70)
    print(f"File: {PPTX_PATH}\n")

    with zipfile.ZipFile(PPTX_PATH, 'r') as zf:
        all_files = zf.namelist()

        # Metadata
        metadata = get_presentation_metadata(zf)
        if metadata:
            print("PRESENTATION METADATA")
            print("-" * 40)
            for k, v in metadata.items():
                print(f"  {k.replace('_', ' ').title()}: {v}")
            print()

        # Slide ordering
        rids = get_slide_order(zf)
        pres_rels = get_presentation_rels(zf)

        # Build ordered slide paths
        ordered_slides = []
        for rid in rids:
            target = pres_rels.get(rid, '')
            if target:
                # target is relative to ppt/ folder
                slide_path = 'ppt/' + target.lstrip('/')
                ordered_slides.append(slide_path)

        # Fallback: discover slides if ordering failed
        if not ordered_slides:
            ordered_slides = sorted(
                [f for f in all_files if re.match(r'ppt/slides/slide\d+\.xml$', f)],
                key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group())
            )

        print(f"TOTAL SLIDES: {len(ordered_slides)}")
        print("=" * 70)
        print()

        for idx, slide_path in enumerate(ordered_slides, 1):
            slide_file = slide_path.split('/')[-1]  # e.g. slide3.xml
            slide_name_no_ext = slide_file.replace('.xml', '')
            rels_path = f'ppt/slides/_rels/{slide_file}.rels'
            notes_path = f'ppt/notesSlides/notesSlide{idx}.xml'

            print(f"SLIDE {idx}")
            print("-" * 50)

            try:
                slide_data = parse_slide(zf, slide_path, rels_path)
            except Exception as e:
                print(f"  [Error parsing slide: {e}]")
                print()
                continue

            # Title
            if slide_data['title']:
                print(f"  TITLE: {slide_data['title']}")
            else:
                print("  TITLE: (no title detected)")

            # Shapes / text content
            if slide_data['shapes']:
                print("  CONTENT:")
                for shape in slide_data['shapes']:
                    if 'graphic_frame' in shape:
                        gf = shape['graphic_frame']
                        print(f"    [{gf['type'].upper()}]")
                        if gf['type'] == 'table':
                            for row in gf['content']:
                                print(f"      | {' | '.join(row)} |")
                        else:
                            for line in gf['content']:
                                print(f"      {line}")
                    else:
                        ph = shape.get('ph_type', '')
                        name = shape.get('name', '')
                        text = shape.get('text', '')
                        label = ''
                        if ph in ('title', 'ctrTitle'):
                            label = '[TITLE] '
                        elif ph in ('subTitle',):
                            label = '[SUBTITLE] '
                        elif ph in ('body', None, 'obj'):
                            label = '[BODY] '
                        else:
                            label = f'[{ph.upper() if ph else name}] '

                        # Skip re-printing the title we already printed
                        if ph in ('title', 'ctrTitle'):
                            continue

                        lines = text.split('\n')
                        print(f"    {label}{lines[0]}")
                        for line in lines[1:]:
                            print(f"          {line}")

            # Images
            imgs = [i for i in slide_data['images'] if i]
            if imgs:
                print("  IMAGES:")
                for img in imgs:
                    parts = []
                    if img.get('name'):
                        parts.append(f"name='{img['name']}'")
                    if img.get('descr'):
                        parts.append(f"description='{img['descr']}'")
                    if img.get('file'):
                        parts.append(f"file='{img['file']}'")
                    print(f"    - {', '.join(parts) if parts else '(unnamed image)'}")

            # Notes
            notes = get_slide_notes(zf, notes_path)
            if notes:
                print("  SPEAKER NOTES:")
                for line in notes.split('\n'):
                    print(f"    {line}")

            print()

        # Summary of embedded media
        media_files = [f for f in all_files if f.startswith('ppt/media/')]
        if media_files:
            print("=" * 70)
            print(f"EMBEDDED MEDIA FILES ({len(media_files)} total):")
            for mf in sorted(media_files):
                print(f"  {mf}")
            print()

    print("=" * 70)
    print("END OF EXTRACTION REPORT")
    print("=" * 70)

if __name__ == '__main__':
    main()
