#!/usr/bin/env python3
"""
Enhance Yamaha.pptx:
- Add professional slide transitions to all slides
- Enhance existing slides with gradient overlays and decorative elements
- Add 6 new content slides (Piano, Digital, Global, Education, Sustainability, Summary)
- Move bibliography to position 15
Total: 15 slides
"""

import zipfile
import os
import shutil
import re

SRC = '/home/user/powerpoints-yamaha/Yamaha.pptx'
DST = '/home/user/powerpoints-yamaha/Yamaha_Enhanced.pptx'
WORK = '/tmp/pptx_enhance_work'

# ── Helper ──────────────────────────────────────────────────────────────────

def rpr(lang="hu-HU", sz=None, bold=False, color_hex=None, typeface="Times New Roman"):
    """Build a run-properties XML snippet."""
    attrs = f'lang="{lang}" dirty="0"'
    if sz:
        attrs += f' sz="{sz}"'
    if bold:
        attrs += ' b="1"'
    fill = ''
    if color_hex:
        fill = f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
    else:
        fill = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="85000"/><a:lumOff val="15000"/></a:schemeClr></a:solidFill>'
    return (f'<a:rPr {attrs}>'
            f'{fill}'
            f'<a:latin typeface="{typeface}" panose="02020603050405020304" pitchFamily="18" charset="0"/>'
            f'<a:cs typeface="{typeface}" panose="02020603050405020304" pitchFamily="18" charset="0"/>'
            f'</a:rPr>')

def run(text, lang="hu-HU", sz=None, bold=False, color_hex=None):
    return f'<a:r>{rpr(lang, sz, bold, color_hex)}<a:t>{text}</a:t></a:r>'

def para(runs_xml, align="l", bullet=True, indent=None):
    """Wrap runs in a paragraph."""
    pPr = f'<a:pPr algn="{align}"'
    if indent is not None:
        pPr += f' marL="{indent}" indent="0"'
    if not bullet:
        pPr += '><a:buNone/></a:pPr>'
    else:
        pPr += '/>'
    return f'<a:p>{pPr}{runs_xml}</a:p>'

def title_sp(title_text, x, y, cx, cy, sz=2800):
    """Title shape."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Cím 1"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="hu-HU" sz="{sz}" dirty="0">
          <a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill>
          <a:latin typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/>
          <a:cs typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/>
        </a:rPr>
        <a:t>{title_text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''

def body_sp(paragraphs_xml, x, y, cx, cy, sp_id=3):
    """Content body shape."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="Tartalom helye {sp_id}"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr><a:noAutofit/></a:bodyPr>
    <a:lstStyle/>
    {paragraphs_xml}
  </p:txBody>
</p:sp>'''

def pic_sp(rid, x, y, cx, cy, sp_id=4, soft_edge=True):
    """Picture shape (rectangular)."""
    effect = '<a:effectLst><a:softEdge rad="76200"/></a:effectLst>' if soft_edge else '<a:effectLst/>'
    return f'''<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{sp_id}" name="Kép {sp_id}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    {effect}
  </p:spPr>
</p:pic>'''

def accent_bar(x, y, cx, cy, color_hex="156082", alpha=80000, sp_id=20):
    """Decorative colored bar."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="AccentBar {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{color_hex}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

def gradient_bg_rect(sp_id=30):
    """Full-slide gradient background rectangle."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="GradBG {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0">
          <a:srgbClr val="0E2841"><a:alpha val="18000"/></a:srgbClr>
        </a:gs>
        <a:gs pos="100000">
          <a:srgbClr val="FFFFFF"><a:alpha val="5000"/></a:srgbClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="0"/>
    </a:gradFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

def separator_line(x, y, cx, color_hex="156082", sp_id=25):
    """Horizontal separator line."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="SepLine {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="6350"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

def shadow_text_box(text_xml, x, y, cx, cy, sp_id=40):
    """Text box with shadow effect."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="ShadowBox {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="85000"/></a:srgbClr></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="156082"><a:alpha val="40000"/></a:srgbClr></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
        <a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody>
    <a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720"><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    {text_xml}
  </p:txBody>
</p:sp>'''

def make_slide(spTree_content, slide_id_suffix=1):
    """Wrap shape tree content in a full slide XML."""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
  <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
  <p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {spTree_content}
  </p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def make_rels(layout_num=2, image_rids=None, hyperlink_rids=None):
    """Build slide .rels file content."""
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{layout_num}.xml"/>',
    ]
    if image_rids:
        for rid, target in image_rids.items():
            lines.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{target}"/>'),
    if hyperlink_rids:
        for rid, url in hyperlink_rids.items():
            lines.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/>'),
    lines.append('</Relationships>')
    return '\n'.join(lines)

# ── Background overlay (used on all new slides) ─────────────────────────────

BG_RECTS = '''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="18" name="BgFill"/>
    <p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="19" name="BgOverlay"/>
    <p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="82766A"><a:alpha val="15000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>'''

# ── Transitions for all 15 slides ────────────────────────────────────────────

TRANSITIONS = [
    # Slide 1: Fade (cinematic opening)
    '<p:transition spd="slow"><p:fade/></p:transition>',
    # Slide 2: Push left
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
    # Slide 3: Cover left
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    # Slide 4: Zoom in (philosophy)
    '<p:transition spd="slow"><p:zoom dir="in"/></p:transition>',
    # Slide 5: Wipe left
    '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
    # Slide 6: Split horizontal out
    '<p:transition spd="med"><p:split dir="out" orient="horz"/></p:transition>',
    # Slide 7: Cover left
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    # Slide 8: Fade (future)
    '<p:transition spd="slow"><p:fade/></p:transition>',
    # Slide 9 (new: Piano): Push left
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
    # Slide 10 (new: Digital): Wipe left
    '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
    # Slide 11 (new: Global): Zoom in
    '<p:transition spd="slow"><p:zoom dir="in"/></p:transition>',
    # Slide 12 (new: Education): Cover left
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    # Slide 13 (new: Sustainability): Split
    '<p:transition spd="med"><p:split dir="out" orient="horz"/></p:transition>',
    # Slide 14 (new: Summary): Fade
    '<p:transition spd="slow"><p:fade/></p:transition>',
    # Slide 15 (bibliography): Push
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
]

# ── Build new slide content ──────────────────────────────────────────────────

def build_slide_piano():
    """Slide 9: Yamaha zongorák – A klasszikus és modern zongora"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "E97132", 100000, sp_id=32) +
        separator_line(358817, 1950000, 5500000, "156082", sp_id=33) +
        title_sp("Yamaha zongorák", 358817, 500000, 5500000, 1200000, sz=3200) +
        body_sp(
            para(run("Yamaha 1900 óta gyárt zongorát", sz=1800, bold=True) +
                 run("", sz=600)) +
            para(run("• CF Series: koncertzongorák – a Carnegue Hallban és a Scala Milanóban", sz=1600)) +
            para(run("• U sorozat: otthoni felállású zongorák, precíz mechanikával", sz=1600)) +
            para(run("• AvantGrand: akusztikus + digitális hibrid technológia", sz=1600)) +
            para(run("• Disklavier: hálózati lejátszás és rögzítés valós időben", sz=1600)) +
            para(run("• CFX: Yamaha zászlóshajó koncertflügel, 9 láb hosszú", sz=1600)),
            358817, 2050000, 5800000, 4600000
        ) +
        pic_sp("rId2", 6250000, 400000, 5700000, 2900000, sp_id=10, soft_edge=True) +
        pic_sp("rId3", 6250000, 3450000, 5700000, 3200000, sp_id=11, soft_edge=True)
    )
    return make_slide(content)

def build_slide_digital():
    """Slide 10: Digitális hangszerek és technológia"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "E97132", 100000, sp_id=32) +
        separator_line(358817, 1950000, 5500000, "E97132", sp_id=33) +
        title_sp("Digitális hangszerek és technológia", 358817, 500000, 5500000, 1200000, sz=2800) +
        body_sp(
            para(run("Yamaha digitális innováció", sz=1800, bold=True)) +
            para(run("• Clavinova sorozat: akusztikus zongorát utánzó kalapácsmechanika", sz=1600)) +
            para(run("• MONTAGE / MODX szintetizátorplatform – élő zenészek kedvence", sz=1600)) +
            para(run("• Reface: kompakt, vintage hangzású szintetizátorok", sz=1600)) +
            para(run("• YC sorozat: professzionális stage orgona és EP hangok", sz=1600)) +
            para(run("• A digitális fejlesztés nem helyettesíti, hanem kiegészíti az akusztikust", sz=1600)),
            358817, 2050000, 5800000, 4600000
        ) +
        pic_sp("rId2", 6250000, 400000, 5700000, 3000000, sp_id=10, soft_edge=True) +
        pic_sp("rId3", 6250000, 3550000, 5700000, 3100000, sp_id=11, soft_edge=True)
    )
    return make_slide(content)

def build_slide_global():
    """Slide 11: Yamaha globális jelenlét"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "196B24", 100000, sp_id=32) +
        separator_line(358817, 1950000, 5500000, "196B24", sp_id=33) +
        title_sp("Yamaha globális jelenléte", 358817, 500000, 5500000, 1200000, sz=2900) +
        body_sp(
            para(run("Világ egyik vezető hangszergyártója", sz=1800, bold=True)) +
            para(run("• Alapítás: 1897, Hamamatsu, Japán", sz=1600)) +
            para(run("• Működés 80+ országban, több mint 40 000 alkalmazott", sz=1600)) +
            para(run("• Éves árbevétel: ~500 milliárd JPY (kb. 3,5 milliárd USD)", sz=1600)) +
            para(run("• Gyárüzemek: Japán, Indonézia, Kína, USA, Mexikó", sz=1600)) +
            para(run("• A Yamaha zenéktől a motorkerékpárokig terjed a termékpaletta", sz=1600)) +
            para(run("  (Yamaha Motor Co. Ltd. – 1955-től különvált leányvállalat)", sz=1400, color_hex="555555")),
            358817, 2050000, 5800000, 4600000
        ) +
        pic_sp("rId2", 6250000, 400000, 5700000, 6250000, sp_id=10, soft_edge=True)
    )
    return make_slide(content)

def build_slide_education():
    """Slide 12: Yamaha zenei oktatás"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "0F9ED5", 100000, sp_id=32) +
        separator_line(358817, 1950000, 5500000, "0F9ED5", sp_id=33) +
        title_sp("Yamaha zenei oktatás", 358817, 500000, 5500000, 1200000, sz=3000) +
        body_sp(
            para(run("Yamaha Music School – globális oktatási hálózat", sz=1800, bold=True)) +
            para(run("• 1954 óta működő zenepedagógiai program", sz=1600)) +
            para(run("• 40+ ország, több mint 500 000 tanuló évente", sz=1600)) +
            para(run("• Gyermekektől felnőttekig: zongora, gitár, ének, fúvósok", sz=1600)) +
            para(run("• Junior Original Concert: gyermekzeneszerzői verseny (1972 óta)", sz=1600)) +
            para(run("• Yamaha ösztöndíjprogram tehetséges fiatal zenészeknek", sz=1600)),
            358817, 2050000, 5800000, 4600000
        ) +
        pic_sp("rId2", 6250000, 400000, 5700000, 3100000, sp_id=10, soft_edge=True) +
        pic_sp("rId3", 6250000, 3600000, 5700000, 3050000, sp_id=11, soft_edge=True)
    )
    return make_slide(content)

def build_slide_sustainability():
    """Slide 13: Fenntarthatóság és értékek"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "196B24", 100000, sp_id=32) +
        separator_line(358817, 1950000, 5500000, "196B24", sp_id=33) +
        title_sp("Fenntarthatóság és Yamaha értékek", 358817, 500000, 5500000, 1200000, sz=2600) +
        body_sp(
            para(run('\u201eHangokkal gazdag\u00edtani az emberek \u00e9let\u00e9t\u201d', sz=1800, bold=True, color_hex="156082")) +
            para(run("• Felelős erdészet: FSC-tanúsított faanyag hangszerekhez", sz=1600)) +
            para(run("• CO₂-semlegesség 2050-re – Yamaha Sustainability Plan 2030", sz=1600)) +
            para(run("• Újrahasznosítható anyagok: rézhangszerek, billentyűzetek", sz=1600)) +
            para(run("• Szociális felelősségvállalás: zeneterápia, közösségi programok", sz=1600)) +
            para(run("• A minőség és a természet tisztelete egymást erősíti", sz=1600)),
            358817, 2050000, 5800000, 4600000
        ) +
        pic_sp("rId2", 6250000, 400000, 5700000, 3000000, sp_id=10, soft_edge=True) +
        pic_sp("rId3", 6250000, 3600000, 5700000, 3050000, sp_id=11, soft_edge=True)
    )
    return make_slide(content)

def build_slide_summary():
    """Slide 14: Összefoglalás"""
    content = (
        BG_RECTS +
        gradient_bg_rect(sp_id=30) +
        accent_bar(0, 0, 12192000, 190500, "0E2841", 90000, sp_id=31) +
        accent_bar(0, 190500, 12192000, 25400, "E97132", 100000, sp_id=32) +
        accent_bar(0, 6668000, 12192000, 190000, "156082", 70000, sp_id=35) +
        separator_line(358817, 1950000, 11434366, "E97132", sp_id=33) +
        title_sp("Összefoglalás", 358817, 500000, 11434366, 1200000, sz=3600) +
        body_sp(
            para(run("A Yamaha Corporation 130+ éves mértékadó örökség", sz=1900, bold=True)) +
            para(run("")) +
            para(run("✦  1887-től máig: orgonától a koncertflügelig, a fuvolától a szintetizátorig", sz=1700)) +
            para(run("✦  Kézművesség + high-tech = páratlan hangminőség", sz=1700)) +
            para(run("✦  Globális jelenlét 80+ országban, világ egyik vezető márkája", sz=1700)) +
            para(run("✦  Oktatási programok 500 000+ tanulónak évente", sz=1700)) +
            para(run("✦  Fenntarthatósági elköteleződés a jövő generációinak", sz=1700)) +
            para(run("✦  Folyamatos innováció digitális és akusztikus területen egyaránt", sz=1700)) +
            para(run("")) +
            para(run('\u201eYamaha \u2013 Making Waves\u201d', sz=1800, bold=True, color_hex="156082")),
            358817, 2050000, 7500000, 4600000
        ) +
        pic_sp("rId2", 8000000, 2200000, 4000000, 3400000, sp_id=10, soft_edge=True)
    )
    return make_slide(content)

# ── New slides data (file, rels) ─────────────────────────────────────────────

NEW_SLIDES = [
    {
        'filename': 'slide10.xml',
        'builder': build_slide_piano,
        'rels': make_rels(2, {'rId2': 'image4.png', 'rId3': 'image5.png'}),
        'rId': 'rId22',
        'sld_id': 270,
    },
    {
        'filename': 'slide11.xml',
        'builder': build_slide_digital,
        'rels': make_rels(2, {'rId2': 'image14.png', 'rId3': 'image11.png'}),
        'rId': 'rId23',
        'sld_id': 271,
    },
    {
        'filename': 'slide12.xml',
        'builder': build_slide_global,
        'rels': make_rels(2, {'rId2': 'image1.png'}),
        'rId': 'rId24',
        'sld_id': 272,
    },
    {
        'filename': 'slide13.xml',
        'builder': build_slide_education,
        'rels': make_rels(2, {'rId2': 'image2.png', 'rId3': 'image3.png'}),
        'rId': 'rId25',
        'sld_id': 273,
    },
    {
        'filename': 'slide14.xml',
        'builder': build_slide_sustainability,
        'rels': make_rels(2, {'rId2': 'image7.png', 'rId3': 'image13.png'}),
        'rId': 'rId26',
        'sld_id': 274,
    },
    {
        'filename': 'slide15.xml',
        'builder': build_slide_summary,
        'rels': make_rels(2, {'rId2': 'image15.png'}),
        'rId': 'rId27',
        'sld_id': 275,
    },
]

# ── Main processing ──────────────────────────────────────────────────────────

def add_transition(xml_content, transition_xml):
    """Inject transition before </p:sld>."""
    # Remove existing transition if any
    xml_content = re.sub(r'<p:transition[^>]*/>', '', xml_content)
    xml_content = re.sub(r'<p:transition[^>]*>.*?</p:transition>', '', xml_content, flags=re.DOTALL)
    # Insert before closing tag
    xml_content = xml_content.replace('</p:sld>', f'\n{transition_xml}\n</p:sld>')
    return xml_content

def add_gradient_to_existing_slide(xml_content, slide_num):
    """
    Enhance existing slides:
    - Add subtle gradient overlay shape after the first background rect
    - Add decorative top/bottom accent bars
    The gradient adds depth without disturbing existing layout.
    """
    # Build enhancement shapes (inserted BEFORE text/images but AFTER bg rects)
    # Top accent bar (dark navy, semi-transparent)
    top_bar = accent_bar(0, 0, 12192000, 152400, "0E2841", 75000, sp_id=90+slide_num)
    # Bottom thin accent line
    bottom_line = separator_line(0, 6705600, 12192000, "156082", sp_id=95+slide_num)
    # Subtle top gradient
    grad = f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{80+slide_num}" name="TopGrad{slide_num}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="1524000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0">
          <a:srgbClr val="0E2841"><a:alpha val="30000"/></a:srgbClr>
        </a:gs>
        <a:gs pos="100000">
          <a:srgbClr val="0E2841"><a:alpha val="0"/></a:srgbClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="0"/>
    </a:gradFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

    # Insert after </p:grpSpPr> (beginning of spTree content)
    insert_point = '</p:grpSpPr>'
    if insert_point in xml_content:
        xml_content = xml_content.replace(
            insert_point,
            insert_point + '\n' + top_bar + '\n' + bottom_line + '\n' + grad,
            1
        )
    return xml_content

def process():
    # 1. Extract original PPTX
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    with zipfile.ZipFile(SRC, 'r') as z:
        z.extractall(WORK)
    print("Extracted PPTX.")

    slides_dir = os.path.join(WORK, 'ppt', 'slides')
    rels_dir = os.path.join(slides_dir, '_rels')

    # 2. Enhance existing slides 1-8 (add transitions + gradient overlays)
    #    Leave slide9 (bibliography) – only add transition
    for i in range(1, 9):
        path = os.path.join(slides_dir, f'slide{i}.xml')
        with open(path, 'r', encoding='utf-8') as f:
            xml = f.read()
        xml = add_gradient_to_existing_slide(xml, i)
        xml = add_transition(xml, TRANSITIONS[i-1])
        with open(path, 'w', encoding='utf-8') as f:
            f.write(xml)
        print(f"Enhanced slide {i}.")

    # 3. Enhance slide 9 (bibliography) – only add transition (keep at position 15 in order)
    path9 = os.path.join(slides_dir, 'slide9.xml')
    with open(path9, 'r', encoding='utf-8') as f:
        xml = f.read()
    xml = add_transition(xml, TRANSITIONS[14])  # slide 15 transition
    with open(path9, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("Enhanced slide 9 (bibliography).")

    # 4. Create new slides 10-15
    for slide_info in NEW_SLIDES:
        fname = slide_info['filename']
        xml_content = slide_info['builder']()
        # Find the transition index (slides 9-14 → index 8-13)
        slide_num = int(fname.replace('slide', '').replace('.xml', ''))
        trans_idx = slide_num - 1  # 0-based
        xml_content = add_transition(xml_content, TRANSITIONS[trans_idx])

        # Write slide XML
        with open(os.path.join(slides_dir, fname), 'w', encoding='utf-8') as f:
            f.write(xml_content)

        # Write rels
        with open(os.path.join(rels_dir, f'{fname}.rels'), 'w', encoding='utf-8') as f:
            f.write(slide_info['rels'])

        print(f"Created {fname}.")

    # 5. Update presentation.xml – add new slide IDs in correct order
    pres_path = os.path.join(WORK, 'ppt', 'presentation.xml')
    with open(pres_path, 'r', encoding='utf-8') as f:
        pres_xml = f.read()

    # Find existing sldIdLst closing tag and insert before bibliography (rId10)
    # Current last entry: <p:sldId id="266" r:id="rId10"/>
    # We need to insert 6 new slides between rId9 (slide8) and rId10 (bibliography)
    new_slide_ids = ''
    for slide_info in NEW_SLIDES:
        new_slide_ids += f'<p:sldId id="{slide_info["sld_id"]}" r:id="{slide_info["rId"]}"/>'

    # Insert after slide8 reference (rId9) and before slide9 (rId10 = bibliography)
    pres_xml = pres_xml.replace(
        '<p:sldId id="265" r:id="rId9"/>',
        '<p:sldId id="265" r:id="rId9"/>' + new_slide_ids
    )
    with open(pres_path, 'w', encoding='utf-8') as f:
        f.write(pres_xml)
    print("Updated presentation.xml.")

    # 6. Update presentation.xml.rels – add new slide relationships
    pres_rels_path = os.path.join(WORK, 'ppt', '_rels', 'presentation.xml.rels')
    with open(pres_rels_path, 'r', encoding='utf-8') as f:
        pres_rels = f.read()

    new_rels = ''
    for slide_info in NEW_SLIDES:
        fname = slide_info['filename']
        rid = slide_info['rId']
        new_rels += f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{fname}"/>'

    pres_rels = pres_rels.replace('</Relationships>', new_rels + '</Relationships>')
    with open(pres_rels_path, 'w', encoding='utf-8') as f:
        f.write(pres_rels)
    print("Updated presentation.xml.rels.")

    # 7. Update [Content_Types].xml – register new slide files
    ct_path = os.path.join(WORK, '[Content_Types].xml')
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct_xml = f.read()

    new_overrides = ''
    for slide_info in NEW_SLIDES:
        fname = slide_info['filename']
        new_overrides += f'<Override PartName="/ppt/slides/{fname}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'

    ct_xml = ct_xml.replace('</Types>', new_overrides + '</Types>')
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct_xml)
    print("Updated [Content_Types].xml.")

    # 8. Pack back into PPTX
    if os.path.exists(DST):
        os.remove(DST)
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(WORK):
            for file in files:
                fpath = os.path.join(root, file)
                arcname = os.path.relpath(fpath, WORK)
                zout.write(fpath, arcname)
    print(f"\nDone! Enhanced PPTX saved to: {DST}")
    print(f"Total slides: 9 existing + 6 new = 15 slides")

if __name__ == '__main__':
    process()
