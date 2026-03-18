#!/usr/bin/env python3
"""
Enhance Yamaha.pptx:
- Egységes háttér minden diára (sötétkék + gradient overlay)
- Fotóismétlés nélkül: az új diák (10-15) dekoratív alakzatokat használnak fotók helyett
- 6 új tartalom dia + átmenetek az összes diára
Total: 15 dia
"""

import zipfile
import os
import shutil
import re

SRC = '/home/user/powerpoints-yamaha/Yamaha.pptx'
DST = '/home/user/powerpoints-yamaha/Yamaha_Enhanced.pptx'
WORK = '/tmp/pptx_enhance_work'

# Egységes háttérszínek
BG_DARK   = "0E2841"   # sötétkék alap
BG_MID    = "1A3A5C"   # közép kék
ACCENT_OR = "E97132"   # narancs akcent
ACCENT_BL = "156082"   # kék akcent
ACCENT_GR = "196B24"   # zöld akcent
ACCENT_LB = "0F9ED5"   # világoskék akcent

# ── Helper ──────────────────────────────────────────────────────────────────

def rpr(lang="hu-HU", sz=None, bold=False, color_hex=None, typeface="Times New Roman"):
    attrs = f'lang="{lang}" dirty="0"'
    if sz:
        attrs += f' sz="{sz}"'
    if bold:
        attrs += ' b="1"'
    fill = ''
    if color_hex:
        fill = f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
    else:
        fill = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    return (f'<a:rPr {attrs}>'
            f'{fill}'
            f'<a:latin typeface="{typeface}" panose="02020603050405020304" pitchFamily="18" charset="0"/>'
            f'<a:cs typeface="{typeface}" panose="02020603050405020304" pitchFamily="18" charset="0"/>'
            f'</a:rPr>')

def run(text, lang="hu-HU", sz=None, bold=False, color_hex=None):
    return f'<a:r>{rpr(lang, sz, bold, color_hex)}<a:t>{text}</a:t></a:r>'

def para(runs_xml, align="l", bullet=True, indent=None):
    pPr = f'<a:pPr algn="{align}"'
    if indent is not None:
        pPr += f' marL="{indent}" indent="0"'
    if not bullet:
        pPr += '><a:buNone/></a:pPr>'
    else:
        pPr += '/>'
    return f'<a:p>{pPr}{runs_xml}</a:p>'

def title_sp(title_text, x, y, cx, cy, sz=2800, color_hex="FFFFFF"):
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
        <a:rPr lang="hu-HU" sz="{sz}" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
          <a:latin typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/>
          <a:cs typeface="Times New Roman" panose="02020603050405020304" pitchFamily="18" charset="0"/>
        </a:rPr>
        <a:t>{title_text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''

def body_sp(paragraphs_xml, x, y, cx, cy, sp_id=3):
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

def accent_bar(x, y, cx, cy, color_hex=ACCENT_BL, alpha=100000, sp_id=20):
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

def separator_line(x, y, cx, color_hex=ACCENT_BL, sp_id=25):
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

def info_box(text_xml, x, y, cx, cy, bg_color=ACCENT_BL, alpha=25000, sp_id=40):
    """Dekoratív szövegdoboz fotó helyett."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="InfoBox {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{bg_color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="{bg_color}"><a:alpha val="60000"/></a:srgbClr></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr lIns="182880" rIns="182880" tIns="91440" bIns="91440"><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    {text_xml}
  </p:txBody>
</p:sp>'''

def stat_block(number, label, x, y, cx, cy, accent_color=ACCENT_OR, sp_id=50):
    """Statisztika blokk – fotó helyett vizuális elem."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="StatBlock {sp_id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="{BG_MID}"><a:alpha val="80000"/></a:srgbClr></a:solidFill>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr lIns="91440" rIns="91440" tIns="91440" bIns="45720" anchor="ctr"><a:normAutofit/></a:bodyPr>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/>
      <a:r><a:rPr lang="hu-HU" sz="3600" b="1" dirty="0">
        <a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>
        <a:latin typeface="Times New Roman" pitchFamily="18" charset="0"/>
      </a:rPr><a:t>{number}</a:t></a:r>
    </a:p>
    <a:p><a:pPr algn="ctr"/>
      <a:r><a:rPr lang="hu-HU" sz="1200" dirty="0">
        <a:solidFill><a:srgbClr val="CCDDEE"/></a:solidFill>
        <a:latin typeface="Times New Roman" pitchFamily="18" charset="0"/>
      </a:rPr><a:t>{label}</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>'''

# ── Egységes háttér minden új diához ─────────────────────────────────────────

def unified_bg(accent_color=ACCENT_OR):
    """Egységes sötétkék háttér + akcent sáv minden új diához."""
    return f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="18" name="UnifiedBG"/>
    <p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{BG_DARK}"/></a:gs>
        <a:gs pos="60000"><a:srgbClr val="{BG_MID}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{BG_DARK}"/></a:gs>
      </a:gsLst>
      <a:lin ang="10800000" scaled="0"/>
    </a:gradFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="19" name="TopBar"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="228600"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="17" name="BottomBar"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="6629400"/><a:ext cx="12192000" cy="228600"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{accent_color}"><a:alpha val="60000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

# ── Transitions ────────────────────────────────────────────────────────────

TRANSITIONS = [
    '<p:transition spd="slow"><p:fade/></p:transition>',
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    '<p:transition spd="slow"><p:zoom dir="in"/></p:transition>',
    '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
    '<p:transition spd="med"><p:split dir="out" orient="horz"/></p:transition>',
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    '<p:transition spd="slow"><p:fade/></p:transition>',
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
    '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
    '<p:transition spd="slow"><p:zoom dir="in"/></p:transition>',
    '<p:transition spd="med"><p:cover dir="l"/></p:transition>',
    '<p:transition spd="med"><p:split dir="out" orient="horz"/></p:transition>',
    '<p:transition spd="slow"><p:fade/></p:transition>',
    '<p:transition spd="med"><p:push dir="l"/></p:transition>',
]

def make_slide(spTree_content):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
  <p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {spTree_content}
  </p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def make_rels(layout_num=2, image_rids=None, hyperlink_rids=None):
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

# ── Új diák – fotók nélkül, dekoratív elemekkel ──────────────────────────────

def build_slide_piano():
    """Slide 9: Yamaha zongorák"""
    content = (
        unified_bg(ACCENT_OR) +
        separator_line(457200, 1980000, 5400000, ACCENT_OR, sp_id=33) +
        title_sp("Yamaha zongorák", 457200, 350000, 11277600, 1400000, sz=3600) +
        body_sp(
            para(run("A klasszikus és modern zongora mestere", sz=1800, bold=True, color_hex=ACCENT_OR)) +
            para(run("• CF Series: koncertzongorák a Carnegie Hallban és a Scala Milanóban", sz=1600)) +
            para(run("• U sorozat: otthoni felállású zongorák, precíz mechanikával", sz=1600)) +
            para(run("• AvantGrand: akusztikus + digitális hibrid technológia", sz=1600)) +
            para(run("• Disklavier: hálózati lejátszás és rögzítés valós időben", sz=1600)) +
            para(run("• CFX: Yamaha zászlóshajó koncertflügel, 9 láb hosszú", sz=1600)),
            457200, 2050000, 5800000, 4300000
        ) +
        stat_block("1900", "óta gyárt zongorát", 6500000, 450000, 2600000, 1400000, ACCENT_OR, sp_id=51) +
        stat_block("CFX", "zászlóshajó flügel", 9300000, 450000, 2600000, 1400000, ACCENT_BL, sp_id=52) +
        stat_block("CF Series", "koncerttermi\nstandard", 6500000, 2050000, 2600000, 1500000, ACCENT_OR, sp_id=53) +
        stat_block("Disklavier", "hálózati MIDI\nrendszer", 9300000, 2050000, 2600000, 1500000, ACCENT_BL, sp_id=54) +
        info_box(
            para(run("\"A Yamaha CFX zongora az élő zene\ntökéletes megtestesítője – átlátszó hangzás,\nkifinomult mechanika, páratlan dinamika.\"",
                     sz=1400, color_hex="CCDDEE"), align="ctr", bullet=False),
            6500000, 3750000, 5400000, 2800000, ACCENT_BL, 20000, sp_id=60
        )
    )
    return make_slide(content)

def build_slide_digital():
    """Slide 10: Digitális hangszerek"""
    content = (
        unified_bg(ACCENT_BL) +
        separator_line(457200, 1980000, 5400000, ACCENT_BL, sp_id=33) +
        title_sp("Digitális hangszerek és technológia", 457200, 350000, 11277600, 1400000, sz=3000) +
        body_sp(
            para(run("Yamaha digitális innováció – az akusztikus élmény digitálisan", sz=1800, bold=True, color_hex=ACCENT_LB)) +
            para(run("• Clavinova sorozat: akusztikus zongorát utánzó kalapácsmechanika", sz=1600)) +
            para(run("• MONTAGE / MODX szintetizátorplatform – élő zenészek kedvence", sz=1600)) +
            para(run("• Reface: kompakt, vintage hangzású szintetizátorok", sz=1600)) +
            para(run("• YC sorozat: professzionális stage orgona és EP hangok", sz=1600)) +
            para(run("• A digitális fejlesztés nem helyettesíti, hanem kiegészíti az akusztikust", sz=1600)),
            457200, 2050000, 5800000, 4300000
        ) +
        stat_block("Clavinova", "otthoni digitális\nzongora-élmény", 6500000, 450000, 2600000, 1500000, ACCENT_LB, sp_id=51) +
        stat_block("MONTAGE", "professzionális\nszintetizátor", 9300000, 450000, 2600000, 1500000, ACCENT_OR, sp_id=52) +
        stat_block("Reface", "vintage kompakt\nszintetizátorok", 6500000, 2150000, 2600000, 1500000, ACCENT_LB, sp_id=53) +
        stat_block("YC sorozat", "stage orgona\nprofik számára", 9300000, 2150000, 2600000, 1500000, ACCENT_OR, sp_id=54) +
        info_box(
            para(run("\"A digitális és az akusztikus hangszerek\nnem versengenek egymással – a Yamaha\nmindkét világban otthon van.\"",
                     sz=1400, color_hex="CCDDEE"), align="ctr", bullet=False),
            6500000, 3850000, 5400000, 2500000, ACCENT_LB, 20000, sp_id=60
        )
    )
    return make_slide(content)

def build_slide_global():
    """Slide 11: Yamaha globális jelenlét"""
    content = (
        unified_bg(ACCENT_GR) +
        separator_line(457200, 1980000, 5400000, ACCENT_GR, sp_id=33) +
        title_sp("Yamaha globális jelenléte", 457200, 350000, 11277600, 1400000, sz=3200) +
        body_sp(
            para(run("A világ egyik vezető hangszergyártója", sz=1800, bold=True, color_hex=ACCENT_GR)) +
            para(run("• Alapítás: 1897, Hamamatsu, Japán – 130+ év tapasztalat", sz=1600)) +
            para(run("• Működés 80+ országban, több mint 40 000 alkalmazott", sz=1600)) +
            para(run("• Éves árbevétel: ~500 milliárd JPY (kb. 3,5 milliárd USD)", sz=1600)) +
            para(run("• Gyárüzemek: Japán, Indonézia, Kína, USA, Mexikó", sz=1600)) +
            para(run("• A hangszergyártástól a szórakoztatás-elektronikáig", sz=1600)),
            457200, 2050000, 5800000, 4300000
        ) +
        stat_block("80+", "ország ahol\njelen van", 6500000, 450000, 2600000, 1400000, ACCENT_GR, sp_id=51) +
        stat_block("40 000+", "alkalmazott\nvilágszerte", 9300000, 450000, 2600000, 1400000, ACCENT_OR, sp_id=52) +
        stat_block("130+", "év tapasztalat\n(1887 óta)", 6500000, 2050000, 2600000, 1400000, ACCENT_GR, sp_id=53) +
        stat_block("3,5 Mrd $", "éves árbevétel\n(~500 Mrd JPY)", 9300000, 2050000, 2600000, 1400000, ACCENT_OR, sp_id=54) +
        stat_block("500 000+", "zenetanuló\névente", 6500000, 3650000, 2600000, 1400000, ACCENT_LB, sp_id=55) +
        stat_block("#1", "hangszergyártó\nJapánban", 9300000, 3650000, 2600000, 1400000, ACCENT_LB, sp_id=56)
    )
    return make_slide(content)

def build_slide_education():
    """Slide 12: Yamaha zenei oktatás"""
    content = (
        unified_bg(ACCENT_LB) +
        separator_line(457200, 1980000, 5400000, ACCENT_LB, sp_id=33) +
        title_sp("Yamaha zenei oktatás", 457200, 350000, 11277600, 1400000, sz=3400) +
        body_sp(
            para(run("Yamaha Music School – globális oktatási hálózat", sz=1800, bold=True, color_hex=ACCENT_LB)) +
            para(run("• 1954 óta működő zenepedagógiai program", sz=1600)) +
            para(run("• 40+ ország, több mint 500 000 tanuló évente", sz=1600)) +
            para(run("• Gyermekektől felnőttekig: zongora, gitár, ének, fúvósok", sz=1600)) +
            para(run("• Junior Original Concert: gyermekzeneszerzői verseny (1972 óta)", sz=1600)) +
            para(run("• Yamaha ösztöndíjprogram tehetséges fiatal zenészeknek", sz=1600)),
            457200, 2050000, 5800000, 4300000
        ) +
        stat_block("1954", "az oktatási\nprogram kezdete", 6500000, 450000, 2600000, 1400000, ACCENT_LB, sp_id=51) +
        stat_block("500 000+", "tanuló\névente", 9300000, 450000, 2600000, 1400000, ACCENT_OR, sp_id=52) +
        stat_block("40+", "ország ahol\nműködik", 6500000, 2050000, 2600000, 1400000, ACCENT_LB, sp_id=53) +
        stat_block("1972", "Junior Original\nConcert indulása", 9300000, 2050000, 2600000, 1400000, ACCENT_OR, sp_id=54) +
        info_box(
            para(run("\"A zenei oktatás az emberi fejlődés\nalappillére – a Yamaha ezt vallja\n1954 óta szerte a világban.\"",
                     sz=1400, color_hex="CCDDEE"), align="ctr", bullet=False),
            6500000, 3650000, 5400000, 2700000, ACCENT_LB, 20000, sp_id=60
        )
    )
    return make_slide(content)

def build_slide_sustainability():
    """Slide 13: Fenntarthatóság"""
    content = (
        unified_bg(ACCENT_GR) +
        separator_line(457200, 1980000, 5400000, ACCENT_GR, sp_id=33) +
        title_sp("Fenntarthatóság és Yamaha értékek", 457200, 350000, 11277600, 1400000, sz=2800) +
        body_sp(
            para(run('\u201eHangokkal gazdagítani az emberek életét\u201d', sz=1800, bold=True, color_hex=ACCENT_GR)) +
            para(run("• Felelős erdészet: FSC-tanúsított faanyag hangszerekhez", sz=1600)) +
            para(run("• CO₂-semlegesség 2050-re – Yamaha Sustainability Plan 2030", sz=1600)) +
            para(run("• Újrahasznosítható anyagok: rézhangszerek, billentyűzetek", sz=1600)) +
            para(run("• Szociális felelősségvállalás: zeneterápia, közösségi programok", sz=1600)) +
            para(run("• A minőség és a természet tisztelete egymást erősíti", sz=1600)),
            457200, 2050000, 5800000, 4300000
        ) +
        stat_block("FSC", "tanúsított\nfaanyag", 6500000, 450000, 2600000, 1400000, ACCENT_GR, sp_id=51) +
        stat_block("2050", "CO₂-semlegesség\ncélév", 9300000, 450000, 2600000, 1400000, ACCENT_OR, sp_id=52) +
        stat_block("2030", "Sustainability\nPlan határidő", 6500000, 2050000, 2600000, 1400000, ACCENT_GR, sp_id=53) +
        stat_block("100%", "megújuló energia\ncélkitűzés", 9300000, 2050000, 2600000, 1400000, ACCENT_OR, sp_id=54) +
        info_box(
            para(run("\"A természet ajándékaiból készülnek\na legjobb hangszerek – ezért kötelességünk\nvédeni a természetet.\"",
                     sz=1400, color_hex="CCDDEE"), align="ctr", bullet=False),
            6500000, 3650000, 5400000, 2700000, ACCENT_GR, 20000, sp_id=60
        )
    )
    return make_slide(content)

def build_slide_summary():
    """Slide 14: Összefoglalás"""
    content = (
        unified_bg(ACCENT_OR) +
        separator_line(457200, 1980000, 11277600, ACCENT_OR, sp_id=33) +
        title_sp("Összefoglalás", 457200, 350000, 11277600, 1400000, sz=4000) +
        body_sp(
            para(run("A Yamaha Corporation 130+ éves mértékadó örökség", sz=1900, bold=True, color_hex=ACCENT_OR)) +
            para(run("")) +
            para(run("✦  1887-től máig: orgonától a koncertflügelig, a fuvolától a szintetizátorig", sz=1700)) +
            para(run("✦  Kézművesség + high-tech = páratlan hangminőség", sz=1700)) +
            para(run("✦  Globális jelenlét 80+ országban, világ egyik vezető márkája", sz=1700)) +
            para(run("✦  Oktatási programok 500 000+ tanulónak évente", sz=1700)) +
            para(run("✦  Fenntarthatósági elköteleződés a jövő generációinak", sz=1700)) +
            para(run("✦  Folyamatos innováció digitális és akusztikus területen egyaránt", sz=1700)) +
            para(run("")) +
            para(run('\u201eYamaha \u2013 Making Waves\u201d', sz=2000, bold=True, color_hex=ACCENT_OR), align="ctr", bullet=False),
            457200, 2050000, 11277600, 4500000
        )
    )
    return make_slide(content)

# ── Új diák adatai ───────────────────────────────────────────────────────────

NEW_SLIDES = [
    {
        'filename': 'slide10.xml',
        'builder': build_slide_piano,
        'rels': make_rels(2),
        'rId': 'rId22',
        'sld_id': 270,
    },
    {
        'filename': 'slide11.xml',
        'builder': build_slide_digital,
        'rels': make_rels(2),
        'rId': 'rId23',
        'sld_id': 271,
    },
    {
        'filename': 'slide12.xml',
        'builder': build_slide_global,
        'rels': make_rels(2),
        'rId': 'rId24',
        'sld_id': 272,
    },
    {
        'filename': 'slide13.xml',
        'builder': build_slide_education,
        'rels': make_rels(2),
        'rId': 'rId25',
        'sld_id': 273,
    },
    {
        'filename': 'slide14.xml',
        'builder': build_slide_sustainability,
        'rels': make_rels(2),
        'rId': 'rId26',
        'sld_id': 274,
    },
    {
        'filename': 'slide15.xml',
        'builder': build_slide_summary,
        'rels': make_rels(2),
        'rId': 'rId27',
        'sld_id': 275,
    },
]

# ── Meglévő diák kezelése ────────────────────────────────────────────────────

def add_transition(xml_content, transition_xml):
    xml_content = re.sub(r'<p:transition[^>]*/>', '', xml_content)
    xml_content = re.sub(r'<p:transition[^>]*>.*?</p:transition>', '', xml_content, flags=re.DOTALL)
    xml_content = xml_content.replace('</p:sld>', f'\n{transition_xml}\n</p:sld>')
    return xml_content

def add_unified_bg_to_existing_slide(xml_content, slide_num):
    """
    Meglévő diákhoz egységes háttér hozzáadása:
    - Sötétkék felső sáv + vékony alsó vonal
    - Halvány gradient overlay az egységes megjelenésért
    """
    top_bar = f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{90+slide_num}" name="UnifiedTopBar{slide_num}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="190500"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{BG_DARK}"><a:alpha val="85000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

    accent_strip = f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{95+slide_num}" name="UnifiedAccent{slide_num}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="190500"/><a:ext cx="12192000" cy="38100"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{ACCENT_OR}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

    bottom_line = f'''<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{85+slide_num}" name="UnifiedBottom{slide_num}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="6820800"/><a:ext cx="12192000" cy="38100"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{ACCENT_BL}"><a:alpha val="70000"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''

    insert_point = '</p:grpSpPr>'
    if insert_point in xml_content:
        xml_content = xml_content.replace(
            insert_point,
            insert_point + '\n' + top_bar + '\n' + accent_strip + '\n' + bottom_line,
            1
        )
    return xml_content

# ── Főfolyamat ───────────────────────────────────────────────────────────────

def process():
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    with zipfile.ZipFile(SRC, 'r') as z:
        z.extractall(WORK)
    print("Extracted PPTX.")

    slides_dir = os.path.join(WORK, 'ppt', 'slides')
    rels_dir = os.path.join(slides_dir, '_rels')

    # Meglévő diák 1-8: egységes háttér + átmenet
    for i in range(1, 9):
        path = os.path.join(slides_dir, f'slide{i}.xml')
        with open(path, 'r', encoding='utf-8') as f:
            xml = f.read()
        xml = add_unified_bg_to_existing_slide(xml, i)
        xml = add_transition(xml, TRANSITIONS[i-1])
        with open(path, 'w', encoding='utf-8') as f:
            f.write(xml)
        print(f"Enhanced slide {i}.")

    # Slide 9 (bibliográfia): csak átmenet
    path9 = os.path.join(slides_dir, 'slide9.xml')
    with open(path9, 'r', encoding='utf-8') as f:
        xml = f.read()
    xml = add_unified_bg_to_existing_slide(xml, 9)
    xml = add_transition(xml, TRANSITIONS[14])
    with open(path9, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("Enhanced slide 9 (bibliography).")

    # Új diák 10-15 létrehozása (fotók nélkül – nincs ismétlés)
    for slide_info in NEW_SLIDES:
        fname = slide_info['filename']
        xml_content = slide_info['builder']()
        slide_num = int(fname.replace('slide', '').replace('.xml', ''))
        trans_idx = slide_num - 1
        xml_content = add_transition(xml_content, TRANSITIONS[trans_idx])

        with open(os.path.join(slides_dir, fname), 'w', encoding='utf-8') as f:
            f.write(xml_content)
        with open(os.path.join(rels_dir, f'{fname}.rels'), 'w', encoding='utf-8') as f:
            f.write(slide_info['rels'])
        print(f"Created {fname}.")

    # presentation.xml frissítése
    pres_path = os.path.join(WORK, 'ppt', 'presentation.xml')
    with open(pres_path, 'r', encoding='utf-8') as f:
        pres_xml = f.read()

    new_slide_ids = ''
    for slide_info in NEW_SLIDES:
        new_slide_ids += f'<p:sldId id="{slide_info["sld_id"]}" r:id="{slide_info["rId"]}"/>'

    pres_xml = pres_xml.replace(
        '<p:sldId id="265" r:id="rId9"/>',
        '<p:sldId id="265" r:id="rId9"/>' + new_slide_ids
    )
    with open(pres_path, 'w', encoding='utf-8') as f:
        f.write(pres_xml)
    print("Updated presentation.xml.")

    # presentation.xml.rels frissítése
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

    # [Content_Types].xml frissítése
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

    # Visszacsomagolás
    if os.path.exists(DST):
        os.remove(DST)
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(WORK):
            for file in files:
                fpath = os.path.join(root, file)
                arcname = os.path.relpath(fpath, WORK)
                zout.write(fpath, arcname)
    print(f"\nKész! Mentve: {DST}")
    print("15 dia – 9 eredeti (egységes háttérrel) + 6 új (fotóismétlés nélkül)")

if __name__ == '__main__':
    process()
