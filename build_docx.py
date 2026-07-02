#!/usr/bin/env python3
"""Build the strategy Word document from strategy_doc_source.md."""
import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn


def set_font(run, name_cn, name_en=None, size=None, bold=None, color=None):
    run.font.name = name_en or name_cn
    r = run._element.rPr
    rFonts = r.find(qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), name_cn)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_runs_with_bold(p, text, font_cn='仿宋_GB2312', size=14):
    """Handle **bold** inline markup."""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])
            set_font(run, '黑体', 'SimHei', size=size, bold=True)
        else:
            run = p.add_run(part)
            set_font(run, font_cn, 'FangSong', size=size)


def para(doc, align=None, space_after=6, line=28):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line)
    if align:
        p.alignment = align
    return p


doc = Document()

# page setup: A4, standard official-document margins
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(3.0)
sec.bottom_margin = Cm(2.5)
sec.left_margin = Cm(2.8)
sec.right_margin = Cm(2.6)

lines = open('/workspace/strategy_doc_source.md', encoding='utf-8').read().split('\n')

first_heading_done = False
i = 0
while i < len(lines):
    line = lines[i].rstrip()
    i += 1
    if not line.strip():
        continue
    if line.strip() == '---':
        continue

    if line.startswith('# '):
        # main title
        p = para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10, line=36)
        run = p.add_run(line[2:].strip())
        set_font(run, '方正小标宋简体', 'SimHei', size=22, bold=True, color=(0xC0, 0x00, 0x00))
        first_heading_done = True
    elif line.startswith('## '):
        text = line[3:].strip()
        p = para(doc, space_after=8, line=32)
        p.paragraph_format.space_before = Pt(14)
        run = p.add_run(text)
        set_font(run, '黑体', 'SimHei', size=16, bold=True)
    elif line.startswith('### '):
        text = line[4:].strip()
        p = para(doc, space_after=6, line=30)
        p.paragraph_format.space_before = Pt(8)
        run = p.add_run(text)
        set_font(run, '楷体_GB2312', 'KaiTi', size=15, bold=True)
    elif re.match(r'^\d+\.\s', line.strip()) or line.strip().startswith('- '):
        text = re.sub(r'^(-\s+|\d+\.\s+)', '', line.strip())
        prefix = ''
        m = re.match(r'^(\d+)\.\s', line.strip())
        if m:
            prefix = m.group(1) + '．'
        else:
            prefix = '——'
        p = para(doc, space_after=4, line=28)
        p.paragraph_format.left_indent = Cm(0.74)
        run = p.add_run(prefix)
        set_font(run, '仿宋_GB2312', 'FangSong', size=14)
        add_runs_with_bold(p, text)
    elif line.strip().startswith('*') and line.strip().endswith('*') and not line.strip().startswith('**'):
        # italic footnote
        p = para(doc, space_after=4, line=24)
        run = p.add_run(line.strip().strip('*'))
        set_font(run, '楷体_GB2312', 'KaiTi', size=12)
        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    elif line.strip().startswith('**') and line.strip().endswith('**') and len(line.strip()) < 60:
        # standalone bold line (subtitle / date)
        p = para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6, line=28)
        run = p.add_run(line.strip()[2:-2])
        set_font(run, '楷体_GB2312', 'KaiTi', size=16, bold=True)
    else:
        # body paragraph with first-line indent
        p = para(doc)
        p.paragraph_format.first_line_indent = Cm(0.99)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_runs_with_bold(p, line.strip())

out = '/workspace/《公路行业未来五年数智化AI新范式预判及发展战略与主要任务（2026～2030）》20260702.docx'
doc.save(out)
print('saved:', out)
