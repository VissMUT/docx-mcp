# -*- coding: utf-8 -*-
"""Генерация assets/reference.docx — эталон стилей pandoc для docx_build.

Берётся дефолтный reference.docx pandoc (--print-default-data-file) и
правится python-docx: Times New Roman 12, по ширине, отступ 1.25 см,
заголовки 13-14 pt bold, A4, поля 20/10/20/10 мм, PAGE в футере.

Запуск: python3 tools/make_reference.py [выходной файл]
(требует pandoc и python-docx; в Dockerfile — этап сборки образа).
"""
import subprocess
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

OUT = sys.argv[1] if len(sys.argv) > 1 else 'assets/reference.docx'

r = subprocess.run(
    ['pandoc', '-o', OUT, '--print-default-data-file', 'reference.docx'],
    capture_output=True, text=True)
if r.returncode != 0:
    sys.exit('pandoc: %s' % r.stderr[:300])

doc = Document(OUT)


def set_style(name, size, bold=None, align=None, indent=None):
    try:
        st = doc.styles[name]
    except KeyError:
        return
    st.font.name = 'Times New Roman'
    st.font.size = Pt(size)
    st.font.color.rgb = RGBColor(0, 0, 0)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(attr), 'Times New Roman')
    if bold is not None:
        st.font.bold = bold
    if align is not None:
        st.paragraph_format.alignment = align
    if indent is not None:
        st.paragraph_format.first_line_indent = Cm(indent)


# Обычный текст: pandoc использует Body Text / First Paragraph / Compact,
# Normal — базовый (отступ НЕ ставим в Normal: он протекает в ячейки таблиц)
set_style('Normal', 12, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
for name in ('Body Text', 'First Paragraph', 'Compact'):
    set_style(name, 12, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=1.25)
set_style('Block Text', 12, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=1.25)

# Заголовки: H1 по центру 14 bold, H2 слева 13 bold
set_style('Heading 1', 14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
set_style('Heading 2', 13, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
set_style('Heading 3', 12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
set_style('Title', 16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

# Страница: A4, поля 20/10/20/10 мм
sec = doc.sections[0]
sec.page_width, sec.page_height = Mm(210), Mm(297)
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(1.0)
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(2.0)

# PAGE в футере (по центру)
p = sec.footer.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run()
fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText')
instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
for el in (fld1, instr, fld2):
    r._r.append(el)

doc.save(OUT)
print('reference.docx -> %s' % OUT)
