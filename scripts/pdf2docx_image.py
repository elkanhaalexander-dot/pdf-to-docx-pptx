# -*- coding: utf-8 -*-
"""pdf2docx_image.py —— 像素级还原：把 PDF 每页渲染成高清位图，整页铺满嵌入 DOCX。

适用：PPT 式课件、图形密集、含背景图/矢量色块的 PDF。
产出：DOCX 每页 = 原 PDF 该页的高清位图，页面尺寸与原页一致、页边距为 0、
      图片以「浮动锚定」方式精确铺满整页（不受行高约束），
      因此在 Word/WPS 里看到的版式与 PDF 完全相同（像素级），代价是文字不可编辑。

用法：
    python pdf2docx_image.py <输入.pdf> <输出.docx> [--dpi 200] [--quality 92] [--keep]

依赖：同目录 deps/ 内的 PyMuPDF + python-docx + Pillow（均无需安装）。
"""
import argparse
import os
import sys
sys.dont_write_bytecode=True
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
DEPS = os.path.join(os.path.dirname(HERE), 'vendor')
if DEPS not in sys.path:
    sys.path.insert(0, DEPS)

import pymupdf  # noqa: E402
from docx import Document  # noqa: E402
from docx.shared import Pt  # noqa: E402
from docx.enum.text import WD_LINE_SPACING  # noqa: E402
from docx.oxml.ns import qn, nsdecls  # noqa: E402
from docx.oxml import parse_xml  # noqa: E402

EMU_PER_PT = 12700


def render_pages(pdf_path, out_dir, dpi, quality):
    """把每一页渲染成位图；纯文字页用 PNG，含位图页用 JPEG（体积小很多）。"""
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    made = []
    for pno in range(doc.page_count):
        page = doc[pno]
        has_raster = bool(page.get_images(full=True))
        ext = 'jpg' if has_raster else 'png'
        path = os.path.join(out_dir, 'p%03d.%s' % (pno + 1, ext))
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        if ext == 'jpg':
            pix.save(path, jpg_quality=quality)
        else:
            pix.save(path)
        made.append((path, page.rect.width, page.rect.height))
    doc.close()
    return made


def add_full_page_picture(paragraph, image_path, page_w, page_h):
    """把图片作为浮动对象钉在页面 (0,0)，尺寸等于整页，实现无缝铺满。"""
    run = paragraph.add_run()
    run.add_picture(image_path, width=Pt(page_w), height=Pt(page_h))
    drawing = run._element.find(qn('w:drawing'))
    inline = drawing.find(qn('wp:inline'))
    blip = inline.find('.//' + qn('a:blip'))
    embed = blip.get(qn('r:embed'))
    docPr = inline.find(qn('wp:docPr'))
    pid = docPr.get('id')
    cx = int(round(page_w * EMU_PER_PT))
    cy = int(round(page_h * EMU_PER_PT))
    anchor_xml = (
        '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="%d" cy="%d"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        '<wp:docPr id="%s" name="Page %s"/>'
        '<a:graphic %s><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic %s><pic:nvPicPr><pic:cNvPr id="%s" name="page%s"/><pic:cNvPicPr/>'
        '</pic:nvPicPr><pic:blipFill><a:blip r:embed="%s"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:anchor>'
    ) % (nsdecls('wp', 'r', 'a', 'pic'), cx, cy, pid, pid,
         nsdecls('a'), nsdecls('pic'), pid, pid, embed, cx, cy)
    drawing.replace(inline, parse_xml(anchor_xml))


def build_docx(pages, out_path):
    doc = Document()
    body = doc.element.body
    for p in list(doc.paragraphs):
        body.remove(p._element)

    page_w, page_h = pages[0][1], pages[0][2]
    sec = doc.sections[0]
    sec.page_width = Pt(page_w)
    sec.page_height = Pt(page_h)
    for attr in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin',
                 'header_distance', 'footer_distance'):
        setattr(sec, attr, Pt(0))

    # 去掉页眉页脚引用，避免多出空行把图片往下推
    sectPr = sec._sectPr
    for tag in ('w:headerReference', 'w:footerReference'):
        for el in sectPr.findall(qn(tag)):
            sectPr.remove(el)

    # 正文样式也压到 1pt，杜绝任何默认行高占位
    normal = doc.styles['Normal']
    normal.font.size = Pt(1)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    normal.paragraph_format.line_spacing = Pt(1)

    for idx, (path, w_pt, h_pt) in enumerate(pages):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Pt(0)
        pf.right_indent = Pt(0)
        pf.first_line_indent = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(1)
        if idx > 0:
            pf.page_break_before = True
        add_full_page_picture(p, path, w_pt, h_pt)

    doc.save(out_path)
    return out_path


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('docx')
    ap.add_argument('--dpi', type=int, default=200)
    ap.add_argument('--quality', type=int, default=92)
    ap.add_argument('--keep', action='store_true', help='保留渲染出的位图')
    args = ap.parse_args(argv[1:])

    output_parent=os.path.dirname(os.path.abspath(args.docx))
    os.makedirs(output_parent,exist_ok=True)
    tmp = tempfile.mkdtemp(prefix='.pdf2docx-images-',dir=output_parent)
    pages = render_pages(args.pdf, tmp, args.dpi, args.quality)
    build_docx(pages, args.docx)
    total = sum(os.path.getsize(p) for p, *_ in pages)
    print('pages=%d  dpi=%d  images=%.1f MB  ->  %s' % (
        len(pages), args.dpi, total / 1024 / 1024, args.docx))
    if not args.keep:
        for p, *_ in pages:
            try:
                os.remove(p)
            except OSError:
                pass
        try:
            os.rmdir(tmp)
        except OSError:
            pass
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
