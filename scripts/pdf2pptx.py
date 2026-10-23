# -*- coding: utf-8 -*-
"""One page per editable slide, with native/residual OCR text and artwork.

The CLI uses layout_common; the older construction helpers remain for imports.
Pictures, logos and diagram graphics are raster, not editable native shapes.
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
from pptx import Presentation  # noqa: E402
from pptx.util import Pt, Emu  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

WHITE_FALLBACK = RGBColor(0x1F, 0x1F, 0x1F)

FONT_MAP = {
    'SimSun': 'SimSun',
    'TimesNewRomanPS-BoldMT': 'Times New Roman',
    'TimesNewRomanPSMT': 'Times New Roman',
    'Arial': 'Arial',
    'Arial-BoldMT': 'Arial',
}


def norm_font(name):
    base = name.split(',')[0].strip()
    if base in FONT_MAP:
        return FONT_MAP[base]
    low = base.lower()
    if 'simsun' in low or 'song' in low:
        return 'SimSun'
    if 'times' in low:
        return 'Times New Roman'
    if 'arial' in low:
        return 'Arial'
    if 'calibri' in low:
        return 'Calibri'
    # PDF 里等线/DengXian 的字体名是乱码字节，统一映射到「等线」
    return 'DengXian'


def set_ea_font(run, font_name):
    """python-pptx 只设置 latin 字体，中文需要额外设置 ea（eastAsian）字体。"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font_name)


def clear_bold_inheritance(run):
    """显式写 b="0"，否则某些主题/占位符会让中文被渲染成粗体。"""
    rPr = run._r.get_or_add_rPr()
    for attr in ('b', 'i'):
        rPr.set(attr, '0')


def span_color(s, on_fill=False):
    """返回文字颜色。

    on_fill=True 表示这段文字位于有色块之上：此时纯白文字保持白色，
    因为色块已被保留，白字在色块上正是原稿的样子。
    on_fill=False 表示底下没有色块：白字改为深灰，否则白纸上看不见。
    """
    if s['color'] == 0xFFFFFF:
        return RGBColor(0xFF, 0xFF, 0xFF) if on_fill else WHITE_FALLBACK
    c = s['color']
    return RGBColor((c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF)


def _rgb(fill):
    r = min(255, max(0, int(round(fill[0] * 255))))
    g = min(255, max(0, int(round(fill[1] * 255))))
    b = min(255, max(0, int(round(fill[2] * 255))))
    return RGBColor(r, g, b)


def _path_points(dr, max_points=160):
    """把 PyMuPDF 的 drawing items 转成多边形顶点（曲线离散成折线）。

    返回 (points, closed)。closed 由路径首尾是否重合决定：
    箭头这类图形在原稿里是「首尾重合的开放路径」，若强行 close=True
    会多连一条线，在相邻色块上留下斜线轮廓。
    """
    pts = []
    for item in dr['items']:
        op = item[0]
        if op == 'l':                       # 直线: ('l', p1, p2)
            p1, p2 = item[1], item[2]
            if not pts:
                pts.append((p1.x, p1.y))
            pts.append((p2.x, p2.y))
        elif op == 'c':                     # 三次贝塞尔: ('c', p1, p2, p3, p4)
            p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
            if not pts:
                pts.append((p1.x, p1.y))
            for k in range(1, 9):           # 每段曲线离散成 8 段折线
                t = k / 8.0
                mt = 1 - t
                x = (mt ** 3 * p1.x + 3 * mt * mt * t * p2.x
                     + 3 * mt * t * t * p3.x + t ** 3 * p4.x)
                y = (mt ** 3 * p1.y + 3 * mt * mt * t * p2.y
                     + 3 * mt * t * t * p3.y + t ** 3 * p4.y)
                pts.append((x, y))
        elif op == 're':                    # 矩形
            r = item[1]
            pts.extend([(r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)])
        elif op == 'qu':                    # 四边形
            q = item[1]
            pts.extend([(q.ul.x, q.ul.y), (q.ur.x, q.ur.y),
                        (q.lr.x, q.lr.y), (q.ll.x, q.ll.y)])

    # 去重（相邻点过近）
    out = []
    for p in pts:
        if not out or abs(p[0] - out[-1][0]) > 0.3 or abs(p[1] - out[-1][1]) > 0.3:
            out.append(p)

    # 首尾是否重合 → 决定是否闭合
    closed = False
    if len(out) >= 3:
        dx = abs(out[0][0] - out[-1][0])
        dy = abs(out[0][1] - out[-1][1])
        closed = dx < 1.0 and dy < 1.0

    # 限流：点太多会让 PPTX 体积暴涨
    if len(out) > max_points:
        step = len(out) / float(max_points)
        out = [out[int(i * step)] for i in range(max_points)]
    return out, closed


def _is_convex(pts):
    """判断多边形是否凸。圆角矩形/圆是凸的，箭头是凹的。"""
    n = len(pts)
    if n < 3:
        return False
    sign = 0
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        x3, y3 = pts[(i + 2) % n]
        cross = (x2 - x1) * (y3 - y2) - (y2 - y1) * (x3 - x2)
        if abs(cross) < 1e-6:
            continue
        s = 1 if cross > 0 else -1
        if sign == 0:
            sign = s
        elif s != sign:
            return False
    return True


def _is_panel(dr, rect, page_w, page_h):
    """判断填充图形是否「面板型」（圆角框、高亮条、色带）。

    面板型：凸多边形（矩形/圆角矩形/圆/椭圆），保留原稿投影更接近观感。
    非面板：凹多边形（箭头、折线块），必须排在面板下方，否则会盖住面板。
    """
    pts, closed = _path_points(dr, max_points=400)
    if len(pts) < 3:
        return True
    if not _is_convex(pts):
        return False
    w, h = rect.width, rect.height
    if w <= 0 or h <= 0:
        return False
    ratio = max(w, h) / min(w, h)
    if ratio > 6:                           # 细长条不加投影
        return False
    return True


def add_vector_fills(slide, page, page_w, page_h):
    """把页面上的矢量填充色块还原成 PPT 自由形状。

    与 DOCX 版不同：PPT 是演示用途，保留色块/高亮/箭头更利于观看，
    它们也是原稿版式的一部分（如「个人」「企业」的黄色高亮、彩色圆角框、箭头）。
    - 用真实路径顶点（含贝塞尔离散）还原形状，圆就是圆、箭头就是箭头；
    - 跳过整页背景（面积 > 页面 40%）；
    - 保留原稿的投影效果，让层次与原稿一致。
    """
    from pptx.util import Emu
    added = []
    shapes = []          # (是否面板, shape元素)
    for dr in page.get_drawings():
        fill = dr.get('fill')
        if fill is None:
            continue
        rect = dr['rect']
        if rect.width <= 0.5 or rect.height <= 0.5:
            continue
        if (rect.width * rect.height) > page_w * page_h * 0.4:
            continue

        pts, closed = _path_points(dr)
        if len(pts) < 3:
            continue

        # build_freeform 的坐标单位是 EMU（1pt = 12700 EMU），scale=1 表示不缩放
        emu = lambda v: int(round(v * 12700))
        try:
            builder = slide.shapes.build_freeform(
                start_x=emu(pts[0][0]), start_y=emu(pts[0][1]), scale=1)
            builder.add_line_segments(
                [(emu(x), emu(y)) for x, y in pts[1:]], close=closed)
            shape = builder.convert_to_shape()
        except Exception:
            continue

        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)
        shape.line.fill.background()
        shape.line.width = Pt(0)

        panel = _is_panel(dr, rect, page_w, page_h)
        shape.shadow.inherit = panel

        # 先摘下元素，稍后统一按层次插回
        sp = shape._element
        sp.getparent().remove(sp)
        shapes.append((panel, sp))
        added.append((rect.x0, rect.y0, rect.x1, rect.y1))

    # build_freeform 创建的形状 id 全部是 2，必须重新编号，
    # 否则 PowerPoint/LibreOffice 无法区分它们，渲染顺序会错乱。
    for i, (panel, sp) in enumerate(sorted(shapes, key=lambda t: t[0])):
        cNvPr = sp.find('.//' + qn('p:cNvPr'))
        if cNvPr is not None:
            cNvPr.set('id', str(100 + i))
            cNvPr.set('name', 'Shape %d' % (100 + i))
        slide.shapes._spTree.insert(2, sp)
    return added

def add_page_bitmap(slide, page, page_w, page_h, tmp_dir, pno, dpi=200):
    """整页渲染成位图铺满（用于纯图片页）。"""
    os.makedirs(tmp_dir, exist_ok=True)
    path = os.path.join(tmp_dir, 'page%03d.png' % (pno + 1))
    page.get_pixmap(dpi=dpi, alpha=False).save(path)
    pic = slide.shapes.add_picture(path, 0, 0, Pt(page_w), Pt(page_h))
    sp = pic._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return path


def _drop_full_page_images(page, page_area):
    """把「整页背景图」从页面资源里摘掉，使其不再被绘制。

    背景图在 PDF 里是一张大 XObject。把它从 /Resources/XObject 字典中删除后，
    内容流里的 `/ImN Do` 就成了空操作：背景消失，而矢量色块、箭头、文字不受影响。
    """
    doc = page.parent
    big_names = []
    for im in page.get_images(full=True):
        xref = im[0]
        name = im[7] if len(im) > 7 else None
        if not name:
            continue
        for r in page.get_image_rects(xref):
            if r.width * r.height > page_area * 0.40:
                big_names.append(name)
                break
    if not big_names:
        return

    for name in big_names:
        try:
            # 删除 /Resources/XObject/<name> 这一项
            doc.xref_set_key(page.xref, 'Resources/XObject/%s' % name, 'null')
        except Exception:
            pass


def add_artwork_bitmap(slide, page, page_w, page_h, tmp_dir, pno, dpi=220,
                       drop_background=True):
    """把「图形层」（色块/箭头/高亮/投影）渲染成底图，文字另行叠加为可编辑文本框。

    为什么不用 PPT 自由形状还原箭头：PDF 里的箭头是多段折线构成的凹多边形，
    转成 custGeom 后，LibreOffice 及部分 PowerPoint 版本会忽略 path 的闭合方式，
    按外接矩形填满，出现盖住面板的斜线（实测多轮都无法稳定规避）。
    改用「图形层位图垫底 + 文字可编辑」：视觉与原稿一致，文字仍可自由编辑。

    drop_background=True 时把大图（背景照片、页脚色带、logo）一起抹掉，
    只留色块/箭头/高亮这些真正参与排版的图形。
    """
    os.makedirs(tmp_dir, exist_ok=True)
    path = os.path.join(tmp_dir, 'art%03d.png' % (pno + 1))

    tmp = pymupdf.open()
    tmp.insert_pdf(page.parent, from_page=pno, to_page=pno)
    p2 = tmp[0]

    page_area = page_w * page_h
    # 摘掉整页背景照片：把页面 XObject 资源里那张大图移除，
    # 再重写内容流，使其不再绘制该图。
    if drop_background:
        _drop_full_page_images(p2, page_area)

    # 抹掉文字：只删文字，保留 line art（色块、箭头、投影）
    for block in p2.get_text('dict')['blocks']:
        if block['type'] != 0:
            continue
        for line in block['lines']:
            for s in line['spans']:
                if s['text'].strip():
                    r = pymupdf.Rect(s['bbox'])
                    r.normalize()
                    p2.add_redact_annot(r, fill=None)

    try:
        p2.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                            text=pymupdf.PDF_REDACT_TEXT_REMOVE)
    except TypeError:
        p2.apply_redactions()

    pix = p2.get_pixmap(dpi=dpi, alpha=False)
    pix.save(path)
    tmp.close()

    pic = slide.shapes.add_picture(path, 0, 0, Pt(page_w), Pt(page_h))
    sp = pic._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return path


def build(pdf_path, pptx_path, tmp_dir=None, keep_fills=True):
    src = pymupdf.open(pdf_path)
    page_w = src[0].rect.width
    page_h = src[0].rect.height
    if tmp_dir is None:
        tmp_dir = os.path.join(tempfile.gettempdir(), 'pdf2pptx_tmp')

    prs = Presentation()
    prs.slide_width = Pt(page_w)
    prs.slide_height = Pt(page_h)
    blank = prs.slide_layouts[6]  # 空白版式

    for pno in range(src.page_count):
        page = src[pno]
        slide = prs.slides.add_slide(blank)

        # 纯图片页（如目录页）：整页保留，避免空白
        if not page.get_text().strip():
            add_page_bitmap(slide, page, page_w, page_h, tmp_dir, pno)
            continue

        # 图形层（色块/箭头/高亮/投影）作为底图，文字作为可编辑文本框叠加在上层。
        # 这样既保证视觉与原稿一致，又保留文字可编辑。
        fill_rects = []
        if keep_fills:
            add_artwork_bitmap(slide, page, page_w, page_h, tmp_dir, pno)
            # 底图保留了色块，因此色块上的白字应保持白色
            for dr in page.get_drawings():
                if dr.get('fill') is None:
                    continue
                r = dr['rect']
                if r.width * r.height > page_w * page_h * 0.4:
                    continue
                fill_rects.append((r.x0, r.y0, r.x1, r.y1))

        for block in page.get_text('dict')['blocks']:
            if block['type'] != 0:
                continue
            for ln in block['lines']:
                spans = [s for s in ln['spans'] if s['text']]
                text = ''.join(s['text'] for s in spans)
                if not text.strip():
                    continue

                lx0, ly0, lx1, ly1 = ln['bbox']
                lw = max(lx1 - lx0, 1.0)
                lh = max(ly1 - ly0, 1.0)
                size = max(round(s['size'], 1) for s in spans)

                # 对齐方式：按行中心 vs 页面中心判断
                center = (lx0 + lx1) / 2.0
                if abs(center - page_w / 2.0) < 12 and lw < page_w * 0.9:
                    align = PP_ALIGN.CENTER
                elif lx1 > page_w - 60 and lx0 > page_w * 0.5:
                    align = PP_ALIGN.RIGHT
                else:
                    align = PP_ALIGN.LEFT

                # 文本框留余量（PowerPoint 字宽与 PDF 度量有差异），按对齐方式锚定
                want_w = lw * 1.45 + 30
                if align == PP_ALIGN.CENTER:
                    w = min(want_w, page_w)
                    x = center - w / 2.0
                elif align == PP_ALIGN.RIGHT:
                    w = min(want_w, page_w)
                    x = (lx1 + 8) - w
                else:
                    x = lx0
                    w = min(want_w, page_w - x)
                x = max(0.0, min(x, max(page_w - w, 0)))
                w = max(min(w, page_w - x), 1.0)
                y = max(ly0 - 2.0, 0.0)
                h = min(lh * 1.6 + 6, page_h - y)

                box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
                tf = box.text_frame
                tf.word_wrap = False
                tf.auto_size = None
                tf.margin_left = 0
                tf.margin_right = 0
                tf.margin_top = 0
                tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.TOP

                para = tf.paragraphs[0]
                para.alignment = align
                para.line_spacing = Pt(lh)

                for s in spans:
                    run = para.add_run()
                    run.text = s['text']
                    font = norm_font(s['font'])
                    run.font.size = Pt(round(s['size'], 1))
                    run.font.name = font
                    run.font.bold = 'Bold' in s['font']
                    # 判断这段文字是否落在某个色块内：是则保留白色
                    sb = s['bbox']
                    cx = (sb[0] + sb[2]) / 2.0
                    cy = (sb[1] + sb[3]) / 2.0
                    on_fill = any(
                        f[0] <= cx <= f[2] and f[1] <= cy <= f[3]
                        for f in fill_rects
                    ) if keep_fills else False
                    run.font.color.rgb = span_color(s, on_fill=on_fill)
                    set_ea_font(run, font)

    prs.save(pptx_path)
    src.close()
    return pptx_path, page_w, page_h


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('pptx')
    ap.add_argument('--no-fills', action='store_true', help='不还原矢量色块')
    ap.add_argument('--keep', action='store_true', help='保留纯图片页的渲染位图')
    ap.add_argument('--ocr',choices=['auto','off'],default='auto')
    ap.add_argument('--ocr-backend',choices=['rapid','windows'],default='rapid')
    ap.add_argument('--ocr-all-regions',action='store_true')
    ap.add_argument('--visuals',choices=['preserve','none'],default='preserve')
    args = ap.parse_args(argv[1:])
    from layout_common import convert
    if args.keep:print('NOTE: the new shared engine always removes task-local temporary renderings.')
    report=convert(args.pdf,args.pptx,'pptx',args.ocr,args.ocr_backend,'none' if args.no_fills else args.visuals,args.ocr_all_regions)
    print('OK -> %s (%s pages)'%(args.pptx,len(report['pages'])))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
