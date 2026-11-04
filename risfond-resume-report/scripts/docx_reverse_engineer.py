# -*- coding: utf-8 -*-
"""反编译 docx：段落级格式 + 页眉页脚 XML + 图片插入尺寸 + 表格属性
用途：拿到一份真实样本报告，逐属性对标，不靠猜。
用法：python docx_reverse_engineer.py <样本.docx>
"""
import sys, zipfile, io, os, re
from docx import Document
from docx.shared import Emu
from docx.oxml.ns import qn
from PIL import Image


def dump(path):
    print('=' * 72)
    print('FILE:', os.path.basename(path))
    d = Document(path)

    # ---------- 1. 页面设置 ----------
    for si, sec in enumerate(d.sections):
        print(f'\n--- section{si} 页面设置 ---')
        print(f'  页边距 T{sec.top_margin.cm:.2f} B{sec.bottom_margin.cm:.2f} '
              f'L{sec.left_margin.cm:.2f} R{sec.right_margin.cm:.2f}')
        hd = sec.header_distance.cm if sec.header_distance else None
        fd = sec.footer_distance.cm if sec.footer_distance else None
        print(f'  页眉距边 {hd} 页脚距边 {fd}')

    # ---------- 2. Normal 样式 ----------
    st = d.styles['Normal']
    print(f'\nNormal: font={st.font.name} size={st.font.size.pt if st.font.size else None} '
          f'line_spacing={st.paragraph_format.line_spacing} '
          f'space_before={st.paragraph_format.space_before} '
          f'space_after={st.paragraph_format.space_after}')

    # ---------- 3. 页眉 / 页脚（含 XML 原文，判断图片到底在不在页眉） ----------
    print('\n--- 页眉/页脚 ---')
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if re.match(r'word/(header|footer)\d*\.xml', n):
                xml = z.read(n).decode('utf-8', 'ignore')
                kind = '页眉' if 'header' in n else '页脚'
                texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
                jc = re.findall(r'<w:jc w:val="([^"]+)"', xml)
                imgs = re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', xml)
                print(f'  [{kind}] {n}')
                print(f'      文本: {texts}')
                print(f'      对齐: {jc}')
                for cx, cy in imgs:
                    print(f'      图片: {int(cx) / 360000:.2f}cm x {int(cy) / 360000:.2f}cm')

    # ---------- 4. 正文图片插入尺寸 ----------
    print('\n--- 正文图片插入尺寸 ---')
    for i, para in enumerate(d.paragraphs):
        if 'graphic' not in para._p.xml:
            continue
        for el in para._p.findall('.//' + qn('wp:inline')):
            ext = el.find(qn('wp:extent'))
            if ext is not None:
                print(f'  para[{i}] {Emu(int(ext.get("cx"))).cm:.2f}cm x '
                      f'{Emu(int(ext.get("cy"))).cm:.2f}cm')

    # ---------- 5. 段落逐条格式 ----------
    print('\n--- 段落格式（仅非空/含图） ---')
    for i, para in enumerate(d.paragraphs):
        txt = para.text
        has_img = 'graphic' in para._p.xml
        if not txt.strip() and not has_img:
            continue
        pf = para.paragraph_format
        bits = []
        if pf.alignment is not None:
            bits.append(f'align={pf.alignment}')
        if pf.first_line_indent is not None:
            bits.append(f'首行缩进={Emu(pf.first_line_indent).cm:.2f}cm')
        if pf.space_before is not None:
            bits.append(f'段前={pf.space_before.pt}pt')
        if pf.space_after is not None:
            bits.append(f'段后={pf.space_after.pt}pt')
        if pf.line_spacing is not None:
            bits.append(f'行距={pf.line_spacing}')
        szs = sorted({r.font.size.pt for r in para.runs if r.font.size})
        if szs:
            bits.append(f'字号={szs}')
        tag = '<图>' if has_img else ''
        print(f'  [{i}]{tag} {txt[:50]!r} | {" ".join(bits)}')
        # run 级：加粗/颜色
        for r in para.runs:
            if not r.text.strip():
                continue
            col = r.font.color.rgb if (r.font.color and r.font.color.rgb) else None
            ea = None
            try:
                ea = r.element.rPr.rFonts.get(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
            except Exception:
                pass
            if r.bold or col or ea:
                print(f'        run b={r.bold} col={col} ea={ea} {r.text[:30]!r}')

    # ---------- 6. 表格 ----------
    print('\n--- 表格 ---')
    for ti, t in enumerate(d.tables):
        print(f'  table{ti}: style={t.style.name if t.style else None} '
              f'{len(t.rows)}x{len(t.columns)}')
        for ri, row in enumerate(t.rows):
            widths = [f'{Emu(c.width).cm:.2f}' if c.width else '?' for c in row.cells]
            print(f'    row{ri} widths={widths}')
            print(f'      {" | ".join(c.text.replace(chr(10), "/") for c in row.cells)}')

    # ---------- 7. media 原图尺寸 ----------
    print('\n--- media 原图 ---')
    try:
        with zipfile.ZipFile(path) as z:
            for n in sorted(z.namelist()):
                if n.startswith('word/media/') and not n.endswith('/'):
                    try:
                        im = Image.open(io.BytesIO(z.read(n)))
                        print(f'    {n} {im.size}')
                    except Exception:
                        pass
    except Exception as e:
        print('    err', e)


if __name__ == '__main__':
    for f in sys.argv[1:]:
        try:
            dump(f)
        except Exception as e:
            print('FAIL', f, e)
