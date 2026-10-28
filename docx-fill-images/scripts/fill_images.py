#!/usr/bin/env python3
"""把 docx 裡的圖片佔位段落（預設「【此處貼圖】」）依序換成圖片。

只改 word/document.xml、rels、[Content_Types].xml、新增 media；其餘部件原封複製
（不用 python-docx 存檔，避免頁首頁尾參照跑掉）。

用法：
  python3 fill_images.py 輸入.docx 輸出.docx 圖1.png 圖2.png ... [選項]
  python3 fill_images.py 輸入.docx --list          # 只列出佔位段落與所在欄寬，不改檔

選項：
  --marker TEXT        佔位文字（預設【此處貼圖】）；段落文字被拆成多個 run 也找得到
  --width-cm N         圖寬（cm）；預設 auto＝所在表格欄寬或版心寬 × 0.95
  --max-height-cm N    圖高上限（cm），超過就等比縮小（預設 不限）
  --crop-menubar N[,N] 第 N 張圖（從 1 起，或 all）裁掉 macOS 選單列（含時鐘，避免日期對不上）
  --crop-top PX        自訂裁掉上方像素，套用在 --crop-menubar 指定的圖
  --skip N[,N]         第 N 個佔位（從 1 起）不換圖、保留原樣；圖片數要對應扣掉
  --fix-trailing-blank 最後一個表格後的空段落縮成極小，避免多擠出一頁空白頁
  --check-pages        產完用 LibreOffice 轉 PDF 回報頁數（PDF 放暫存目錄，不留在輸出旁）

產完會自動檢查：佔位是否都換掉、rels 有無斷掉、settings.xml updateFields／w:dirty。
"""
import argparse, io, os, re, subprocess, sys, tempfile, zipfile

try:
    from PIL import Image
except ImportError:
    sys.exit('需要 Pillow：pip install pillow')

EMU_PER_CM = 360000
DXA_PER_CM = 567.0
P_RE = re.compile(r'<w:p[ >].*?</w:p>', re.S)
T_RE = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')


def para_text(p):
    return ''.join(T_RE.findall(p))


def find_placeholders(x, marker):
    """回傳 [(start, end, 段落xml)]，依文件順序。"""
    out = []
    for m in P_RE.finditer(x):
        if marker in para_text(m.group(0)):
            out.append((m.start(), m.end(), m.group(0)))
    return out


def text_width_cm(x):
    w = re.search(r'<w:pgSz w:w="(\d+)"', x)
    mar = re.search(r'<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"', x)
    if not w or not mar:
        return 15.0
    return (int(w.group(1)) - int(mar.group(1)) - int(mar.group(2))) / DXA_PER_CM


def cell_width_cm(x, pos):
    """佔位段落若在表格儲存格內，回傳該格寬（cm），否則 None。"""
    tc = x.rfind('<w:tc>', 0, pos)
    tc2 = x.rfind('<w:tc ', 0, pos)
    tc = max(tc, tc2)
    if tc < 0 or x.rfind('</w:tc>', 0, pos) > tc:
        return None
    m = re.search(r'<w:tcW w:w="(\d+)" w:type="dxa"', x[tc:tc + 600])
    if not m:
        return None
    # 扣掉儲存格左右預設邊界約 0.38cm
    return int(m.group(1)) / DXA_PER_CM - 0.38


def prep_image(path, crop_top, crop_menubar):
    im = Image.open(path)
    if not crop_menubar:
        crop_top = None
    elif crop_top is None:
        # macOS 選單列約 37 邏輯點；Retina 截圖寬度 > 2000px 視為 2x
        crop_top = 74 if im.size[0] > 2000 else 37
    if crop_top:
        im = im.crop((0, crop_top, im.size[0], im.size[1]))
    buf = io.BytesIO()
    im.save(buf, 'PNG')
    return buf.getvalue(), im.size


def drawing_run(rid, pid, name, cx, cy):
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{pid}" name="圖片 {pid}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{pid}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )


def replace_para(p, run_xml):
    """保留段落 pPr，所有 run 換成圖片 run；行距 exact 改 auto，避免圖片被裁切。"""
    ppr = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
    ppr = ppr.group(0) if ppr else ''
    ppr = re.sub(r'(<w:spacing[^>]*?) w:line="\d+" w:lineRule="exact"', r'\1', ppr)
    head = re.match(r'<w:p[ >][^>]*>|<w:p>', p).group(0)
    return head + ppr + run_xml + '</w:p>'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('out', nargs='?')
    ap.add_argument('images', nargs='*')
    ap.add_argument('--list', action='store_true')
    ap.add_argument('--marker', default='【此處貼圖】')
    ap.add_argument('--width-cm', type=float)
    ap.add_argument('--max-height-cm', type=float)
    ap.add_argument('--crop-menubar', default='')
    ap.add_argument('--crop-top', type=int)
    ap.add_argument('--skip', default='')
    ap.add_argument('--fix-trailing-blank', action='store_true')
    ap.add_argument('--check-pages', action='store_true')
    a = ap.parse_args()

    zin = zipfile.ZipFile(a.src)
    x = zin.read('word/document.xml').decode('utf8')
    phs = find_placeholders(x, a.marker)
    tw = text_width_cm(x)

    if a.list or not a.out:
        print(f'佔位 {len(phs)} 個（版心寬 {tw:.1f}cm）：')
        for i, (s, e, p) in enumerate(phs, 1):
            cw = cell_width_cm(x, s)
            prev = [para_text(m) for m in P_RE.findall(x[max(0, s - 4000):s])]
            prev = [t for t in prev if t.strip()][-1:] or ['']
            print(f'  {i}. 欄寬 {cw:.1f}cm' if cw else f'  {i}. 不在表格內', '｜前一段：', prev[0][:50])
        return

    crop_all = a.crop_menubar.strip() == 'all'
    crop_set = set() if crop_all else {int(n) for n in a.crop_menubar.split(',') if n.strip()}
    skip = {int(n) for n in a.skip.split(',') if n.strip()}
    targets = [ph for i, ph in enumerate(phs, 1) if i not in skip]
    if len(targets) != len(a.images):
        sys.exit(f'佔位 {len(targets)} 個（已扣 skip），圖片 {len(a.images)} 張，數量不符')

    rels = zin.read('word/_rels/document.xml.rels').decode('utf8')
    ct = zin.read('[Content_Types].xml').decode('utf8')
    max_rid = max([int(n) for n in re.findall(r'Id="rId(\d+)"', rels)] or [0])
    max_pid = max([int(n) for n in re.findall(r'<wp:docPr id="(\d+)"', x)] or [0])
    names = set(zin.namelist())
    new_media = {}

    # 從後往前換，位置不會跑掉
    for k in range(len(targets) - 1, -1, -1):
        s, e, p = targets[k]
        data, (w, h) = prep_image(a.images[k], a.crop_top, crop_all or (k + 1) in crop_set)
        width = a.width_cm or (cell_width_cm(x, s) or tw) * 0.95
        if a.max_height_cm and width * h / w > a.max_height_cm:
            width = a.max_height_cm * w / h
        cx = int(width * EMU_PER_CM)
        cy = int(cx * h / w)
        rid = f'rId{max_rid + 1 + k}'
        n = 1
        while f'word/media/fill_{k + 1}_{n}.png' in names:
            n += 1
        media = f'word/media/fill_{k + 1}_{n}.png'
        new_media[media] = data
        rels = rels.replace('</Relationships>',
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{media[5:]}"/></Relationships>')
        x = x[:s] + replace_para(p, drawing_run(rid, max_pid + 1 + k, os.path.basename(media), cx, cy)) + x[e:]
        print(f'  佔位→{os.path.basename(a.images[k])}  {width:.1f}×{width * h / w:.1f}cm')

    if 'Extension="png"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1)

    if a.fix_trailing_blank:
        k = x.rfind('</w:tbl>')
        m = P_RE.search(x, k) if k >= 0 else None
        if m and not para_text(m.group(0)).strip() and '<w:drawing' not in m.group(0) and '<w:sectPr' not in m.group(0):
            tiny = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
            x = x[:m.start()] + tiny + x[m.end():]
            print('  已縮小表格後空段落')

    with zipfile.ZipFile(a.out + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zo:
        for orig in zin.infolist():
            # 複製一份 ZipInfo：直接重用會改到 zin 的 header_offset，之後再讀 zin 就壞掉
            info = zipfile.ZipInfo(orig.filename, orig.date_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = orig.external_attr
            if info.filename == 'word/document.xml':
                zo.writestr(info, x.encode('utf8'))
            elif info.filename == 'word/_rels/document.xml.rels':
                zo.writestr(info, rels.encode('utf8'))
            elif info.filename == '[Content_Types].xml':
                zo.writestr(info, ct.encode('utf8'))
            else:
                zo.writestr(info, zin.read(orig.filename))
        for n, d in new_media.items():
            zo.writestr(n, d)
    os.replace(a.out + '.tmp', a.out)

    # 檢查
    problems = []
    left = len(find_placeholders(x, a.marker)) - len(skip)
    if left:
        problems.append(f'還有 {left} 個佔位沒換')
    ids = set(re.findall(r'Id="(rId\d+)"', rels))
    bad = sorted(set(re.findall(r'r:(?:embed|id)="(rId\d+)"', x)) - ids)
    if bad:
        problems.append(f'斷掉的關聯：{bad}')
    settings = zin.read('word/settings.xml').decode('utf8') if 'word/settings.xml' in zin.namelist() else ''
    if 'updateFields' in settings or 'w:dirty="true"' in x:
        problems.append('有 updateFields／w:dirty 旗標（開檔會跳更新提示）')
    print('檢查：' + ('；'.join(problems) if problems else 'OK'))

    if a.check_pages:
        with tempfile.TemporaryDirectory() as td:
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', td, a.out],
                           capture_output=True)
            pdf = os.path.join(td, os.path.splitext(os.path.basename(a.out))[0] + '.pdf')
            try:
                import fitz
                d = fitz.open(pdf)
                blank = [i + 1 for i, pg in enumerate(d) if len(pg.get_text().strip().splitlines()) <= 1 and not pg.get_images()]
                print(f'頁數（LibreOffice）：{len(d)}' + (f'；疑似空白頁：{blank}' if blank else ''))
            except Exception as ex:
                print('頁數檢查失敗：', ex)


if __name__ == '__main__':
    main()
