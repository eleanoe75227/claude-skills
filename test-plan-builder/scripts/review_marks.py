# -*- coding: utf-8 -*-
"""抽取審閱者在Word裡留下的標記，並核對修正後文件是否仍殘留被標示的文字。

用法：
  python3 review_marks.py extract <審閱回饋.docx>                 # 列出所有標記（含上下文），做成核對清單
  python3 review_marks.py residual <審閱回饋.docx> <新版.docx> [...]  # 被反黃的文字是否仍出現在新版文件

標記種類：反黃(或其他螢光色)、非黑色文字（審閱者說明）、Word註解（含被註解的文字）、追蹤修訂(插入/刪除)。
註解與說明文字無法自動判斷「改好了沒」，請依 extract 結果逐條寫成可檢查的條件（見 sa_feedback_verification.md）。
"""
import sys, re, zipfile
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def load(path, part="word/document.xml"):
    z = zipfile.ZipFile(path)
    return etree.fromstring(z.read(part)) if part in z.namelist() else None


def ptext(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def context(p):
    """所在表格列的文字（若在表格內），否則段落本身。"""
    tr = p
    while tr is not None and tr.tag != W + "tr":
        tr = tr.getparent()
    return " | ".join(ptext(c) for c in tr.iter(W + "p")) if tr is not None else ptext(p)


def is_heading(p):
    st = p.find(W + "pPr/" + W + "pStyle")
    v = st.get(W + "val") if st is not None else ""
    return v.lower().startswith("heading") or v in ("1", "2", "3", "10")


def marks(path):
    doc = load(path)
    out, heading = [], ""
    for p in doc.iter(W + "p"):
        if is_heading(p):
            heading = ptext(p)
        hl_runs, color_runs = {}, {}
        for r in p.iter(W + "r"):
            t = "".join(x.text or "" for x in r.findall(W + "t"))
            if not t:
                continue
            h = r.find(W + "rPr/" + W + "highlight")
            c = r.find(W + "rPr/" + W + "color")
            if h is not None and h.get(W + "val") not in (None, "none"):
                hl_runs.setdefault(h.get(W + "val"), []).append(t)
            if c is not None and c.get(W + "val", "").lower() not in ("000000", "auto", ""):
                color_runs.setdefault(c.get(W + "val"), []).append(t)
        for col, ts in hl_runs.items():
            out.append(("反" + col, heading, "".join(ts), context(p)))
        for col, ts in color_runs.items():
            out.append(("色字" + col, heading, "".join(ts), context(p)))
        for tag, name in ((W + "ins", "修訂-插入"), (W + "del", "修訂-刪除")):
            for e in p.iter(tag):
                t = "".join(x.text or "" for x in e.iter(W + "t", W + "delText"))
                if t:
                    out.append((name, heading, t, context(p)))
        for cs in p.iter(W + "commentRangeStart"):
            out.append(("註解#" + cs.get(W + "id"), heading, "", context(p)))
    com = load(path, "word/comments.xml")
    ctext = {}
    if com is not None:
        for c in com.iter(W + "comment"):
            ctext[c.get(W + "id")] = "".join(c.itertext()).strip()
    return [(k, h, (ctext.get(k[3:], "") if k.startswith("註解#") else t), ctx) for k, h, t, ctx in out]


def cmd_extract(path):
    rows = marks(path)
    print("| # | 種類 | 章節 | 標記內容／註解 | 上下文 | 處理方式 |")
    print("|---|---|---|---|---|---|")
    seen = set()
    n = 0
    for k, h, t, ctx in rows:
        key = (k, t, ctx)
        if key in seen:
            continue
        seen.add(key)
        n += 1
        clean = lambda s: s.replace("|", "／").replace("\n", " ")[:120]
        print(f"| {n} | {k} | {clean(h)} | {clean(t)} | {clean(ctx)} |  |")
    print(f"\n共 {n} 筆標記（標記種類：{sorted(set(r[0].rstrip('0123456789#') for r in rows))}）", file=sys.stderr)


def cmd_residual(path, news, color="yellow"):
    frags = []
    for k, h, t, ctx in marks(path):
        if k == "反" + color and len(t.strip()) >= 3:
            frags.append(t.strip())
    frags = list(dict.fromkeys(frags))
    texts = {}
    for n in news:
        texts[n] = "\n".join(ptext(p) for p in load(n).iter(W + "p"))
    bad = 0
    for f in frags:
        key = re.sub(r"^\d+\.\s*", "", f).strip("。；，、 ")
        hit = [n for n, t in texts.items() if key and key in t]
        if hit:
            bad += 1
        print(("仍殘留 " if hit else "已處理 ") + f[:60] + (f"  → {hit}" if hit else ""))
    print(f"\n反{color}片段 {len(frags)} 段，仍殘留 {bad} 段（殘留者需逐一判斷：漏改，或為需求原文等合理保留）")


if __name__ == "__main__":
    if sys.argv[1] == "extract":
        cmd_extract(sys.argv[2])
    elif sys.argv[1] == "residual":
        cmd_residual(sys.argv[2], sys.argv[3:])
