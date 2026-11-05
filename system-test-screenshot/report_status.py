#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""測試報告 docx 截圖格盤點（流程D 用）。
每輪補圖後跑一次，取代手動數「剩幾格」，數字直接貼對內彙整／回報使用者。

格式假設：一需求一張表，某列第一格是項目編號（如 2-1、115-7），最後一列是截圖列；
截圖列內「一般文字段落」＝格子標題，標題段或其後段落有圖＝已插圖；
藍字【截圖】…＝待補；【說明】…＝以文字說明代替；紅字【請SA確認】另計。

用法：python3 report_status.py <報告.docx> [--detail]
"""
import sys, re
from docx import Document
from docx.oxml.ns import qn

ID=re.compile(r"^(\d{3}-\d{1,2}|\d-\d)$")

def paras(cell):
    for p in cell._tc.iter(qn("w:p")):
        txt="".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
        yield txt, len(p.findall(".//"+qn("w:drawing")))

def scan(tbl):
    item=next((r.cells[0].text.strip() for r in tbl.rows if ID.match(r.cells[0].text.strip())), None)
    if not item: return None
    slots=[]; sa=0; in_sa=False
    for txt,img in paras(tbl.rows[-1].cells[0]):
        if txt.startswith("【請SA確認】"): sa+=1; in_sa=True; continue
        if in_sa: continue   # 【請SA確認】之後的條列（1. 2. …）都屬紅字待確認段，不是格子
        if txt.startswith("【截圖】"):
            if slots and slots[-1][1]=="?": slots[-1][1]="待補"
            continue
        if txt.startswith("【說明】") or txt.startswith("【備註】"):
            if slots and slots[-1][1]=="?": slots[-1][1]="說明"
            continue
        if txt:
            slots.append([txt, "圖" if img else "?"])
        elif img and slots:
            slots[-1][1]="圖"
    for s in slots:
        if s[1]=="?": s[1]="空白"
    return item, slots, sa

def main():
    if len(sys.argv)<2: sys.exit(__doc__)
    doc=Document(sys.argv[1]); detail="--detail" in sys.argv
    tot={"圖":0,"待補":0,"說明":0,"空白":0}; sa_tot=0
    print(f"{'項目':<8}{'圖':>3}{'待補':>4}{'說明':>4}{'空白':>4}{'SA':>4}")
    for t in doc.tables:
        r=scan(t)
        if not r: continue
        item,slots,sa=r; c={k:sum(1 for s in slots if s[1]==k) for k in tot}
        for k in tot: tot[k]+=c[k]
        sa_tot+=sa
        print(f"{item:<8}{c['圖']:>3}{c['待補']:>4}{c['說明']:>4}{c['空白']:>4}{sa:>4}")
        if detail:
            for txt,st in slots:
                if st!="圖": print(f"    - [{st}] {txt[:40]}")
    print(f"{'合計':<8}{tot['圖']:>3}{tot['待補']:>4}{tot['說明']:>4}{tot['空白']:>4}{sa_tot:>4}")
    print("※ 圖＝已插圖格數；判定（通過/未執行…）仍以對內彙整為準，本表只數格子。")

if __name__=="__main__":
    main()
