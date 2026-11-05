#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""協作登入＋沿途截圖（正式機需本人身分登入，如單一簽入／自然人憑證等「使用者本人操作」的登入流程）。
開有頭視窗到起始網址，使用者自己操作；每次換頁自動截一張（page＋含網址列的整窗由使用者另截）。

結束條件（任一成立就收工；避免登入失敗停在原頁時腳本關不掉）：
  ① 成功：回到目標主機且網址不含 login
  ② 出現提示訊息：頁面文字含 --stop 指定字樣（預設「查無」「失敗」「錯誤」「逾時」）
  ③ 閒置：連續 --idle 秒（預設 60）沒換頁
  ④ 總時限：--max 秒（預設 300）

用法：
  python3 assist_login.py <起始網址> <輸出資料夾> [--host app.example.gov.tw] [--stop 查無,失敗] [--idle 60] [--max 300]
輸出：<輸出資料夾>/NN_<說明>_page.png、url_log.txt（每步網址）。
⚠ 正式機截圖常含身分證字號、使用者帳號 → 插報告前先打碼；不能用的原圖移到 `_勿用/`（見 SKILL.md 資安底線）。
"""
import sys, time, argparse, pathlib
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright

UA="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"

ap=argparse.ArgumentParser()
ap.add_argument("url"); ap.add_argument("out")
ap.add_argument("--host", help="目標系統主機（預設取起始網址的主機）")
ap.add_argument("--stop", default="查無,失敗,錯誤,逾時")
ap.add_argument("--idle", type=int, default=60)
ap.add_argument("--max", type=int, default=300)
a=ap.parse_args()
host=(a.host or urlparse(a.url).netloc).lower()
stops=[s for s in a.stop.split(",") if s]
out=pathlib.Path(a.out); out.mkdir(parents=True, exist_ok=True)
log=open(out/"url_log.txt","a",encoding="utf-8")

def shot(pg, n, tag):
    fn=out/f"{n:02d}_{tag}_page.png"
    try: pg.screenshot(path=str(fn))
    except Exception: return
    log.write(f"{n:02d}\t{tag}\t{pg.url}\n"); log.flush()
    print(f"  📸 {fn.name}  {urlparse(pg.url).netloc}", flush=True)

with sync_playwright() as p:
    b=p.chromium.launch(headless=False)
    ctx=b.new_context(ignore_https_errors=True, viewport={"width":1920,"height":1080}, user_agent=UA)
    pg=ctx.new_page(); pg.goto(a.url, wait_until="domcontentloaded")
    print(f"⚠ 請在彈出視窗自行登入（閒置{a.idle}秒或{a.max}秒後自動結束）", flush=True)
    n=0; last=None; t0=tlast=time.time(); why="總時限"
    while time.time()-t0 < a.max:
        try: pg.wait_for_load_state("domcontentloaded", timeout=3000)
        except Exception: pass
        u=pg.url
        if u!=last:
            n+=1; last=u; tlast=time.time(); time.sleep(1.5)
            shot(pg, n, urlparse(u).netloc.split(".")[0] or "page")
        nl=urlparse(u).netloc.lower()
        if nl==host and "login" not in u.lower():
            why="登入成功"; break
        try: txt=pg.inner_text("body", timeout=2000)
        except Exception: txt=""
        hit=next((s for s in stops if s in txt), None)
        if hit and nl==host:
            n+=1; shot(pg, n, "提示訊息"); why=f"出現提示「{hit}」"; break
        if time.time()-tlast > a.idle:
            n+=1; shot(pg, n, "閒置結束"); why="閒置逾時"; break
        time.sleep(1)
    print(f"結束：{why}；截圖 {n} 張 → {out}", flush=True)
    b.close()
