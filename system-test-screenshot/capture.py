#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批次截頁：用 roles.py 取得登入狀態（session 復用，失效才彈窗），依序截多個頁面。
資安：帳密只讀本機 .env、腳本不印密碼；截圖存本機。

用法：
  python3 capture.py                         # 只截 APP_HOME_URL（驗證登入用）
  python3 capture.py 名稱=URL 名稱2=URL2 ...  # 名稱結尾加 * ＝整頁截，否則截可視範圍
  --role <角色>：用哪個角色登入（預設 admin）
  --out <資料夾>：截圖存放處（預設 <APP_OUT>/<YYYYMMDD>/）
"""
import sys, datetime, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from roles import get, sync_playwright, OUT, HOME

def main():
    args=sys.argv[1:]; role="admin"; outarg=None
    for flag in ("--role","--out"):
        if flag in args:
            i=args.index(flag); v=args[i+1]; del args[i:i+2]
            if flag=="--role": role=v
            else: outarg=pathlib.Path(v).expanduser()
    targets=[]
    for a in args:
        if "=" in a:
            name,u=a.split("=",1); name=name.strip(); full=name.endswith("*")
            targets.append((name.rstrip("*").strip(), u.strip(), full))
    if not targets: targets=[("首頁", HOME, False)]
    outdir=outarg or OUT/datetime.date.today().strftime("%Y%m%d"); outdir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        g=get(p, role)
        if not g: print("❌ 無法取得登入狀態，中止。"); return
        b,ctx,pg=g
        for i,(name,u,full) in enumerate(targets,1):
            try:
                pg.goto(u, wait_until="networkidle", timeout=20000)
                if "login" in pg.url.lower(): print(f"  ⚠ 第{i}項 {name}：被導回登入頁，session 可能剛失效")
                fn=outdir/f"{i:02d}_{name}.png"; pg.screenshot(path=str(fn), full_page=full)
                print(f"  ✅ {name} → {fn} ({'整頁' if full else '可視範圍'})")
            except Exception as e:
                print(f"  ❌ {name} 失敗：{str(e)[:80]}")
        b.close(); print(f"\n完成，截圖存於：{outdir}")

if __name__=="__main__":
    main()
