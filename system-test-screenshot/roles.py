#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用多角色登入（Playwright）。每個角色各存一份 session，失效才彈窗請使用者手動登入（例如輸入驗證碼）。

設定全放在本資料夾的 .env（不要進 git）：
  APP_LOGIN_URL=https://your-system/login        # 登入頁
  APP_HOME_URL=https://your-system/home          # 需登入才看得到的頁，用來判斷 session 是否有效
  APP_SUCCESS=home                               # 登入成功後網址會包含的字串
  APP_USER_SEL=#username                         # 帳號欄 selector
  APP_PASS_SEL=#password                         # 密碼欄 selector
  APP_USER=... / APP_PASS=...                    # 管理員（預設角色 admin）
  APP_<角色>_USER=... / APP_<角色>_PASS=...      # 其他角色，例：APP_STAFF_USER → get(p,"staff")
  APP_OUT=outputs/截圖測試                        # session 與截圖存放處（選填；相對於目前目錄）

開工前檢查：python3 roles.py --check（各角色是否同一組帳號＋session 是否有效；不彈窗、不印帳號）
登入/預熱：python3 roles.py admin staff   （逐一彈窗）

程式內用法：
  sys.path.insert(0, "<本skill路徑>"); from roles import get, sync_playwright
  with sync_playwright() as p:
      b, ctx, pg = get(p, "staff")
"""
import sys, pathlib
from playwright.sync_api import sync_playwright

SKILL=pathlib.Path(__file__).resolve().parent
ENV=SKILL/".env"
VIEW={"width":1920,"height":1080}
# 無頭 Chromium 預設 UA 含 HeadlessChrome，部分系統會判成手機版 → 固定用桌機 UA
UA="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"

def load_env():
    cfg={}
    if not ENV.exists(): sys.exit(f"找不到 {ENV}，請先依 roles.py 開頭說明建立 .env")
    for l in ENV.read_text(encoding="utf-8").splitlines():
        l=l.strip()
        if l.startswith("#") or "=" not in l: continue
        k,v=l.split("=",1); cfg[k.strip()]=v.split("  #")[0].strip()
    return cfg
cfg=load_env()
OUT=pathlib.Path(cfg.get("APP_OUT","outputs/截圖測試")).expanduser()
HOME=cfg["APP_HOME_URL"]; OK=cfg.get("APP_SUCCESS","")

def keys(role):
    if role=="admin": return "APP_USER","APP_PASS","session_state.json"
    R=role.upper(); return f"APP_{R}_USER", f"APP_{R}_PASS", f"session_{role}.json"

def roles():
    rs=["admin"]+[k[4:-5].lower() for k in cfg if k.startswith("APP_") and k.endswith("_USER") and k!="APP_USER"]
    return [r for r in rs if cfg.get(keys(r)[0])]

def authed(pg):
    u=pg.url.lower(); return (OK.lower() in u if OK else True) and "login" not in u

def _ctx(b, st=None):
    kw=dict(ignore_https_errors=True, viewport=VIEW, user_agent=UA)
    if st: kw["storage_state"]=str(st)
    return b.new_context(**kw)

def get(p, role="admin", headless=True, popup=True, wait=240):
    """回傳 (browser, ctx, page)；session 有效就無頭復用，否則彈窗登入（popup=False 時直接回 None）"""
    uk,pk,sf=keys(role); st=OUT/sf
    if st.exists():
        b=p.chromium.launch(headless=headless); ctx=_ctx(b, st); pg=ctx.new_page()
        try: pg.goto(HOME, wait_until="networkidle", timeout=15000)
        except Exception: pass
        if authed(pg): return b,ctx,pg
        b.close()
    if not popup: return None
    b=p.chromium.launch(headless=False); ctx=_ctx(b); pg=ctx.new_page()
    pg.goto(cfg["APP_LOGIN_URL"], wait_until="domcontentloaded")
    pg.fill(cfg.get("APP_USER_SEL","#username"), cfg[uk]); pg.fill(cfg.get("APP_PASS_SEL","#password"), cfg[pk])
    print(f"⚠ 帳密已填好，請在彈出視窗完成登入（驗證碼等）【{role}】，最多 {wait} 秒", flush=True)
    try: pg.wait_for_url(lambda u: authed(pg), timeout=wait*1000)
    except Exception: pass
    if authed(pg):
        OUT.mkdir(parents=True, exist_ok=True); ctx.storage_state(path=str(st))
        print(f"✅ {role} 登入成功"); return b,ctx,pg
    b.close(); return None

def session_ok(p, role):
    """不彈窗：用存檔 session 打需登入的頁，回 True/False/None(無 session 檔)。
    要實際打頁面判斷，不能看檔案修改時間（檔案新但伺服器端可能已失效）"""
    st=OUT/keys(role)[2]
    if not st.exists(): return None
    b=p.chromium.launch(headless=True); ctx=_ctx(b, st); pg=ctx.new_page()
    try: pg.goto(HOME, wait_until="networkidle", timeout=15000)
    except Exception: pass
    ok=authed(pg); b.close(); return ok

if __name__=="__main__":
    args=sys.argv[1:]
    if args==["--check"]:
        rs=roles(); users={r:cfg[keys(r)[0]] for r in rs}
        print("角色:", rs)
        for i,a in enumerate(rs):
            for c in rs[i+1:]:
                if users[a]==users[c]: print(f"⚠ {a} 與 {c} 是同一組帳號")
        with sync_playwright() as p:
            for r in rs:
                ok=session_ok(p,r)
                print(f"  {r}: session", {True:"有效",False:"已過期（跑 roles.py "+r+" 重登）",None:"尚未登入過"}[ok])
        print("檢查完成（未印出帳號）")
    else:
        with sync_playwright() as p:
            for r in args or ["admin"]:
                g=get(p,r)
                print(r, "登入成功" if g else "登入失敗")
                if g: g[0].close()
