#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""背景開啟 Excel 檔並截圖（不搶前景、不受視窗切換影響）。
- open -g 背景開檔；System Events 自動點掉「格式不符」對話框的「是(Y)」（不設 frontmost）；
- Quartz 取文件視窗ID；screencapture -l 抓視窗內容緩衝（背景/被遮擋也能截）。
需求：已授權執行腳本的終端機 App 的「螢幕錄製」+「輔助使用」。
用法：python excel_shot.py <xls/xlsx路徑> <輸出png路徑>
"""
import sys, subprocess, time

def run_osa(script):
    return subprocess.run(["osascript","-e",script], capture_output=True, text=True).stdout.strip()

def click_dialog_yes():
    # 不設 frontmost；找名稱以「是」或「Yes」開頭的按鈕點掉
    osa='''
    tell application "System Events"
      tell process "Microsoft Excel"
        repeat 8 times
          repeat with w in windows
            try
              repeat with b in (buttons of w)
                set bn to name of b
                if bn starts with "是" or bn starts with "Yes" or bn starts with "開啟" then
                  click b
                  return "clicked:" & bn
                end if
              end repeat
            end try
          end repeat
          delay 0.8
        end repeat
        return "no-dialog"
      end tell
    end tell'''
    return run_osa(osa)

def excel_window_id(min_area=300000):
    """回傳最大面積的 Excel 視窗ID（文件視窗，排除小對話框）"""
    import Quartz
    wins=Quartz.CGWindowListCopyWindowInfo(
        Quartz.kCGWindowListOptionOnScreenOnly|Quartz.kCGWindowListExcludeDesktopElements,
        Quartz.kCGNullWindowID)
    cand=[]
    for w in wins:
        if 'Excel' in w.get('kCGWindowOwnerName',''):
            b=w.get('kCGWindowBounds',{}); area=b.get('Width',0)*b.get('Height',0)
            cand.append((area,int(w['kCGWindowNumber'])))
    cand=[c for c in cand if c[0]>=min_area]
    return max(cand)[1] if cand else None

def shoot(path, out):
    subprocess.run(["osascript","-e",'tell application "Microsoft Excel" to quit saving no'],
                   capture_output=True)
    time.sleep(2)
    subprocess.run(["open","-g","-a","Microsoft Excel", path])
    time.sleep(7)
    print("dialog:", click_dialog_yes())
    time.sleep(3)                       # 等文件視窗渲染
    wid=excel_window_id()
    if not wid:
        print("找不到 Excel 文件視窗"); return False
    subprocess.run(["screencapture","-x","-o","-l",str(wid), out])
    subprocess.run(["osascript","-e",'tell application "Microsoft Excel" to quit saving no'],
                   capture_output=True)
    return True

if __name__=="__main__":
    ok=shoot(sys.argv[1], sys.argv[2])
    if ok:
        from PIL import Image, ImageStat; import numpy as np
        im=Image.open(sys.argv[2]).convert("L"); a=np.asarray(im)
        print("OK", sys.argv[2], "尺寸", Image.open(sys.argv[2]).size,
              "非白 %.1f%%"%((a<235).mean()*100))
