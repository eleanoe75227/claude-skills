# REFERENCE — 系統功能測試截圖（卡住才讀）

## 1. Selector 庫（先查這裡，不要偵察）

每個受測系統在下方開一節，記錄實測過的 selector。

**登錄規則**：
- 每節標頭寫**驗證日期**與環境（測試機／正式機）；沒有日期的條目無法判斷是否過期。
- 欄位：元素｜selector｜備註。備註要寫**前置條件**（要先點哪個頁籤、展開哪個面板，元素才會出現在 DOM）與**陷阱**（為什麼不能用另一種直覺寫法）。
- 站台改版、selector 失效 → 重新偵察後**改同一筆並更新日期**，不要另加一筆。
- 優先序：id／name／onclick／title 等穩定屬性 > 可見文字 > class。不用開發者工具「複製 selector」產生的路徑。

### 範例：<系統名稱> 登入頁（YYYY-MM-DD 驗證，測試機）
| 元素 | selector | 備註 |
|---|---|---|
| 帳號 | `#username` | |
| 密碼 | `#password` | |
| 驗證碼圖 | `#captchaImg` | 不做 OCR，由使用者在彈窗輸入 |
| 登入成功判斷 | URL 含 `home` 且不含 `login` | 對應 .env 的 `APP_SUCCESS` |

### 範例：<系統名稱> 某模組（YYYY-MM-DD 驗證）
| 元素 | selector | 備註 |
|---|---|---|
| 首頁模組磁磚 | `[onclick*='<module名>']` | 點擊**開新分頁**，要 `ctx.expect_page()` 接 |
| 進階搜尋 | `div.more[title='進階搜尋']` | 點了才展開篩選面板；展開後**原本的**搜尋鈕被隱藏 → 改點面板內的鈕 |
| 匯出Excel | `a.xls[title='匯出Excel']` | 觸發下載，用 `expect_download()` 攔 |
| 彈窗 | `.modal.in, .modal.show` innerText | 登入後固定跳的公告彈窗會攔截所有點擊 → 先 `pg.evaluate("()=>$('#xxxModal').modal('hide')")` |
| select2 下拉 | 點 `#select2-<id>-container` → `.select2-search__field` 打字 → 點 `.select2-results__option` | 原生 select 沒有 option，`select_option` 無效 |
| 隱藏的檔案上傳 | `input[type=file]` 直接 `set_input_files` | 每檔等上傳 POST 回 200 |
| 換頁會先 ajax 存檔 | 之後 `wait_for_url(...)`＋`networkidle` | 否則報 "page is navigating" |

> 新偵察到的 selector：加進對應系統的表（含備註），下次才不用再偵察。

## 2. 偵察範本（selector 庫沒有時才用）
無頭跑，列出候選元素；**別把整頁 HTML 印出來**（耗 token 又可能含內部資料）：
```python
js="""() => { const out=[];
  for(const el of document.querySelectorAll('a,button,select,input,i,div')){
    const r=el.getBoundingClientRect(); if(r.width===0) continue;
    const cls=(el.className||'').toString();
    out.push({tag:el.tagName, txt:(el.innerText||'').trim().slice(0,12),
      cls:cls.slice(0,25), title:el.getAttribute('title')||'',
      oc:(el.getAttribute('onclick')||'').slice(0,30),
      x:Math.round(r.x), y:Math.round(r.y)}); }
  return out.filter(e=>e.txt||e.title||e.oc).slice(0,40); }"""
for e in page.evaluate(js): print(e)
```
可再按座標區域過濾（如右上角：`r.left>1200 && r.top<200`）。

## 3. 機制原理（為什麼這樣設計）
- **session 復用**：驗證碼只在登入那刻要；登入後伺服器發 session cookie，存成 `session_*.json` 跨次復用 → 平常無頭全自動。過期才需使用者重登一次（roles.py 自動彈窗、帳密已填好）。**不裝 OCR**：session 復用已足夠。
- **憑證警告**：測試機常是自簽憑證，`ignore_https_errors=True` 直接略過。
- **紅框＋箭頭**（annotate.py）：框粗 7px、箭頭自動從空間最大側指向框、尖端與框留 16px 間距。改樣式調 `BOX_W/ARROW_W/GAP`。
- **同帳號只能一處登入的系統**：多個角色共用同一帳號時，登一次後複製 session 檔，不要並行登入。

## 4. Excel 背景截圖原理（excel_shot.py，macOS）
- 系統「匯出Excel」常是 **HTML 表格＋假副檔名 .xls/.xlsx**：Chromium 渲染無樣式、pandas 轉檔表頭會跑掉 → 都不採用。
- 正解：`open -g` 背景開真 Excel（不搶焦點）→ System Events 點掉「格式與副檔名不相符」對話框（按鈕名**用 starts with "是"/"Yes" 比對**，不設 frontmost）→ Quartz `CGWindowListCopyWindowInfo` 取**最大面積**視窗ID（排除小對話框）→ `screencapture -x -o -l <ID>` 抓**視窗內容緩衝**（被遮擋/在背景也截得到）。
- **權限（一次性）**：執行腳本的終端機 App 需在「系統設定→隱私權與安全性」勾**螢幕錄製**＋**輔助使用**。螢幕錄製要重開 App 才生效。
- 沒有 Quartz 時：`soffice --convert-to html` → Playwright 開 html 截圖（轉 PDF 會列高不撐開、文字疊字）。

## 5. 踩雷記錄（別再犯）
| 雷 | 教訓 |
|---|---|
| 用螢幕區域截圖＋搬前景 | 使用者一切視窗就截錯＋被打擾 → 一律用 `-l 視窗ID` 抓內容緩衝 |
| 對話框按鈕寫死「是 (Y)」 | 實名可能無空格 → 用 starts with 比對 |
| `get_by_text().first` 亂點 | 同名文字常有多個匹配（含隱藏選單）→ 用 onclick/title 屬性定位 |
| 展開面板後點原搜尋鈕 | 展開後原鈕被藏（not visible）→ 用面板內的鈕 |
| full_page=True 當預設 | 首頁常超長；預設可視範圍，需要整頁才加 `*` |
| 圖插入後被壓扁 | 段落固定行高 → 必設單行間距 |
| 下載檔用 openpyxl 開 | BadZipFile（HTML 假 xlsx）→ 走 excel_shot.py |
| 只用 DOM 抓字寫說明 | 純圖示按鈕（＋、放大鏡、日曆）抓不到 → 截圖後一定開圖目視補 |
| 工具列按鈕密集時的紅框箭頭 | 自動挑空間最大側，可能壓住相鄰按鈕文字 → 目視檢查，必要時改用 annotate_multi |
| 裁下方留白被左側樹狀選單騙 | 樹的灰底延伸到底 → 只看右側內容區判斷最後一列內容 |
| 畫面全無樣式（手機版排版） | 無頭 Chromium 的 UA 含 HeadlessChrome 被判成行動裝置 → new_context 加桌機 `user_agent` |
| 多角色測試 | 每個帳號各存一份 session；先確認各角色帳號不是同一組（只比對、不印出）→ `roles.py --check` |
| 登入失敗但沒有任何訊息 | 某些系統帳密錯誤只是 302 回登入頁 → 先懷疑密碼過期 |

### 判定類常見誤判
| 雷 | 教訓 |
|---|---|
| 說「防呆沒做／功能沒做」 | 先走三段：①在原本資料上出現 ②換一筆乾淨資料仍出現 ③**同一條規則的正向案例會亮**（證明檢核機制本身是活的）。少了③只證明「這次沒看到提示」，不能寫「未實作」 |
| 必填沒填就送出，沒看到後面規則的提示 | 系統在必填檢核就擋下，後面的規則根本沒跑到 → 標「未執行」，不是「未實作」 |
| 欄位停用、按鈕灰掉 | 先確認是不是該狀態本來就不能編輯（案件階段鎖定），再判缺陷 |
| 查詢結果 0 筆就下結論 | 先確認 session 沒過期（過期常回空清單或導回登入頁），再確認查詢條件 |
| 「沒觀察到」當成「不存在」 | 先確認操作真的有送出（截到送出後畫面），觀察方式真的看得到 |
| 間歇性問題只跑一次 | 跑一次成功不代表已修好，至少重跑一次 |
| 多數樣本都被判異常 | 通常是判斷標準錯了，不是系統壞 |
| 同一組帳號同時跑兩批 | session 互搶，看起來像登入異常 → 同帳號不並行 |
| 測試機沒更版就判不符 | 先確認測試機版本與待測版本一致，不一致整項標「未執行」 |

## 6. 擴充到新系統
`.env` 填好登入頁、首頁、成功判斷字串、帳密欄 selector 與各角色帳密 → `python3 roles.py --check` → `python3 roles.py admin` 登入一次 → 在 §1 開一節記錄該系統 selector。登入流程特殊（多步驟、單一簽入轉導）時，複製 roles.py 另寫一支登入模組，介面維持 `get(p, role)`。
