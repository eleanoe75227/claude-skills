---
name: docx-fill-images
description: 把截圖依序插進 docx 的「【此處貼圖】」佔位（變更單、測試報告、維護報告、交付文件都適用），自動配欄寬、裁 macOS 選單列、防空白頁、產完自檢。當使用者丟截圖說「補圖」「把圖插進去」「這是第N張圖」，或文件有【此處貼圖】佔位要結案時觸發。
---

# docx-fill-images（docx 補圖）

## 開工時先跟使用者說（需要使用者提供的資料）
1. 要補圖的 docx
2. 截圖本身，並說明每張對應哪個佔位；沒講就依佔位順序，先用 `--list` 列出來給使用者對
3. 截圖裡如果看不到主機名稱／IP、指令或畫面跟文件寫的不一樣，要先問，**不能自己補**

## 流程
1. **動手前**：確認檔案沒有被 Word 開著（同資料夾有 `~$` 開頭的鎖定檔就先請使用者關檔），並先備份原檔。
2. **列佔位**：`python3 scripts/fill_images.py 檔.docx --list`，看有幾個佔位、各在多寬的欄、前一段在講什麼。
3. **先看截圖，再核對文字**。這一步最常要回頭改，逐張檢查：
   - 截圖用的指令跟文件寫的一樣嗎？（例如文件寫 `php -i`，實際跑的是 `cat php.ini | grep`）
   - 數值跟文件一樣嗎？（例如文件寫 `Off`，截圖顯示 `0`）
   - 是不是多台主機合在一張圖裡？是的話就把佔位、圖號、圖說合併
   - 截圖裡有沒有文件沒提到的主機？要補一句說明，資訊不足就問
   - 文件寫了某個動作，截圖看不出來？問使用者要刪掉還是補圖
   - 截圖裡的時間（例如 `date` 輸出）可以拿來填完成時間；macOS 選單列的時鐘若跟作業日期不同，該張加 `--crop-menubar N`
   - 要改的文字整理好，一次問使用者
4. **改文字**：多出來的佔位段落和圖說要刪掉，圖號重新編。Word 常把一句話拆成好幾個 run，改之前先印出 run 清單。
5. **插圖**：
   ```
   python3 scripts/fill_images.py 輸入.docx 輸出.docx 圖1.png 圖2.png \
     [--crop-menubar 2] [--skip 3] [--fix-trailing-blank] [--check-pages] [--max-height-cm 8]
   ```
   - 圖寬預設是所在表格欄寬（不在表格內則用版心寬）× 0.95；表格最後一列被擠到下一頁時，用 `--max-height-cm` 或 `--width-cm` 縮小
   - `--fix-trailing-blank`：最後一個表格後面的空段落會多擠出一頁空白頁，加這個選項就能縮掉
   - 輸出最後一行「檢查：OK」才能交付；它會檢查佔位都換掉了、圖片關聯沒斷掉、沒有開檔更新功能變數的旗標（updateFields／w:dirty）
6. **看版面**：`--check-pages` 會用 LibreOffice 回報頁數和疑似空白頁；要看細節就轉 PDF 再轉圖目視檢查。頁數以 Word 為準，LibreOffice 只當參考。

## 設計原則
- 只改 `word/document.xml`、rels、`[Content_Types].xml`，並新增 media；其餘部件原封複製。不用 python-docx 存檔，避免頁首頁尾參照跑掉。
- 依賴：Python 3、Pillow；`--check-pages` 另外需要 LibreOffice（`soffice`）和 PyMuPDF（`fitz`）。

## 已知的雷
- 重用 `zipfile.ZipInfo` 寫新檔會改到來源 zip 的 offset，之後再讀來源就壞掉；腳本已處理。
- 圖片段落的行距不能是 `exact`，否則圖會被裁切；腳本會自動拿掉。
