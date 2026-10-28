# claude-skills

可分享的 Claude Code skills，內容已去識別化（不含客戶名稱、契約內容等業務機敏資訊），供下載、參考、自行調整使用。

## 目前收錄

- **test-plan-builder**：依RFP需求、SA文件、會議紀錄與驗收條款，建立可雙向追溯的測試計畫。詳見該資料夾內的 SKILL.md。
- **docx-fill-images**：把截圖依序插進 docx 的「【此處貼圖】」佔位，自動配欄寬、裁 macOS 選單列、防尾端空白頁、產完自檢（佔位、圖片關聯、更新功能變數旗標）。只用 zipfile 改 XML，不經 python-docx 存檔。
- **notebooklm**：透過 notebooklm-py CLI 操作 Google NotebookLM（建筆記本、加來源、問答、產生/下載各類產出）。轉載自 [teng-lin/notebooklm-py](https://github.com/teng-lin/notebooklm-py)（MIT 授權，v0.7.3），未修改。使用前需先安裝並登入：`uv tool install "notebooklm-py[browser]"` → `notebooklm login` → `notebooklm auth check --test --json`。登入憑證存在本機 `~/.notebooklm/`，請各自用自己的 Google 帳號登入，勿互相複製。

## 使用方式

複製對應資料夾到你的 `.claude/skills/` 底下即可。內容供參考與調整，實際使用時請依自己專案的素材與情境判斷，不保證直接適用。

## 更新紀錄

### docx-fill-images

- **2026-10-28**：收錄。

### notebooklm

- **2026-10-10**：收錄（上游 notebooklm-py v0.7.3 原檔）。

### test-plan-builder

- **2026-09-28**
  - 開工前新增三個詢問決策：交付形式（一份主文件／主文件＋附件）、案例撰寫方式（預設全部寫完／先交3～5個樣本給SA確認寫法）、審閱分工。
  - 依SA實際審閱回饋新增10條寫作規則（通用原則7條＋可依客戶慣例調整的預設寫法3條），前後對照範例見 `references/case_writing_rules.md`。
  - 「定測試範圍」擴大到整合、壓力、資料移轉測試；不在RFP範圍的項目不列入。
  - 新增SA審閱回饋的抽取與自動核對流程（`references/sa_feedback_verification.md`、`scripts/review_marks.py`）。
  - 新增「Word與建置程式同步」規則、待確認標記分兩種（給SA確認／PM自行處理）。
- **2026-09-21**：初版。
