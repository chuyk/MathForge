# 📐 阿凱老師的數學考卷改題排版神器 (Web Edition)

本專案為專用於數學科段考、模擬考之「智能改題、精準程式繪圖、與標準 JIS B4 Word 排版輸出」的 Web 應用系統。
支援部署至 **Streamlit Community Cloud**，並由使用者自行提供 Google Gemini API Key（BYOK 模式，零伺服器費用風險）。

---

## 🌟 核心特色

1. **JIS B4 學校大考規格**：紙張 257 × 364 mm、邊界 22 mm、中文標楷體 13pt、英數與公式 Times New Roman 13pt 斜體。
2. **三層式選項智慧調配（全面零表格）**：
   - 簡短選項：單行四欄並列。
   - 中等選項：雙行雙欄（Word 原生製表位 Tab Stop 於 4.2 英吋對齊）。
   - 長選項（> 20 字）：自動切換為一個選項獨立佔一行（四行直列），絕不發生右欄折行溢出。
3. **Word 原生 OMML 方程式（全卷公式統一 13pt）**：依託微軟原生 `MML2OMML.XSL` 與 `latex2mathml`，公式在 Word 中完全可點選編輯，嚴禁水平打字式分數。後端於所有方程式 run 自動注入 `<w:sz w:val="26"/>` 與 `<w:szCs w:val="26"/>`，使題目與詳解之所有公式嚴格鎖定為 **13pt**，徹底杜絕 Word 預設 11pt 縮小落差！
4. **雙版本成果交付**：
   - **檔案 1**：`全卷試題與後附詳解(B4_13pt版).docx`（題目在先，分頁後為簡答表與詳解）
   - **檔案 2**：`逐題詳解教師備課卷(B4_13pt版).docx`（每題題目緊跟著該題詳解）
5. **詳解純化與深紅字體**：詳解文字與公式一律統一為深紅色 (`RGB(180, 0, 0)`)，題目與選項維持標準黑字。
6. **智慧算式脫殼分流（極速防卡死）**：AI 提示詞與後端轉換器雙重防護，純數字、變數與選項代號自動脫殼為輕量級文字 Run，方程式物件精簡 56% 以上，徹底解決微軟 Word 全選複製 (`Ctrl+A` / `Ctrl+C`) 單核 100% 卡死轉圈問題！
7. **全面解除相容模式**：後端自動調用 `upgrade_to_modern_word_mode` 將微軟 Office `compatibilityMode` 鎖定為 `15`（Word 2013/2016/2019/2021/365 原生現代模式），徹底移除視窗標題列的「[相容模式]」，解鎖原生 SVG 向量圖形與現代方程式功能！

---

## 🚀 本機快速啟動測試

1. 安裝相依套件：
   ```bash
   pip install -r requirements.txt
   ```
2. 啟動 Streamlit 網頁：
   ```bash
   streamlit run app.py
   ```
3. 瀏覽器將自動開啟 `http://localhost:8501`。

---

## 🌐 部署至 GitHub 與 Streamlit Community Cloud（超簡單 3 步驟）

### 步驟 1：在 GitHub 建立新倉庫
1. 前往 [GitHub.com](https://github.com/) 點擊 **New repository**。
2. 倉庫名稱輸入 `MathForge-Web`（可設為 Public 或 Private）。

### 步驟 2：將本目錄推送至 GitHub
在終端機中切換至本目錄（`MathForge_Web/`），依序執行：
```bash
git init
git add .
git commit -m "Initial commit of MathForge Web"
git branch -M main
git remote add origin https://github.com/您的帳號/MathForge-Web.git
git push -u origin main
```

### 步驟 3：在 Streamlit Cloud 一鍵部署
1. 前往 [Streamlit Community Cloud](https://share.streamlit.io/) 並使用 GitHub 帳號登入。
2. 點擊右上角 **"New app"**。
3. 選擇剛剛建立的倉庫：`您的帳號/MathForge-Web`，Branch 選擇 `main`，Main file path 填寫 `app.py`。
4. 點擊 **"Deploy!"** 按鈕。
5. **大約 1 分鐘後，您的專屬網址即正式上線！**（例如：`https://mathforge-web.streamlit.app`）

---

## 🔑 使用者操作指南

1. 開啟網頁，於左側核心設定輸入 **系統啟動碼**（`kai`）。
2. 貼上 **Google Gemini API Key**（系統自動記住至本地/瀏覽器 localStorage，下次免再輸入）。
3. 選擇 AI 模型（支援最新官方前瞻模型：`gemini-3.8-flash`、`gemini-3.7-flash`、`gemini-3.6-flash`、`gemini-3.5-flash`、`gemini-3.5-flash-lite`）。
4. 上傳原始試卷（支援 `.docx` 或 `.pdf`）。
5. 點擊 **「🚀 開始智能改題與 B4 考卷排版」**。
6. 系統將即時進行 API 握手確認，並全自動完成素養改題、精確繪圖、OMML 公式轉換與 JIS B4 排版，點擊按鈕即可下載雙版本 Word 檔案！
