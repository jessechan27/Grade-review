# 專案備忘 — 成績查詢網站

> 這份檔案會在這個資料夾開工時自動載入，目的是不用重新想一次設計決策。
> **這個檔案會進公開 repo，所以不要在這裡寫任何查詢碼或密碼。**

最後更新：2026-11-02

---

## 這是什麼

教媒科課程的學生成績查詢靜態網站。學生輸入自己的學號＋查詢碼，看到自己各項成績
與目前加權總分。MLB 球員卡風格（深藍 `#041e42` ＋紅 `#bf0d3e`，Oswald 窄體字）。

- 線上：<https://jessechan27.github.io/Grade-review/>
- Repo：<https://github.com/jessechan27/Grade-review>（**public**）

## 架構

```
Grade review.xlsx  ──┐
secrets.json       ──┼─→  build.py  ──→  index.html   （要上傳的成品）
template.html      ──┘               └─→  成績總表.csv  （本機看全班用）
                                     └─→  查詢碼名單.csv（發給學生用）
```

**只改 `template.html`，不要直接改 `index.html`** —— 它是產生出來的，下次 build 會被覆蓋。

改完一律跑 `python build.py`，然後 `git add index.html template.html` → commit → push。

## 核心設計決策（含原因）

### 1. 每筆成績用「學號＋查詢碼」各自加密

金鑰 = `PBKDF2-HMAC-SHA256(學號 + "|" + 查詢碼, salt, 60000, 64 bytes)`
→ 前 32 bytes 當加密金鑰，後 32 bytes 當查表用的 token。
加密是 SHA-256 counter-mode keystream XOR，再加 HMAC-SHA256 前 16 bytes 當驗證碼。

**為什麼**：這是公開的靜態網站，沒有後端。若成績直接寫在 HTML 裡，看原始碼就能撈走
全班。改用這個設計後，網頁原始碼只有亂碼，而且**知道別人學號也沒用** —— 限制做在
密碼學層，不是只有畫面擋住。

### 2. 網站上刻意「沒有」全班總表

曾經做過管理端（主密碼解開一份全班名冊 blob），**後來整個拿掉了**。
原因：那等於把全班成績加密後放在公開網站上，只靠一組密碼擋著，而老師本來就有
Excel 原檔，不需要在瀏覽器裡看全班。

全班視圖改成本機產生的 `成績總表.csv`（build 時自動更新，含算好的總分，依總分排序）。

> 若日後要求「我要在網頁上看全班」，請先提醒上面這個取捨。

### 3. 查詢碼必須跨 build 穩定

查詢碼存在 `secrets.json`，`build.py` 每次會沿用既有的碼，只幫新學生配新碼。
**為什麼**：碼已經發給學生了，重新產生會讓所有人的碼失效。
→ `secrets.json` 弄丟 = 全班重發，請務必備份。

碼用 8 碼、字母表 `23456789ABCDEFGHJKMNPQRSTUVWXYZ`（去掉易混淆的 0/O/1/I/L），
顯示成 `XXXX-XXXX`。比對時大小寫不分、忽略非英數字元（`normCode()`）。

### 4. 權重從 Excel 標題列讀，不寫死

標題寫「口頭簡報（50%）」就抓 50%。調權重只要改 Excel 標題，不用動程式。
- 含 `點名扣分` 的欄 → 從總分直接扣除
- 含 `學期總分` 的欄 → 忽略（由網頁自己算）

### 5. 未評分一律以 0 分計入總分

這是使用者明確指定的規則。總分 = Σ(分數 × 權重%)，空白儲存格算 0。
另外顯示「已評分權重」與「得分率」(已得分 ÷ 已評分權重) 當輔助資訊。

### 6. 純 JS 密碼學 + WebCrypto 快速路徑

`template.html` 內含手寫的 SHA-256 / HMAC / PBKDF2。
有 `crypto.subtle` 時走 WebCrypto（約 16ms），沒有時退回純 JS（約 1.9 秒）。
**為什麼要留純 JS**：確保 `file://` 直接開啟也能用，不依賴安全內容環境。
兩條路徑的輸出已驗證與 Python 的 `hashlib.pbkdf2_hmac` 完全一致。

---

## 絕對不要 commit 的檔案

| 檔案 | 外洩後果 |
|---|---|
| `Grade review.xlsx` | 全班明文成績 |
| `secrets.json` | 所有查詢碼 ＝ 等同全班成績 |
| `查詢碼名單.csv` | 同上 |
| `成績總表.csv` | 全班明文成績 |

`.gitignore` 已排除這四類。**repo 是公開的，而且 git 歷史一旦寫進去就極難清除**，
所以任何新增檔案前先確認：

```bash
git ls-files | grep -iE "xlsx|secrets|名單|總表"   # 應該沒有輸出
```

每次 push 前值得再掃一次 `index.html` 有沒有明文姓名或查詢碼。

---

## 踩過的坑

- **Excel 學號欄是 int 與 str 混用的**（第一列是數字格式，其餘文字）。`build.py` 兩種都處理。
- **Windows 主控台是 cp950**，直接 print 中文會 `UnicodeEncodeError`。
  → 用 `PYTHONIOENCODING=utf-8`，或把輸出寫成 UTF-8 檔再讀。
- **GitHub Pages 有 CDN 快取**，push 後線上可能還是舊版 30～60 秒。
  驗證時加 `?cb=$RANDOM` 避開快取，不然會誤判成沒更新。
- **檔案總管隱藏副檔名**（`HideFileExt=1`），`index.html` 只顯示成「index」。
  曾因此出現 `index.html.html`（看到「index」就改名成「index.html」，Windows 又補一次）。
  使用者也曾回報圖示看起來像 PDF —— `.html` 與 `.pdf` 都關聯到 Edge，
  實際是圖示快取問題，檔案本身正常。
- **`<span>` 要加 `display:block`** 才吃得到 width/height（成績長條圖曾因此不顯示）。

## 使用者偏好

- 溝通用繁體中文。
- 會自己更新 Excel 後說「我更新了檔案」，這時要重新 build 並比對這次多了哪些成績。
- 很在意隱私與權限邊界，相關取捨要主動講清楚。
