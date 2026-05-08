# 內部業務查詢系統

萊廷照明鋁線槽系列產品的純前端業務查詢網頁，供內部業務人員在手機/平板快速查詢規格與定價。

## 專案目標

- 業務人員輸入料號（如 HB128）或描述需求（如「天花板嵌入」）→ 立即顯示符合產品
- 支援手機、平板響應式，字體與觸控區域針對中年視力優化
- 純靜態無後端，開啟 HTML 即可使用

## 技術架構

```
鋁線槽系列規格表.xlsx
    ↓ python3 build_search.py
products.json
    ↓ fetch（瀏覽器）
search.html（純前端，原生 JS，無框架）
```

## 檔案清單

| 檔案 | 說明 |
|------|------|
| `build_search.py` | 讀取 xlsx → 輸出 products.json（含圖片路徑映射） |
| `products.json` | 產品資料（build_search.py 生成，勿手動編輯） |
| `search.html` | 主搜尋頁面 |
| `鋁線槽系列規格表.xlsx` | 資料來源（勿修改） |
| `build_excel_aluminum.py` | 原始 Excel 生成腳本（資料維護用） |
| `extract_images.py` | PDF 圖片擷取腳本 |
| `圖庫/` | 產品圖片（子目錄名稱 = 料號） |
| `1. 鋁線槽系列P2~P15.pdf` | 廠商原始型錄（勿修改） |
| `PLAN.md` | Codex 詳細執行規格 |

## 目前進度

最後更新：2026-05-08

- 已完成 `build_search.py`
  - 讀取 `鋁線槽系列規格表.xlsx`
  - 產生 `products.json`
  - 掃描 `圖庫/` 並建立主圖與 gallery 路徑
  - 目前輸出 92 筆產品資料
- 已完成 `products.json`
  - `meta.total` 與 `products.length` 皆為 92
  - `HB128.hasImage` 為 `true`
  - `HB128.images.main` 為 `圖庫/HB128/外觀1.jpeg`
  - `HB128.sceneList` 已正常拆分
- 已完成 `search.html`
  - 原生 HTML/CSS/JS，無框架
  - 支援即時搜尋、安裝方式篩選、安裝位置篩選、適用場景篩選
  - 支援產品卡片、無圖佔位符、產品詳細 Modal、圖片 gallery
  - 支援 ESC、背景點擊、關閉鈕關閉 Modal
  - 響應式版面：手機 1 欄、平板 2 欄、桌機 3 欄

## 驗收紀錄

- `python3 build_search.py`：通過，輸出 `products.json` 共 92 筆
- `python3 -m py_compile build_search.py`：通過
- 瀏覽器自動化驗收：通過
  - 初始載入：92 張產品卡
  - 搜尋 `HB128`：顯示 `HB128`、`HB128X`
  - 篩選安裝方式 `明裝`：49 筆，結果皆符合
  - Modal：可開啟，ESC 可關閉
  - 無結果狀態：正常顯示並可清除
  - 1200px：3 欄
  - 768px：2 欄
  - 375px：1 欄，無水平溢出
  - console errors：0

## 執行方式

```bash
# 1. 安裝依賴（首次）
pip install openpyxl

# 2. 生成資料
python3 build_search.py

# 3. 啟動本機伺服器（避免 fetch CORS）
python3 -m http.server 8080

# 4. 開啟瀏覽器
# http://localhost:8080/search.html
```

## 資料更新流程

1. 在 `build_excel_aluminum.py` 的 DATA 陣列中新增/修改產品
2. 執行 `python3 build_excel_aluminum.py` 重產 xlsx
3. 執行 `python3 build_search.py` 重產 products.json
4. 重新整理 search.html 即可

## 設計原則

- 字體最小 17px，定價 20px，高對比（背景 #f5f7fa，文字 #1a1a1a）
- 觸控區域最小 48px（符合 WCAG 2.5.5）
- 明亮主題，適合戶外與各種光線環境
- 手機 1 欄，平板 2 欄，桌機 3 欄
