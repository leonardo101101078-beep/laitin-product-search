# PLAN.md — 內部業務查詢系統 Codex 執行規格

## 專案概述

為萊廷照明業務人員建立純前端產品查詢網頁。資料來源為 Excel 規格表（92 筆產品），透過 Python 腳本轉換成 JSON，再由單一 HTML 頁面讀取並提供即時搜尋與篩選功能。

**需要建立的檔案：**
1. `build_search.py` — 讀取 xlsx，掃描圖庫，輸出 products.json
2. `products.json` — 執行上述腳本後自動生成
3. `search.html` — 主搜尋頁面

**不得修改的檔案：**
- `鋁線槽系列規格表.xlsx`
- `build_excel_aluminum.py`
- `extract_images.py`
- `1. 鋁線槽系列P2~P15.pdf`

---

## Task 1：build_search.py

### 依賴

```python
import openpyxl, json, os, re
from datetime import date
```

只使用 openpyxl（已在現有腳本中使用），不引入 pandas。

### 執行邏輯

```
讀取 鋁線槽系列規格表.xlsx 的第一個工作表
    ↓
跳過第 1 列（標頭列）
    ↓
掃描 圖庫/ 目錄，建立 image_map: { sku: { main, gallery } }
    ↓
逐列（row 2 起）轉換為 product dict
    ↓
寫入 products.json
```

### Excel 欄位對應

工作表名稱：`鋁線槽系列規格表`（若不存在則取第一個工作表）

| 欄索引（1-based） | Excel 標頭 | JSON 欄位 | 處理方式 |
|---|---|---|---|
| 1 | 序號 | — | 跳過 |
| 2 | 型錄頁 | `catalogPage` | str 或 null |
| 3 | 大類 | `category` | str 或 null |
| 4 | 料號 | `sku` | str（同時設為 id） |
| 5 | 品名 | `name` | str |
| 6 | 外寬(mm) | `widthMm` | float 或 null |
| 7 | 外高(mm) | `heightMm` | float 或 null |
| 8 | 電路槽寬(mm) | `channelWidthMm` | float 或 null |
| 9 | 板厚(mm) | `thicknessMm` | float 或 null |
| 10 | 安裝方式 | `mountMethod` | str 或 null |
| 11 | 安裝位置 | `mountLocation` | str 或 null |
| 12 | 出光方向 | `lightDirection` | str 或 null |
| 13 | 電壓 | `voltage` | str 或 null |
| 14 | 色溫選項 | `colorTemp` + `colorTempList` | 見下方 |
| 15 | IP防護等級 | `ipRating` | str 或 null |
| 16 | 適用場景 | `scenes` + `sceneList` | 見下方 |
| 17 | 特殊功能 | `features` | str 或 null |
| 18 | 可選顏色 | `colors` | str 或 null |
| 19 | 對應配件料號 | `relatedParts` | str 或 null |
| 20 | 定價(NT$) | `price` | 見下方 |
| 21 | 計價單位 | `priceUnit` | str 或 null |
| 22 | 備註 | `notes` | str 或 null |

### 欄位處理細則

**通用 None/空值處理：**
```python
def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    return None if s == '' else s
```

**數值欄位（widthMm, heightMm, channelWidthMm, thicknessMm）：**
```python
def to_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None
```

**定價（price）：**
- 若 cell 值是數字（int/float）→ 轉 float
- 若是字串（如 `"150 / 160"`）→ 保留字串原值
- None → null

```python
def parse_price(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    return None if s == '' else s
```

**逗號分隔欄位（colorTemp/colorTempList 和 scenes/sceneList）：**
```python
def split_tags(v):
    s = clean(v)
    if not s or s == 'N/A':
        return (None, [])
    items = [x.strip() for x in s.split(',') if x.strip() and x.strip() != 'N/A']
    return (s, items)
```
- `colorTemp` = 原始字串，`colorTempList` = 拆分後的陣列
- `scenes` = 原始字串，`sceneList` = 拆分後的陣列

**「N/A」值：** 保留為字串 `"N/A"`（前端會過濾顯示），不轉成 null

### 圖庫掃描邏輯

```python
def build_image_map(gallery_dir='圖庫'):
    image_map = {}
    if not os.path.isdir(gallery_dir):
        return image_map
    
    for sku_folder in os.listdir(gallery_dir):
        folder_path = os.path.join(gallery_dir, sku_folder)
        if not os.path.isdir(folder_path):
            continue
        
        files = [f for f in os.listdir(folder_path)
                 if f.lower().endswith(('.jpg', '.jpeg', '.png'))
                 and not f.startswith('.')]
        
        if not files:
            continue
        
        # main 圖優先順序
        main = None
        for priority in ['外觀1.jpeg', '外觀1.jpg', '外觀1.png']:
            if priority in files:
                main = priority
                break
        if not main:
            # 找任意 外觀*.jpeg/jpg
            for f in sorted(files):
                if f.startswith('外觀'):
                    main = f
                    break
        if not main:
            main = sorted(files)[0]
        
        gallery = [f for f in sorted(files) if not f.startswith('全頁參考') and not f.startswith('提取_')]
        # 最多取 5 張
        gallery = gallery[:5]
        
        image_map[sku_folder] = {
            'main': f'{gallery_dir}/{sku_folder}/{main}',
            'gallery': [f'{gallery_dir}/{sku_folder}/{f}' for f in gallery]
        }
    
    return image_map
```

### 輸出 products.json

```python
def row_to_product(row_values, image_map):
    # row_values[0] = 序號（跳過），row_values[1] = 型錄頁，依此類推
    # 注意：openpyxl row 取值後轉 list，index 0 = 序號
    
    sku = clean(row_values[3])  # 欄 4（index 3）
    if not sku:
        return None  # 跳過空列
    
    color_temp_raw, color_temp_list = split_tags(row_values[13])
    scenes_raw, scene_list = split_tags(row_values[15])
    
    img = image_map.get(sku, {})
    has_image = bool(img.get('main'))
    
    return {
        'sku': sku,
        'catalogPage': clean(row_values[1]),
        'category': clean(row_values[2]),
        'name': clean(row_values[4]),
        'widthMm': to_float(row_values[5]),
        'heightMm': to_float(row_values[6]),
        'channelWidthMm': to_float(row_values[7]),
        'thicknessMm': to_float(row_values[8]),
        'mountMethod': clean(row_values[9]),
        'mountLocation': clean(row_values[10]),
        'lightDirection': clean(row_values[11]),
        'voltage': clean(row_values[12]),
        'colorTemp': color_temp_raw,
        'colorTempList': color_temp_list,
        'ipRating': clean(row_values[14]),
        'scenes': scenes_raw,
        'sceneList': scene_list,
        'features': clean(row_values[16]),
        'colors': clean(row_values[17]),
        'relatedParts': clean(row_values[18]),
        'price': parse_price(row_values[19]),
        'priceUnit': clean(row_values[20]),
        'notes': clean(row_values[21]),
        'hasImage': has_image,
        'images': img if has_image else {'main': None, 'gallery': []}
    }

def main():
    wb = openpyxl.load_workbook('鋁線槽系列規格表.xlsx', data_only=True)
    ws = wb.worksheets[0]
    
    image_map = build_image_map('圖庫')
    
    products = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        p = row_to_product(list(row), image_map)
        if p:
            products.append(p)
    
    output = {
        'meta': {
            'generated': str(date.today()),
            'total': len(products),
            'source': '鋁線槽系列規格表.xlsx'
        },
        'products': products
    }
    
    with open('products.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f'✓ 輸出 products.json，共 {len(products)} 筆')

if __name__ == '__main__':
    main()
```

---

## Task 2：search.html

單一 HTML 檔案，所有 CSS 和 JS 內嵌（`<style>` + `<script>`），透過 `fetch('./products.json')` 讀取資料。

### HTML 骨架

```html
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>萊廷照明 — 業務查詢系統</title>
  <style>/* 所有 CSS 在此 */</style>
</head>
<body>
  <header id="site-header">
    <div class="header-inner">
      <h1>萊廷照明</h1>
      <span class="header-sub">業務查詢系統</span>
    </div>
  </header>

  <div id="controls" class="controls-bar">
    <div class="search-row">
      <div class="search-wrap">
        <input type="search" id="search-input" placeholder="輸入料號、品名、場景…" autocomplete="off">
        <button id="clear-search" aria-label="清除">✕</button>
      </div>
    </div>
    <div class="filter-row">
      <select id="filter-mount-method" class="filter-select" data-filter="mountMethod">
        <option value="">安裝方式（全部）</option>
      </select>
      <select id="filter-mount-location" class="filter-select" data-filter="mountLocation">
        <option value="">安裝位置（全部）</option>
      </select>
      <select id="filter-scene" class="filter-select" data-filter="scene">
        <option value="">適用場景（全部）</option>
      </select>
      <button id="clear-filters">清除篩選</button>
    </div>
    <div id="result-count" class="result-count"></div>
  </div>

  <main>
    <div id="product-grid" class="product-grid"></div>
    <div id="no-result" class="no-result" hidden>
      <p>找不到符合條件的產品</p>
      <button onclick="clearAll()">清除所有篩選</button>
    </div>
  </main>

  <!-- Modal -->
  <div id="modal-overlay" class="modal-overlay" hidden>
    <div id="modal" class="modal" role="dialog" aria-modal="true">
      <button id="modal-close" class="modal-close" aria-label="關閉">✕</button>
      <div id="modal-content"></div>
    </div>
  </div>

  <script>/* 所有 JS 在此 */</script>
</body>
</html>
```

---

### CSS 規格

#### CSS 自訂屬性

```css
:root {
  --font-base: 17px;
  --font-sm: 14px;
  --font-lg: 20px;
  --font-xl: 22px;
  --touch: 48px;

  --bg: #f5f7fa;
  --surface: #ffffff;
  --border: #e0e4ea;
  --text: #1a1a1a;
  --text-muted: #555;
  --primary: #1a5fa8;
  --primary-dark: #134b88;
  --price: #0a7c30;
  --tag-bg: #e8f0fb;
  --tag-text: #1a5fa8;

  --radius: 12px;
  --shadow: 0 2px 8px rgba(0,0,0,0.09);
  --shadow-hover: 0 4px 16px rgba(0,0,0,0.15);
  --header-h: 56px;
  --controls-h: auto;
}
```

#### 整體 reset 與 body

```css
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: "微軟正黑體", "Microsoft JhengHei", "PingFang TC", "Noto Sans TC", sans-serif;
  font-size: var(--font-base);
  background: var(--bg);
  color: var(--text);
  line-height: 1.5;
}
```

#### Header

```css
#site-header {
  position: fixed;
  top: 0; left: 0; right: 0;
  height: var(--header-h);
  background: var(--primary);
  color: #fff;
  z-index: 100;
  display: flex;
  align-items: center;
  padding: 0 16px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}
.header-inner { display: flex; align-items: baseline; gap: 10px; }
#site-header h1 { font-size: 20px; font-weight: 700; }
.header-sub { font-size: 13px; opacity: 0.8; }
```

#### 篩選控制列

```css
.controls-bar {
  position: sticky;
  top: var(--header-h);
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  z-index: 90;
  padding: 12px 16px 8px;
}
.search-row { margin-bottom: 8px; }
.search-wrap {
  position: relative;
  max-width: 480px;
}
#search-input {
  width: 100%;
  height: var(--touch);
  padding: 0 44px 0 16px;
  font-size: var(--font-base);
  border: 2px solid var(--border);
  border-radius: 8px;
  outline: none;
  transition: border-color 0.2s;
}
#search-input:focus { border-color: var(--primary); }
#clear-search {
  position: absolute;
  right: 8px; top: 50%;
  transform: translateY(-50%);
  background: none; border: none;
  cursor: pointer; font-size: 18px;
  color: var(--text-muted);
  width: 32px; height: 32px;
  display: flex; align-items: center; justify-content: center;
}

.filter-row {
  display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
}
.filter-select {
  height: var(--touch);
  padding: 0 12px;
  font-size: var(--font-sm);
  border: 1.5px solid var(--border);
  border-radius: 8px;
  background: var(--surface);
  color: var(--text);
  cursor: pointer;
  min-width: 130px;
}
.filter-select:focus { border-color: var(--primary); outline: none; }

#clear-filters {
  height: var(--touch);
  padding: 0 16px;
  font-size: var(--font-sm);
  border: 1.5px solid var(--border);
  border-radius: 8px;
  background: none;
  color: var(--text-muted);
  cursor: pointer;
}
#clear-filters:hover { background: var(--bg); }

.result-count {
  font-size: var(--font-sm);
  color: var(--text-muted);
  margin-top: 8px;
}
```

#### 卡片網格

```css
main {
  margin-top: calc(var(--header-h) + 20px); /* header 高度 + 初始間距 */
  padding: 16px;
  /* 注意：controls-bar 是 sticky，不佔 margin */
}

.product-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 16px;
}

@media (min-width: 600px) {
  .product-grid { grid-template-columns: repeat(2, 1fr); }
}

@media (min-width: 900px) {
  .product-grid {
    grid-template-columns: repeat(3, 1fr);
    max-width: 1200px;
    margin: 0 auto;
  }
  main { padding: 24px; }
}
```

#### 產品卡片

```css
.card {
  background: var(--surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
  cursor: pointer;
  transition: box-shadow 0.2s, transform 0.2s;
  border: 1.5px solid transparent;
}
.card:hover, .card:focus {
  box-shadow: var(--shadow-hover);
  transform: translateY(-2px);
  border-color: var(--primary);
  outline: none;
}
.card:active { transform: translateY(0); }

.card-img {
  width: 100%;
  height: 180px;
  overflow: hidden;
  background: var(--bg);
}
.card-img img {
  width: 100%; height: 100%;
  object-fit: cover;
  display: block;
}

/* 佔位符 */
.img-placeholder {
  width: 100%; height: 100%;
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  gap: 8px;
  font-size: var(--font-sm);
}
.placeholder-sku {
  font-size: 18px;
  font-weight: 700;
  letter-spacing: 1px;
}
.placeholder-cat {
  font-size: 12px;
  opacity: 0.7;
}

.card-body {
  padding: 14px 16px 16px;
  display: flex; flex-direction: column; gap: 6px;
}
.card-sku {
  font-size: var(--font-sm);
  color: var(--text-muted);
  font-weight: 600;
}
.card-name {
  font-size: 18px;
  font-weight: 700;
  line-height: 1.3;
  color: var(--text);
}
.card-size {
  font-size: var(--font-sm);
  color: var(--text-muted);
}
.card-tags {
  display: flex; flex-wrap: wrap; gap: 4px;
  margin: 2px 0;
}
.tag {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 100px;
  font-size: 12px;
  background: var(--tag-bg);
  color: var(--tag-text);
}
.card-meta {
  font-size: var(--font-sm);
  color: var(--text-muted);
}
.dot { margin: 0 4px; }
.card-price {
  font-size: var(--font-lg);
  font-weight: 700;
  color: var(--price);
  margin-top: 4px;
}
```

#### 無結果提示

```css
.no-result {
  text-align: center;
  padding: 60px 16px;
  color: var(--text-muted);
}
.no-result p { font-size: 18px; margin-bottom: 16px; }
.no-result button {
  padding: 12px 24px;
  font-size: var(--font-base);
  background: var(--primary);
  color: #fff;
  border: none;
  border-radius: 8px;
  cursor: pointer;
}
```

#### Modal

```css
.modal-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.5);
  z-index: 200;
  display: flex;
  align-items: flex-end;
  justify-content: center;
  padding: 0;
}
@media (min-width: 600px) {
  .modal-overlay {
    align-items: center;
    padding: 24px;
  }
}

.modal {
  background: var(--surface);
  border-radius: var(--radius) var(--radius) 0 0;
  width: 100%;
  max-height: 90vh;
  overflow-y: auto;
  position: relative;
  padding: 0;
}
@media (min-width: 600px) {
  .modal {
    border-radius: var(--radius);
    max-width: 600px;
    max-height: 85vh;
  }
}

.modal-close {
  position: sticky;
  top: 0;
  float: right;
  margin: 12px 12px 0 0;
  width: 36px; height: 36px;
  border-radius: 50%;
  border: none;
  background: var(--bg);
  font-size: 18px;
  cursor: pointer;
  z-index: 1;
  display: flex; align-items: center; justify-content: center;
}

.modal-header {
  padding: 16px 56px 16px 20px;
  border-bottom: 1px solid var(--border);
}
.modal-sku { font-size: var(--font-sm); color: var(--text-muted); }
.modal-name { font-size: 20px; font-weight: 700; margin-top: 2px; }

/* 圖片輪播 */
.modal-gallery {
  background: #000;
}
.modal-main-img {
  width: 100%;
  height: 240px;
  object-fit: contain;
  display: block;
}
.modal-thumbs {
  display: flex;
  gap: 4px;
  padding: 8px;
  overflow-x: auto;
  background: #111;
}
.modal-thumb {
  width: 60px; height: 60px;
  object-fit: cover;
  border-radius: 4px;
  cursor: pointer;
  opacity: 0.6;
  border: 2px solid transparent;
  flex-shrink: 0;
}
.modal-thumb.active, .modal-thumb:hover { opacity: 1; border-color: #fff; }

/* 規格表 */
.modal-specs { padding: 16px 20px; }
.specs-table {
  width: 100%;
  border-collapse: collapse;
  font-size: var(--font-base);
}
.specs-table th {
  text-align: left;
  padding: 8px 12px 8px 0;
  color: var(--text-muted);
  font-weight: 500;
  white-space: nowrap;
  vertical-align: top;
  width: 100px;
}
.specs-table td {
  padding: 8px 0;
  color: var(--text);
  border-bottom: 1px solid var(--border);
}
.modal-price {
  padding: 16px 20px;
  border-top: 2px solid var(--border);
  font-size: 24px;
  font-weight: 700;
  color: var(--price);
}
.modal-notes {
  padding: 8px 20px 20px;
  font-size: var(--font-sm);
  color: var(--text-muted);
}
```

---

### JavaScript 規格

#### 全域狀態

```js
let allProducts = [];   // 從 products.json 載入後不變動
let filtered = [];      // 當前篩選結果
let savedScrollY = 0;   // modal 開啟前的捲動位置

const state = {
  query: '',
  mountMethod: '',
  mountLocation: '',
  scene: ''
};
```

#### 初始化

```js
async function init() {
  try {
    const res = await fetch('./products.json');
    const data = await res.json();
    allProducts = data.products;
    buildFilterOptions();
    filter();
  } catch (e) {
    document.getElementById('product-grid').innerHTML =
      '<p style="padding:32px;color:#c00">無法載入資料，請確認已執行 build_search.py 並透過本機伺服器開啟頁面。</p>';
  }
}

document.addEventListener('DOMContentLoaded', init);
```

#### 動態生成下拉選項

```js
function buildFilterOptions() {
  // 安裝方式：按 "/" 拆分，取所有子值
  const methods = [...new Set(
    allProducts
      .map(p => p.mountMethod)
      .filter(v => v && v !== 'N/A')
      .flatMap(v => v.split('/').map(s => s.trim()))
  )].sort();

  // 安裝位置：按 "/" 或 "、" 拆分
  const locations = [...new Set(
    allProducts
      .map(p => p.mountLocation)
      .filter(v => v && v !== 'N/A')
      .flatMap(v => v.split(/[\/、]/).map(s => s.trim()))
      .filter(Boolean)
  )].sort();

  // 適用場景：用 sceneList
  const scenes = [...new Set(
    allProducts.flatMap(p => p.sceneList)
  )].filter(Boolean).sort();

  populateSelect('filter-mount-method', methods, '安裝方式（全部）');
  populateSelect('filter-mount-location', locations, '安裝位置（全部）');
  populateSelect('filter-scene', scenes, '適用場景（全部）');
}

function populateSelect(id, values, placeholder) {
  const sel = document.getElementById(id);
  sel.innerHTML = `<option value="">${placeholder}</option>` +
    values.map(v => `<option value="${v}">${v}</option>`).join('');
}
```

#### 搜尋與篩選

```js
function matchSearch(p, q) {
  if (!q) return true;
  const haystack = [
    p.sku, p.name, p.scenes, p.category,
    p.features, p.mountLocation, p.notes
  ].filter(Boolean).join(' ').toLowerCase();
  return haystack.includes(q.toLowerCase().trim());
}

function matchFilters(p) {
  if (state.mountMethod) {
    const m = p.mountMethod || '';
    // 處理複合值如 "明裝/嵌入"：拆分後看是否含選定值
    const parts = m.split('/').map(s => s.trim());
    if (!parts.includes(state.mountMethod)) return false;
  }
  if (state.mountLocation) {
    const l = p.mountLocation || '';
    const parts = l.split(/[\/、]/).map(s => s.trim());
    if (!parts.includes(state.mountLocation)) return false;
  }
  if (state.scene) {
    if (!p.sceneList.includes(state.scene)) return false;
  }
  return true;
}

function filter() {
  filtered = allProducts.filter(p => matchSearch(p, state.query) && matchFilters(p));
  const count = document.getElementById('result-count');
  count.textContent = `共 ${filtered.length} 項產品`;
  renderGrid();
}
```

#### 渲染卡片

```js
// 大類對應佔位符背景色
const CAT_COLORS = {
  '天花回光槽': { bg: '#dbeafe', text: '#1e3a5f' },
  '超薄斜角懸浮鋁槽': { bg: '#ede9fe', text: '#3b1f7a' },
  '嵌入斜照偏光': { bg: '#fce7f3', text: '#6b1342' },
  '斜照懸浮吊頂型材': { bg: '#d1fae5', text: '#064e3b' },
  '嵌入式無邊框偏光洗牆': { bg: '#fef9c3', text: '#713f12' },
  '弧面雙向回光槽': { bg: '#ffedd5', text: '#7c2d12' },
  '防眩弧面懸浮鋁槽': { bg: '#e0f2fe', text: '#0c4a6e' },
  '極窄系列': { bg: '#f0fdf4', text: '#14532d' },
  '單邊線型燈': { bg: '#fdf4ff', text: '#581c87' },
  '明裝免吊頂頂角': { bg: '#fff7ed', text: '#7c2d12' },
  '雙路發光牆沿線型燈': { bg: '#f8f8f8', text: '#111' },
  '頂角線型燈': { bg: '#ecfdf5', text: '#064e3b' },
  '頂角線條燈槽': { bg: '#fef3c7', text: '#78350f' },
  '向上發光回光槽': { bg: '#e0f7fa', text: '#006064' },
  '牆面腰際線型燈': { bg: '#fce4ec', text: '#880e4f' },
  '常用型材': { bg: '#f5f5f5', text: '#333' },
  '配件': { bg: '#fffde7', text: '#5d4037' },
};

function getPlaceholder(p) {
  const c = CAT_COLORS[p.category] || { bg: '#f1f5f9', text: '#374151' };
  return `<div class="img-placeholder" style="background:${c.bg};color:${c.text}">
    <svg width="40" height="40" viewBox="0 0 40 40" fill="none">
      <rect x="4" y="15" width="32" height="8" rx="4" fill="currentColor" opacity="0.35"/>
      <rect x="8" y="9" width="24" height="4" rx="2" fill="currentColor" opacity="0.55"/>
      <circle cx="20" cy="30" r="3" fill="currentColor" opacity="0.25"/>
    </svg>
    <span class="placeholder-sku">${p.sku.slice(0, 6)}</span>
    <span class="placeholder-cat">${p.category || ''}</span>
  </div>`;
}

function renderImg(p) {
  if (!p.hasImage) return getPlaceholder(p);
  return `<img src="${p.images.main}" alt="${p.name}" loading="lazy"
    onerror="this.parentElement.innerHTML=getPlaceholderById('${p.sku}')">`;
}

// 供 onerror 調用（onerror 在 inline 上下文中只能用全域函式）
window.getPlaceholderById = function(sku) {
  const p = allProducts.find(x => x.sku === sku);
  return p ? getPlaceholder(p) : '';
};

function formatSize(p) {
  const parts = [];
  if (p.widthMm) parts.push(`${p.widthMm}`);
  if (p.heightMm) parts.push(`${p.heightMm}`);
  if (parts.length === 0) return '';
  return `${parts.join(' × ')} mm`;
}

function formatPrice(p) {
  if (!p.price) return '';
  const amount = typeof p.price === 'number'
    ? `NT$${p.price.toLocaleString()}`
    : `NT$${p.price}`;
  const unit = p.priceUnit || '';
  return `${amount} ${unit}`.trim();
}

function renderCard(p) {
  const size = formatSize(p);
  const price = formatPrice(p);
  const tags = (p.sceneList || []).slice(0, 3)
    .map(s => `<span class="tag">${s}</span>`).join('');
  const meta = [p.mountMethod, p.mountLocation]
    .filter(v => v && v !== 'N/A').join(' · ') || '';

  return `<article class="card" tabindex="0" data-sku="${p.sku}" role="button" aria-label="${p.sku} ${p.name}">
    <div class="card-img">${renderImg(p)}</div>
    <div class="card-body">
      <div class="card-sku">${p.sku}</div>
      <div class="card-name">${p.name}</div>
      ${size ? `<div class="card-size">${size}</div>` : ''}
      ${tags ? `<div class="card-tags">${tags}</div>` : ''}
      ${meta ? `<div class="card-meta">${meta}</div>` : ''}
      ${price ? `<div class="card-price">${price}</div>` : ''}
    </div>
  </article>`;
}

function renderGrid() {
  const grid = document.getElementById('product-grid');
  const noResult = document.getElementById('no-result');

  if (filtered.length === 0) {
    grid.innerHTML = '';
    noResult.hidden = false;
  } else {
    noResult.hidden = true;
    grid.innerHTML = filtered.map(renderCard).join('');
  }
}
```

#### Modal

```js
function openModal(sku) {
  const p = allProducts.find(x => x.sku === sku);
  if (!p) return;

  savedScrollY = window.scrollY;

  const overlay = document.getElementById('modal-overlay');
  const content = document.getElementById('modal-content');

  // 圖片區
  let galleryHTML = '';
  if (p.hasImage && p.images.gallery.length > 0) {
    const thumbs = p.images.gallery
      .map((src, i) => `<img src="${src}" class="modal-thumb${i === 0 ? ' active' : ''}" data-idx="${i}" loading="lazy" alt="">`)
      .join('');
    galleryHTML = `
      <div class="modal-gallery">
        <img id="modal-main-img" class="modal-main-img" src="${p.images.gallery[0]}" alt="${p.name}">
        ${p.images.gallery.length > 1 ? `<div class="modal-thumbs">${thumbs}</div>` : ''}
      </div>`;
  }

  // 規格列（null/N/A/空值跳過）
  function row(label, value) {
    if (!value || value === 'N/A') return '';
    return `<tr><th>${label}</th><td>${value}</td></tr>`;
  }
  function sizeStr() {
    const parts = [];
    if (p.widthMm) parts.push(`外寬 ${p.widthMm} mm`);
    if (p.heightMm) parts.push(`外高 ${p.heightMm} mm`);
    if (p.channelWidthMm) parts.push(`槽寬 ${p.channelWidthMm} mm`);
    if (p.thicknessMm) parts.push(`板厚 ${p.thicknessMm} mm`);
    return parts.join('　');
  }

  const specsHTML = `
    <div class="modal-specs">
      <table class="specs-table">
        ${row('型錄頁', p.catalogPage)}
        ${row('大類', p.category)}
        ${sizeStr() ? `<tr><th>尺寸</th><td>${sizeStr()}</td></tr>` : ''}
        ${row('安裝方式', p.mountMethod)}
        ${row('安裝位置', p.mountLocation)}
        ${row('出光方向', p.lightDirection)}
        ${row('電壓', p.voltage)}
        ${row('色溫', p.colorTemp)}
        ${row('IP防護', p.ipRating)}
        ${row('適用場景', p.scenes)}
        ${row('特殊功能', p.features)}
        ${row('可選顏色', p.colors)}
        ${row('對應配件', p.relatedParts)}
      </table>
    </div>`;

  const priceStr = formatPrice(p);
  const priceHTML = priceStr ? `<div class="modal-price">${priceStr}</div>` : '';
  const notesHTML = p.notes ? `<div class="modal-notes">備註：${p.notes}</div>` : '';

  content.innerHTML = `
    <div class="modal-header">
      <div class="modal-sku">${p.sku}</div>
      <div class="modal-name">${p.name}</div>
    </div>
    ${galleryHTML}
    ${specsHTML}
    ${priceHTML}
    ${notesHTML}`;

  // 縮圖切換事件
  content.querySelectorAll('.modal-thumb').forEach(thumb => {
    thumb.addEventListener('click', () => {
      const src = thumb.src;
      document.getElementById('modal-main-img').src = src;
      content.querySelectorAll('.modal-thumb').forEach(t => t.classList.remove('active'));
      thumb.classList.add('active');
    });
  });

  overlay.hidden = false;
  document.body.style.overflow = 'hidden';
  document.getElementById('modal-close').focus();
}

function closeModal() {
  document.getElementById('modal-overlay').hidden = true;
  document.body.style.overflow = '';
  window.scrollTo(0, savedScrollY);
}
```

#### 事件綁定

```js
function bindEvents() {
  // 搜尋
  document.getElementById('search-input').addEventListener('input', e => {
    state.query = e.target.value;
    filter();
  });
  document.getElementById('clear-search').addEventListener('click', () => {
    document.getElementById('search-input').value = '';
    state.query = '';
    filter();
  });

  // 篩選下拉
  document.querySelectorAll('.filter-select').forEach(sel => {
    sel.addEventListener('change', e => {
      state[e.target.dataset.filter] = e.target.value;
      filter();
    });
  });

  // 清除篩選
  document.getElementById('clear-filters').addEventListener('click', clearAll);

  // 卡片點擊（event delegation）
  document.getElementById('product-grid').addEventListener('click', e => {
    const card = e.target.closest('.card');
    if (card) openModal(card.dataset.sku);
  });
  document.getElementById('product-grid').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      const card = e.target.closest('.card');
      if (card) { e.preventDefault(); openModal(card.dataset.sku); }
    }
  });

  // Modal 關閉
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeModal();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !document.getElementById('modal-overlay').hidden) {
      closeModal();
    }
  });
}

function clearAll() {
  document.getElementById('search-input').value = '';
  document.querySelectorAll('.filter-select').forEach(s => s.value = '');
  state.query = '';
  state.mountMethod = '';
  state.mountLocation = '';
  state.scene = '';
  filter();
}

document.addEventListener('DOMContentLoaded', () => {
  init();
  bindEvents();
});
```

**注意：** 將 `init()` 和 `bindEvents()` 的呼叫統一放在 `DOMContentLoaded` 中，不要重複呼叫 `init()`。

---

## 驗收標準

執行以下指令後，逐項確認：

```bash
cd 內部業務查詢系統
python3 build_search.py
python3 -m http.server 8080
# 開啟 http://localhost:8080/search.html
```

| # | 測試項目 | 預期結果 |
|---|---|---|
| 1 | `python3 build_search.py` 執行成功 | 輸出「✓ 輸出 products.json，共 N 筆」，N >= 90 |
| 2 | products.json 中 HB128 的 hasImage | `true`，images.main 為 `圖庫/HB128/外觀1.jpeg` |
| 3 | products.json 中有 sceneList 陣列 | `["客廳", "臥室", "餐廳", ...]` 非空 |
| 4 | 頁面載入 | Header 顯示「萊廷照明 業務查詢系統」，卡片網格出現 |
| 5 | 搜尋「HB128」 | 立即篩選，只顯示 HB128 相關卡片 |
| 6 | 篩選「安裝方式」選「明裝」 | 所有卡片 mountMethod 含「明裝」 |
| 7 | 點擊卡片 | Modal 開啟，顯示完整規格 |
| 8 | ESC 鍵 | Modal 關閉，頁面恢復原 scroll 位置 |
| 9 | 375px 寬（Chrome DevTools 手機模式） | 1 欄，字體清晰，不橫向溢出 |
| 10 | 768px 寬 | 2 欄 |
| 11 | 無圖產品（非 HB128/HB128X/HB381A） | 顯示彩色佔位符，無 JS 錯誤 |
| 12 | 搜尋無結果 | 顯示「找不到符合條件的產品」提示與清除按鈕 |
