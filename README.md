# BIM 混凝土成本動態估算系統

選手大樓專題：用 Autodesk Forma Takeoff 算出結構混凝土體積，搭配政府公開的最新單價，自動估算混凝土成本。

🌐 **網站：** https://nkhs-arch.github.io/bim-concrete-cost/

## 運作方式

```
Forma Takeoff 算出體積（樓板、柱、梁）
   │  模型改版時匯出 Excel，覆蓋 data/quantities.xlsx
   ▼
scripts/build_data.py（GitHub Actions 每週一 08:00 自動執行）
   │  體積 × 最新單價
   │  最新單價 = 工程會基準價 × 主計總處物價指數調整
   ▼
data.json → index.html 網頁儀表板
```

## 單價來源（政府開放資料）

| 資料 | 提供機關 | 更新頻率 |
|---|---|---|
| [大宗資材之混凝土價格調查](https://data.gov.tw/dataset/6819)（北區各強度單價） | 行政院公共工程委員會 | 每月 |
| [營造工程物價指數](https://data.gov.tw/dataset/8246)「水泥及其製品類」 | 行政院主計總處 | 每月 |

最新單價 = 工程會基準價 × 最新月份指數 ÷ 基準價月份指數

## 碳排放（6D 永續）

| 資料 | 提供機關 | 更新頻率 |
|---|---|---|
| [碳足跡排放係數](https://data.gov.tw/dataset/28176)（預拌混凝土各強度，kgCO₂e/m³） | 環境部 | 不定期 |

蘊含碳 = 混凝土體積 × 係數。同一強度有多筆係數時，取最新公告年份（同年多筆取平均），水中、自充填、加飛灰爐石等特殊配比不採用。評估方式參照建研所《低碳（低蘊含碳）建築評估手冊》（LEBR）。目前只算混凝土，鋼筋重量完成後再加入。

> 工程會網站無法從 GitHub 的海外主機連線，所以自動更新時沿用 `data/price_cache.json` 裡的基準價，只更新物價指數。要更新基準價時，在自己電腦執行一次 `python scripts/build_data.py` 再上傳即可。

## 常見操作

### 模型改版了
1. Forma Takeoff →「混凝土體積計算」→ 匯出 Excel
2. 改名為 `quantities.xlsx`，覆蓋 `data/quantities.xlsx`
3. 修改 `data/model_info.json` 的 `note`（例如「第二版」）
4. 上傳後會自動重算並建立新版本（V2、V3…）

### 設定混凝土強度
Takeoff 匯出檔沒有強度資料，預設全部用 280 kgf/cm²。要指定時修改 `data/model_info.json`：

```json
"grades": { "columns": "350", "beams": "280", "floors": "280" }
```

### 手動重新估算
GitHub → **Actions** → **BIM 混凝土成本自動估算** → **Run workflow**

### 產生離線版 / PDF
```bash
python scripts/export_offline.py
```
會產生可直接雙擊開啟的 `BIM混凝土成本儀表板_離線版_<日期>.html`。

## 檔案說明

| 路徑 | 說明 |
|---|---|
| `index.html` | 網頁儀表板 |
| `data.json` | 計算結果（自動產生，不要手動改） |
| `data/quantities.xlsx` | Forma Takeoff 匯出的體積 |
| `data/model_info.json` | 模型名稱、版本備註、工地區域、混凝土強度 |
| `data/versions.json` | 各版本體積紀錄（自動維護） |
| `data/price_cache.json` | 最近一次下載的單價資料（自動維護） |
| `data/carbon_cache.json` | 最近一次下載的碳排係數（自動維護） |
| `scripts/build_data.py` | 主程式：讀體積、抓單價、算成本 |
| `scripts/export_offline.py` | 匯出離線版 HTML |
| `scripts/certs/` | 補上主計總處網站缺少的 TWCA 中繼憑證 |
| `.github/workflows/bim_auto_cost.yml` | 自動排程設定 |
