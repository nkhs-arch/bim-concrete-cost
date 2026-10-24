# BIM 混凝土成本動態估算系統

選手大樓專題：用 Autodesk Forma Takeoff 或 Revit 明細表算出結構混凝土體積，搭配政府公開的最新單價，自動估算混凝土與鋼筋成本。

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

蘊含碳 = 混凝土體積 × 係數。同一強度有多筆係數時，取最新公告年份（同年多筆取平均），水中、自充填、加飛灰爐石等特殊配比不採用。評估方式參照建研所《低碳（低蘊含碳）建築評估手冊》（LEBR）。目前只算混凝土，尚未含鋼筋碳排。

## 鋼筋重量與成本（練習假設）

鋼筋重量 = 各類型構件體積 × 鋼筋比（kg/m³）× (1 + 損耗率)

- **鋼筋比**寫在 `data/rebar.json`，依「類別＋類型」對照（例如梁 `45X75`、樓板 `30cm`）。數字來自練習配筋檔 S-0002，**是練習假設，不是正式結構設計**。體積檔需要有「類型」欄；對不到鋼筋比的類型會在日誌出現 `WARN`，該列以 0 計。
- **損耗＋搭接** 10%：業界 #5 以下約 5%、#6 以上依號數增加；模型沒有畫搭接與加密區，所以一併算進來。
- **單價**：

| 資料 | 提供機關 | 更新頻率 |
|---|---|---|
| [大宗資材之鋼筋價格調查](https://data.gov.tw/dataset/6820)（北區 SD420W，元/T） | 行政院公共工程委員會 | 每月 |
| [營造工程物價指數](https://data.gov.tw/dataset/8246)「金屬製品類」 | 行政院主計總處 | 每月 |

鋼筋單價 = 工程會基準價 × 最新月份金屬製品類指數 ÷ 基準價月份指數。同規格有多種直徑時取最細的（目前是 SD420W D32mm）。工程會調查沒有 Revit 配筋用的 #4、#6、#8（D13、D19、D25），以最接近的品項估算；依工程會資料，各號數價差約 2%。工程會與快取都沒有資料時，才改用 `rebar.json` 的 `fallback_price`，網頁會標示「參考市場行情，非政府資料」。鋼筋只算目前模型（舊版本的體積檔沒有類型資料）。

> 工程會網站無法從 GitHub 的海外主機連線，所以自動更新時沿用 `data/price_cache.json` 裡的混凝土與鋼筋基準價，只更新物價指數。要更新基準價時，在自己電腦執行一次 `python scripts/build_data.py` 再上傳即可。

## 常見操作

### 模型改版了
1. Revit 結構明細表（計量類型、樓層、類型、個數、體積）存成 `data/quantities.csv`；或 Forma Takeoff 匯出 Excel 改名為 `data/quantities.xlsx`（兩個都在時先讀 xlsx，換格式時記得刪掉另一個）
2. 有新的構件類型時，在 `data/rebar.json` 補上它的鋼筋比
3. 修改 `data/model_info.json` 的 `note`（例如「第二版」）
4. 上傳後會自動重算並建立新版本（V2、V3…）

### 設定混凝土強度
Takeoff 匯出檔沒有強度資料，預設全部用 280 kgf/cm²。要指定時修改 `data/model_info.json`：

```json
"grades": { "columns": "350", "beams": "280", "floors": "280" }
```

### 寫更新日誌
網站功能或資料有變動時，在 `data/changelog.json` 最上面加一筆（日期＋項目），`type` 可用「新功能」「修正」「資料」「其他」。上傳後網站會自動更新。

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
| `data/quantities.csv` | 結構體積（Revit 明細表：計量類型、樓層、類型、個數、體積），也可改放 Forma Takeoff 匯出的 `quantities.xlsx`（兩個都在時先讀 xlsx） |
| `data/rebar.json` | 鋼筋比、損耗率、鋼筋規格（練習假設） |
| `data/model_info.json` | 模型名稱、版本備註、工地區域、混凝土強度 |
| `data/changelog.json` | 網站更新日誌（手動維護） |
| `data/versions.json` | 各版本體積紀錄（自動維護） |
| `data/price_cache.json` | 最近一次下載的單價資料（自動維護） |
| `data/carbon_cache.json` | 最近一次下載的碳排係數（自動維護） |
| `scripts/build_data.py` | 主程式：讀體積、抓單價、算成本 |
| `scripts/export_offline.py` | 匯出離線版 HTML |
| `scripts/certs/` | 補上政府網站缺少的中繼憑證。網站換憑證時程式會自動下載新的中繼憑證存成 `auto_*.pem`，不用手動處理 |
| `.github/workflows/bim_auto_cost.yml` | 自動排程設定 |
