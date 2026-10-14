"""
BIM 混凝土成本自動估算：把 Forma Takeoff 匯出的體積 × 最新單價，產生網頁用的 data.json。

輸入（都在 data/ 資料夾）：
  quantities.xlsx 或 quantities.csv  ← 從 Forma Takeoff 匯出，模型改版時覆蓋這一份就好
  model_info.json                    ← 模型名稱、備註、工地區域、各構件混凝土強度

單價（自動下載，皆為政府開放資料）：
  基準價：公共工程委員會「大宗資材之混凝土價格調查」各強度、各區域單價
  調整：  主計總處「營造工程物價指數」水泥及其製品類指數，每月更新
  最新單價 = 基準價 × 最新月份指數 ÷ 基準價月份指數

碳排（6D 永續，自動下載）：
  環境部「碳足跡排放係數」開放資料，取各強度「預拌混凝土」最新公告的係數（kgCO2e/m³）
  蘊含碳 = 體積 × 係數

自動維護（不要手動改）：
  versions.json      ← 偵測到 quantities 內容變動時，自動新增一個版本
  price_cache.json   ← 最近一次成功下載的價格資料，網站連不上時拿來備用
  carbon_cache.json  ← 最近一次成功下載的碳排係數，網站連不上時拿來備用

輸出：
  data.json          ← 給 index.html 讀取
"""

import csv
import hashlib
import io
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

sys.stdout.reconfigure(encoding="utf-8")  # Windows 終端機預設 cp950，印不出 m³

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
OUTPUT_FILE = os.path.join(ROOT, "data.json")
CACHE_FILE = os.path.join(DATA_DIR, "price_cache.json")
CARBON_CACHE_FILE = os.path.join(DATA_DIR, "carbon_cache.json")

TW = timezone(timedelta(hours=8))

PCC_PRICE_URL = ("https://pcic.pcc.gov.tw/pwc-web/api/service/opendata-file/document/"
                 + urllib.parse.quote("預拌混凝土.csv"))
DGBAS_INDEX_URL = "https://ws.dgbas.gov.tw/001/Upload/461/relfile/11525/230553/pr0501a1m.xml"
INDEX_ITEM = "水泥及其製品類"
# api_key 是政府資料開放平臺（data.gov.tw/dataset/28176）公開提供的下載金鑰
CFP_URL = ("https://data.moenv.gov.tw/api/v2/cfp_p_02?api_key=4c89a32a-a214-461b-bf29-30ff32a61a8a"
           "&limit=1000&offset={offset}&format=CSV")
# 只取一般預拌混凝土，例如「預拌混凝土(280 kgf/cm2)」；水中、自充填、加飛灰爐石等特殊配比不算
CFP_CONCRETE = re.compile(r"^預拌混凝土\((\d{2,3})\s*kgf/cm2\)$")

DEFAULT_GRADE = "280"
DEFAULT_REGION = "北區"
HISTORY_MONTHS = 12

# Takeoff 匯出的欄位名稱可能是中文或英文，這裡列出可接受的名稱
COLUMN_ALIASES = {
    "category": ["計量類型", "類別", "構件類別", "Takeoff type", "Category", "Classification"],
    "spec": ["材質", "規格", "材料", "Material"],
    "volume": ["體積 (M3)", "體積", "體積(m3)", "體積 (m³)", "Volume (M3)", "Volume"],
    "count": ["個數", "Count"],  # Takeoff 詳細資料一列就是一個構件，沒有這欄時每列算 1 個
    "item": ["行項名稱", "Item name"],  # 只取「體積」這個行項，避免日後加了面積等行項被加進來
}
VOLUME_ITEMS = {"體積", "volume"}

# 把 Revit/Takeoff 的類別名稱歸到四大類
CATEGORY_KEYWORDS = {
    "floors": ["樓板", "板", "floor", "slab"],
    "columns": ["柱", "column"],
    "beams": ["梁", "樑", "framing", "beam"],
    "walls": ["牆", "wall"],
}

logs = []


def log(tag, msg):
    now = datetime.now(TW).strftime("%H:%M:%S")
    logs.append({"time": now, "tag": tag, "msg": msg})
    print(f"[{now}] {tag} {msg}")


def today():
    # BUILD_DATE 只用來測試（模擬不同日期執行），平常不用設定
    return os.environ.get("BUILD_DATE") or datetime.now(TW).strftime("%Y-%m-%d")


def load_model_info():
    path = os.path.join(DATA_DIR, "model_info.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------- 讀取 Takeoff 匯出檔 ----------

def read_table(path):
    if path.endswith(".xlsx"):
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
        # Takeoff 匯出檔第一頁是摘要，逐構件明細在「詳細資料 - …」那一頁
        detail = [ws for ws in wb.worksheets if ws.title.startswith(("詳細資料", "Details"))]
        ws = detail[0] if detail else wb.active
        rows = [[("" if c is None else str(c)).strip() for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    rows = [r for r in rows if any(r)]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, r)) for r in body]


def pick(row, field):
    for name in COLUMN_ALIASES[field]:
        if name in row and row[name] != "":
            return row[name]
    return None


def classify(category):
    text = (category or "").lower()
    for key, words in CATEGORY_KEYWORDS.items():
        if any(w in text for w in words):
            return key
    return None


def parse_grade(text):
    m = re.search(r"(\d{3})\s*kgf", text or "")
    return m.group(1) if m else None


def to_number(value):
    return float(re.sub(r"[^\d.\-]", "", value or "0") or 0)


def load_quantities(grades):
    for name in ("quantities.xlsx", "quantities.csv"):
        path = os.path.join(DATA_DIR, name)
        if os.path.exists(path):
            break
    else:
        raise FileNotFoundError("找不到 data/quantities.xlsx 或 data/quantities.csv")

    with open(path, "rb") as f:
        digest = hashlib.sha256(f.read()).hexdigest()[:12]

    elements = {}
    for row in read_table(path):
        item = pick(row, "item")
        if item and item.strip().lower() not in VOLUME_ITEMS:
            continue
        key = classify(pick(row, "category"))
        if not key:
            continue
        spec = pick(row, "spec") or ""
        # Takeoff 匯出檔沒有混凝土強度，可在 model_info.json 的 "grades" 指定，例如 {"columns": "350"}
        grade = parse_grade(spec) or str(grades.get(key, DEFAULT_GRADE))
        el = elements.setdefault(key, {
            "volume": 0.0, "count": 0, "grade": grade,
            "spec": spec or f"{grade} kgf/cm² 預拌混凝土",
        })
        el["volume"] += to_number(pick(row, "volume"))
        el["count"] += int(to_number(pick(row, "count") or "1"))

    for el in elements.values():
        el["volume"] = round(el["volume"], 2)

    log("TAKEOFF", f"讀取 {name}：{len(elements)} 類構件，總體積 {sum(e['volume'] for e in elements.values()):,.2f} m³")
    return digest, elements


# ---------- 單價（政府開放資料） ----------

def ssl_context():
    # 政府網站常漏送中繼憑證，所以額外載入 scripts/certs/ 裡的 TWCA 中繼憑證；仍是完整驗證
    # 主計總處 2026 年改用「TWCA CYBER Root CA」，舊版 Windows 沒有這張根憑證，有裝 certifi 就一併載入
    ctx = ssl.create_default_context()
    try:
        import certifi
        ctx.load_verify_locations(cafile=certifi.where())
    except ImportError:
        pass
    cert_dir = os.path.join(ROOT, "scripts", "certs")
    for name in sorted(os.listdir(cert_dir)):
        if name.endswith(".pem"):
            ctx.load_verify_locations(cafile=os.path.join(cert_dir, name))
    return ctx


def fetch_missing_intermediate(host):
    """網站換憑證、又漏送中繼憑證時，照憑證裡寫的「CA Issuers」網址把中繼憑證抓回來存到 scripts/certs/。
    抓回來的憑證還是要能接到受信任的根憑證才會通過驗證，所以不會降低安全性。"""
    import socket
    from cryptography import x509
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID

    # 先不驗證地拿到網站憑證，只為了讀出中繼憑證的下載網址
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with socket.create_connection((host, 443), timeout=30) as sock:
        with ctx.wrap_socket(sock, server_hostname=host) as tls:
            leaf = x509.load_der_x509_certificate(tls.getpeercert(binary_form=True))

    aia = leaf.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
    url = next(d.access_location.value for d in aia
               if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS)
    with urllib.request.urlopen(url, timeout=60) as res:
        raw = res.read()
    pem = raw.decode() if raw.startswith(b"-----") else ssl.DER_cert_to_PEM_cert(raw)

    name = re.sub(r"[^\w.-]", "_", os.path.basename(urllib.parse.urlparse(url).path)).rsplit(".", 1)[0]
    with open(os.path.join(ROOT, "scripts", "certs", f"auto_{name}.pem"), "w") as f:
        f.write(pem)
    log("CERT", f"{host} 憑證缺中繼憑證，已自動從 {url} 補上")


def download(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (bim-concrete-cost)"})
    try:
        with urllib.request.urlopen(req, timeout=60, context=ssl_context()) as res:
            return res.read()
    except urllib.error.URLError as e:
        if not isinstance(e.reason, ssl.SSLCertVerificationError):
            raise
        fetch_missing_intermediate(urllib.parse.urlparse(url).hostname)
    with urllib.request.urlopen(req, timeout=60, context=ssl_context()) as res:
        return res.read()


def roc_to_month(roc):
    """民國年月 11504 → 2025-04"""
    roc = roc.strip()
    return f"{int(roc[:-2]) + 1911}-{roc[-2:]}"


def fetch_base_prices(region):
    """工程會預拌混凝土價格調查：回傳 (基準月份, {強度: 單價})，只取最新一期、指定區域。"""
    raw = download(PCC_PRICE_URL)
    for enc in ("utf-8-sig", "big5"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue

    by_period = {}
    for row in csv.DictReader(io.StringIO(text)):
        item = row.get("調查項目", "")
        if "預拌混凝土材料費" not in item or row.get("調查地區", "").strip() != region:
            continue
        grade = parse_grade(item)
        if grade:
            by_period.setdefault(roc_to_month(row["調查時間"]), {})[grade] = to_number(row["價格"])

    if not by_period:
        raise ValueError(f"價格檔裡找不到「{region}」的預拌混凝土資料")
    period = max(by_period)
    return period, by_period[period]


def fetch_cement_index():
    """主計總處營造工程物價指數：回傳 {"2026-08": 123.24, ...}（水泥及其製品類、原始值）。"""
    root = ET.fromstring(download(DGBAS_INDEX_URL))
    series = {}
    for obs in root.iter("Obs"):
        if INDEX_ITEM not in (obs.findtext("Item") or "") or obs.findtext("TYPE") != "原始值":
            continue
        value = (obs.findtext("Item_VALUE") or "").strip()
        if value:
            y, m = obs.findtext("TIME_PERIOD").split("M")
            series[f"{y}-{m}"] = float(value)
    if not series:
        raise ValueError(f"物價指數檔裡找不到「{INDEX_ITEM}」")
    return series


def load_price_data(region):
    cache = {}
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            cache = json.load(f)

    try:
        base_period, base_prices = fetch_base_prices(region)
        log("PRICE", f"工程會基準價（{base_period}・{region}）280 kgf = NT$ {base_prices.get('280', 0):,.0f}")
        cache.update({"region": region, "base_period": base_period, "base_prices": base_prices})
    except Exception as e:
        if cache.get("region") != region:
            raise
        log("WARN", f"工程會價格下載失敗（{e}），改用上次的資料")

    try:
        index = fetch_cement_index()
        latest = max(index)
        log("PRICE", f"主計總處水泥類指數 最新 {latest} = {index[latest]}")
        cache["index"] = index
    except Exception as e:
        if "index" not in cache:
            raise
        log("WARN", f"物價指數下載失敗（{e}），改用上次的資料")

    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    return cache


def prices_for(month, cache):
    """把基準價依指數換算到指定月份。基準月份沒有指數時，直接用基準價。"""
    index = cache["index"]
    base_idx = index.get(cache["base_period"])
    ratio = index[month] / base_idx if base_idx and month in index else 1.0
    return {g: round(p * ratio) for g, p in cache["base_prices"].items()}


def cost_of(elements, prices):
    fallback = prices.get(DEFAULT_GRADE) or next(iter(prices.values()))
    return sum(el["volume"] * prices.get(el["grade"], fallback) for el in elements.values())


# ---------- 碳排係數（環境部開放資料） ----------

def fetch_carbon_factors():
    """回傳 {強度: {"factor": kgCO2e/m³, "year": 公告年, "source": 公告單位}}。
    同一強度有多筆時取最新年份；同一年有多筆時取平均。"""
    rows, offset = [], 0
    while True:
        text = download(CFP_URL.format(offset=offset)).decode("utf-8-sig")
        page = list(csv.DictReader(io.StringIO(text)))
        rows += page
        if len(page) < 1000:
            break
        offset += 1000

    found = {}
    for row in rows:
        m = CFP_CONCRETE.match(row.get("name", "").strip())
        if not m or "m3" not in row.get("unit", ""):
            continue
        year = int(row["announcementyear"])
        best = found.get(m.group(1))
        if best is None or year > best["year"]:
            found[m.group(1)] = {"year": year, "values": [float(row["coe"])], "sources": {row["departmentname"]}}
        elif year == best["year"]:
            best["values"].append(float(row["coe"]))
            best["sources"].add(row["departmentname"])

    if not found:
        raise ValueError("碳足跡係數檔裡找不到預拌混凝土")
    return {g: {"factor": round(sum(f["values"]) / len(f["values"]), 1), "year": f["year"],
                "source": "、".join(sorted(f["sources"]))}
            for g, f in sorted(found.items(), key=lambda kv: int(kv[0]))}


def load_carbon_factors():
    try:
        factors = fetch_carbon_factors()
        with open(CARBON_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(factors, f, ensure_ascii=False, indent=2)
    except Exception as e:
        if not os.path.exists(CARBON_CACHE_FILE):
            raise
        log("WARN", f"碳排係數下載失敗（{e}），改用上次的資料")
        with open(CARBON_CACHE_FILE, encoding="utf-8") as f:
            factors = json.load(f)
    f280 = factors.get(DEFAULT_GRADE)
    if f280:
        log("CARBON", f"環境部碳足跡係數 280 kgf = {f280['factor']} kgCO2e/m³（{f280['year']} 年公告）")
    return factors


def carbon_of(elements, factors):
    """回傳公噸 CO2e。"""
    fallback = factors.get(DEFAULT_GRADE) or next(iter(factors.values()))
    return sum(el["volume"] * factors.get(el["grade"], fallback)["factor"] for el in elements.values()) / 1000


# ---------- 版本 ----------

def update_versions(date, digest, elements, info):
    path = os.path.join(DATA_DIR, "versions.json")
    versions = []
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            versions = json.load(f)

    if versions and versions[-1]["hash"] == digest and versions[-1]["elements"] == elements:
        log("MODEL", f"體積檔案沒有變動，沿用 {versions[-1]['version']}")
        return versions

    if versions and versions[-1]["hash"] == digest:
        # 同一份檔案，只是強度設定改了：更新目前版本，不另開新版本
        versions[-1]["elements"] = elements
        log("MODEL", f"混凝土強度設定變更，更新 {versions[-1]['version']}")
    else:
        versions.append({
            "version": f"V{len(versions) + 1}",
            "model_name": info.get("model_name", "quantities"),
            "note": info.get("note", ""),
            "author": info.get("author", ""),
            "detected_at": date,
            "hash": digest,
            "elements": elements,
        })
        log("MODEL", f"偵測到新的體積檔案，建立 {versions[-1]['version']}（{versions[-1]['model_name']}）")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(versions, f, ensure_ascii=False, indent=2)
    return versions


# ---------- 主程式 ----------

def main():
    date = today()
    info = load_model_info()
    region = info.get("region", DEFAULT_REGION)
    log("INFO", f"開始估算（{date}）")

    digest, elements = load_quantities(info.get("grades", {}))
    versions = update_versions(date, digest, elements, info)

    cache = load_price_data(region)
    months = sorted(cache["index"])
    latest, previous = months[-1], months[-2]
    prices = prices_for(latest, cache)
    carbon_factors = load_carbon_factors()
    log("PRICE", f"換算後 {latest} 單價：" + "、".join(f"{g} kgf NT$ {prices[g]:,}" for g in ("210", "280", "350") if g in prices))

    # 每個版本都用「目前單價」計算，版本之間才能公平比較
    for v in versions:
        v["total_volume"] = round(sum(e["volume"] for e in v["elements"].values()), 2)
        v["element_count"] = sum(e["count"] for e in v["elements"].values())
        v["cost_at_current_price"] = round(cost_of(v["elements"], prices))
        v["carbon_tco2e"] = round(carbon_of(v["elements"], carbon_factors), 1)

    # 成本走勢：目前模型 × 過去 12 個月的單價
    latest_version = versions[-1]
    cost_history = [
        {"month": m, "cost": round(cost_of(latest_version["elements"], prices_for(m, cache)))}
        for m in months[-HISTORY_MONTHS:]
    ]

    log("SUCCESS", f"{latest_version['version']} 總體積 {latest_version['total_volume']:,.2f} m³，"
                   f"總成本 NT$ {latest_version['cost_at_current_price']:,}，"
                   f"蘊含碳 {latest_version['carbon_tco2e']:,.1f} tCO2e")

    output = {
        "generated_at": datetime.now(TW).strftime("%Y-%m-%d %H:%M"),
        "is_sample": bool(info.get("is_sample")),
        "region": region,
        "price_month": latest,
        "base_period": cache["base_period"],
        "base_index": cache["index"].get(cache["base_period"]),
        "latest_index": cache["index"][latest],
        "prices": prices,
        "previous_prices": prices_for(previous, cache),
        "versions": versions,
        "cost_history": cost_history,
        "carbon_factors": carbon_factors,
        "logs": logs,
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
