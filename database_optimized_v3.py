from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "products.db"
DATA_DIR.mkdir(exist_ok=True)

_ENGINE: Engine | None = None
_DB_INIT_DONE = False

# Normalized products table. 兼容原卫浴报价表，也兼容新版 PPR 宽表。
PRODUCT_COLUMNS = [
    "sap", "product_type", "subcategory", "category",
    "cn_name", "en_name", "model", "description",
    "color", "size_mm", "weight", "material_description",
    "price", "price_cny", "stock", "packing_volume", "qty_per_ctn",
    "package_length", "package_width", "package_height",
    "unit", "currency", "package_info", "image_url", "active",
]

TEXT_COLUMNS = [
    "sap", "product_type", "subcategory", "category", "cn_name", "en_name", "model",
    "description", "color", "size_mm", "material_description", "unit", "currency",
    "package_info", "image_url",
]
NUMERIC_COLUMNS = [
    "price", "price_cny", "stock", "packing_volume", "qty_per_ctn",
    "package_length", "package_width", "package_height", "weight", "active",
]

EXCEL_ERROR_VALUES = {"#N/A", "#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "nan", "none", "null", ""}

COLUMN_ALIASES = {
    "sap号": "sap", "sap": "sap", "sap code": "sap", "SAP": "sap", "SAP号": "sap", "物料": "sap", "物料编码": "sap",
    "分类": "category", "类别": "category", "category": "category", "仓位描述": "category", "定价小组分类": "category",
    "品名": "cn_name", "中文名": "cn_name", "中文品名": "cn_name", "产品名称": "cn_name", "name": "cn_name",
    "英文名": "en_name", "英文品名": "en_name", "英文产品名": "en_name", "en_name": "en_name", "english name": "en_name",
    "型号": "model", "model": "model", "Size(mm)": "size_mm", "size(mm)": "size_mm", "size": "size_mm", "规格": "size_mm",
    "描述": "description", "产品描述": "description", "description": "description",
    "颜色": "color", "color": "color", "colour": "color",
    "重量": "weight", "weight": "weight", "Weight\n(kg/m or pc)": "weight", "weight(kg/m or pc)": "weight", "kg/m": "weight",
    "物料描述": "material_description", "material_description": "material_description", "物料描述（灰色）": "material_description", "物料描述（绿色）": "material_description",
    "价格": "price", "报价": "price", "fob": "price", "FOB价": "price", "FOB价（USD/PC）": "price", "fob usd": "price", "单价": "price",
    "灰色基准价格（USD）/m(pcs)": "price", "绿色基准价格（USD）/m(pcs)": "price",
    "人民币价格": "price_cny", "price_cny": "price_cny", "灰色基准价格（CNY）/m(pcs)": "price_cny", "绿色基准价格（CNY）/m(pcs)": "price_cny",
    "库存": "stock", "stock": "stock",
    "包装体积": "packing_volume", "体积": "packing_volume", "cbm": "packing_volume", "CBM": "packing_volume", "CBM/PC": "packing_volume", "cbm/pc": "packing_volume",
    "Qty/CTN": "qty_per_ctn", "qty/ctn": "qty_per_ctn", "QTY/CTN": "qty_per_ctn", "qty_per_ctn": "qty_per_ctn",
    "Pcs/Carton": "qty_per_ctn", "pcs/carton": "qty_per_ctn", "一箱所含件": "qty_per_ctn", "每箱数量": "qty_per_ctn", "装箱数量": "qty_per_ctn", "每箱件数": "qty_per_ctn", "pcs/ctn": "qty_per_ctn", "PCS/CTN": "qty_per_ctn",
    "L": "package_length", "length": "package_length", "长": "package_length",
    "W": "package_width", "width": "package_width", "宽": "package_width",
    "H": "package_height", "height": "package_height", "高": "package_height",
    "单位": "unit", "unit": "unit", "UOM": "unit", "uom": "unit",
    "币种": "currency", "currency": "currency",
    "包装": "package_info", "包装信息": "package_info", "package": "package_info", "package_info": "package_info",
    "图片": "image_url", "图片链接": "image_url", "图片URL": "image_url", "image": "image_url", "image_url": "image_url", "image link": "image_url", "Picture": "image_url", "Picture ": "image_url", "灰色图片": "image_url",
    "状态": "active", "active": "active", "是否启用": "active",
}

# Extra aliases for the simple product-master template the user uses for quick uploads.
# Example headers: SAP No. / product describe / TYPE / 白色物料描述（5.8m） / 规格（mm) / 壁厚（mm) / PRICE
COLUMN_ALIASES.update({
    "SAP No.": "sap", "SAP No": "sap", "sap no.": "sap", "sap no": "sap",
    "产品大类": "product_type", "大类": "product_type", "TYPE": "product_type", "type": "product_type", "Product Type": "product_type", "product type": "product_type",
    "产品小类": "subcategory", "小类": "subcategory", "子分类": "subcategory", "subcategory": "subcategory",
    "product describe": "subcategory", "Product Describe": "subcategory", "product description type": "subcategory",
    "白色物料描述（5.8m）": "material_description", "白色物料描述(5.8m)": "material_description", "白色物料描述": "material_description",
    "规格（mm)": "size_mm", "规格（mm）": "size_mm", "规格(mm)": "size_mm",
    "PRICE": "price", "price": "price",
})


def _read_secret(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value:
        return value
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return default


def get_database_url() -> str:
    url = _read_secret("DATABASE_URL")
    if not url:
        return f"sqlite:///{DB_PATH}"
    url = str(url).strip()
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    if url.startswith("postgresql://"):
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
    return url


DATABASE_URL = get_database_url()


def get_database_backend_name() -> str:
    if DATABASE_URL.startswith("postgresql"):
        return "PostgreSQL / Supabase"
    if DATABASE_URL.startswith("sqlite"):
        return "SQLite / Local"
    return "Unknown"


def get_engine() -> Engine:
    global _ENGINE
    if _ENGINE is None:
        connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
        _ENGINE = create_engine(DATABASE_URL, future=True, pool_pre_ping=True, connect_args=connect_args)
    return _ENGINE


def _dialect() -> str:
    return get_engine().dialect.name


def _table_exists(conn, table_name: str) -> bool:
    if _dialect() == "sqlite":
        row = conn.execute(text("SELECT name FROM sqlite_master WHERE type='table' AND name=:t"), {"t": table_name}).fetchone()
        return row is not None
    row = conn.execute(
        text("SELECT table_name FROM information_schema.tables WHERE table_schema=current_schema() AND table_name=:t"),
        {"t": table_name},
    ).fetchone()
    return row is not None


def _product_columns(conn) -> set[str]:
    if not _table_exists(conn, "products"):
        return set()
    if _dialect() == "sqlite":
        rows = conn.execute(text("PRAGMA table_info(products)")).fetchall()
        return {row[1] for row in rows}
    rows = conn.execute(text("""
        SELECT column_name FROM information_schema.columns
        WHERE table_name='products' AND table_schema=current_schema()
    """)).fetchall()
    return {row[0] for row in rows}


def _backup_invalid_products_if_needed(conn) -> None:
    cols = _product_columns(conn)
    if cols and "sap" not in cols:
        suffix = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"products_invalid_{suffix}"
        conn.execute(text(f"ALTER TABLE products RENAME TO {backup_name}"))


def init_db() -> None:
    """Create/migrate tables.

    The Streamlit process may keep _DB_INIT_DONE=True while the Supabase products table
    is manually deleted. Therefore, even after initialization we quickly re-check the
    required tables and rebuild if missing.
    """
    global _DB_INIT_DONE
    engine = get_engine()
    if _DB_INIT_DONE:
        try:
            with engine.connect() as conn:
                if _table_exists(conn, "products") and _table_exists(conn, "users") and _table_exists(conn, "quote_history"):
                    return
        except Exception:
            pass
        _DB_INIT_DONE = False

    with engine.begin() as conn:
        _backup_invalid_products_if_needed(conn)
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS products (
                sap VARCHAR(80) PRIMARY KEY,
                product_type TEXT DEFAULT '',
                subcategory TEXT DEFAULT '',
                category TEXT,
                cn_name TEXT,
                en_name TEXT DEFAULT '',
                model TEXT,
                description TEXT,
                color TEXT DEFAULT '',
                size_mm TEXT DEFAULT '',
                weight REAL DEFAULT 0,
                material_description TEXT DEFAULT '',
                price REAL DEFAULT 0,
                price_cny REAL DEFAULT 0,
                stock INTEGER DEFAULT 0,
                packing_volume REAL DEFAULT 0,
                qty_per_ctn REAL DEFAULT 0,
                package_length REAL DEFAULT 0,
                package_width REAL DEFAULT 0,
                package_height REAL DEFAULT 0,
                unit TEXT DEFAULT 'PC',
                currency TEXT DEFAULT 'USD',
                package_info TEXT,
                image_url TEXT,
                active INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))
        if _dialect() == "postgresql":
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS quote_history (
                    id SERIAL PRIMARY KEY,
                    quote_no TEXT,
                    customer TEXT,
                    total_amount REAL,
                    created_by TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """))
        else:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS quote_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    quote_no TEXT,
                    customer TEXT,
                    total_amount REAL,
                    created_by TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS users (
                username VARCHAR(80) PRIMARY KEY,
                password_hash TEXT NOT NULL,
                role VARCHAR(20) NOT NULL DEFAULT 'user',
                active INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))

    migrate_products_table()
    with engine.begin() as conn:
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_active_category_sap ON products(active, category, sap)"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_category_sap ON products(category, sap)"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_type_subcategory_sap ON products(product_type, subcategory, sap)"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_sap_lookup ON products(sap)"))
    _DB_INIT_DONE = True


def migrate_products_table() -> None:
    required = {
        "product_type": "TEXT DEFAULT ''",
        "subcategory": "TEXT DEFAULT ''",
        "category": "TEXT",
        "cn_name": "TEXT",
        "en_name": "TEXT DEFAULT ''",
        "model": "TEXT",
        "description": "TEXT",
        "color": "TEXT DEFAULT ''",
        "size_mm": "TEXT DEFAULT ''",
        "weight": "REAL DEFAULT 0",
        "material_description": "TEXT DEFAULT ''",
        "price": "REAL DEFAULT 0",
        "price_cny": "REAL DEFAULT 0",
        "stock": "INTEGER DEFAULT 0",
        "packing_volume": "REAL DEFAULT 0",
        "qty_per_ctn": "REAL DEFAULT 0",
        "package_length": "REAL DEFAULT 0",
        "package_width": "REAL DEFAULT 0",
        "package_height": "REAL DEFAULT 0",
        "unit": "TEXT DEFAULT 'PC'",
        "currency": "TEXT DEFAULT 'USD'",
        "package_info": "TEXT",
        "image_url": "TEXT DEFAULT ''",
        "active": "INTEGER DEFAULT 1",
        "created_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "updated_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
    }
    with get_engine().begin() as conn:
        existing = _product_columns(conn)
        if existing and "sap" not in existing:
            _backup_invalid_products_if_needed(conn)
            existing = _product_columns(conn)
        for col, definition in required.items():
            if col not in existing:
                conn.execute(text(f"ALTER TABLE products ADD COLUMN {col} {definition}"))


def _storage_public_url(storage_path: str) -> str:
    supabase_url = (_read_secret("SUPABASE_URL", "https://oevlzhvdgojgzacnbfka.supabase.co") or "").rstrip("/")
    bucket = (_read_secret("SUPABASE_STORAGE_BUCKET", "product-images") or "product-images").strip("/")
    encoded_path = quote(str(storage_path).strip().lstrip("/"), safe="/")
    if not supabase_url or not bucket or not encoded_path:
        return ""
    return f"{supabase_url}/storage/v1/object/public/{bucket}/{encoded_path}"


def normalize_image_url(value: object, sap: object | None = None) -> str:
    raw = clean_scalar(value)
    if not raw:
        return ""
    if raw.startswith("http://") or raw.startswith("https://"):
        return raw

    bucket = (_read_secret("SUPABASE_STORAGE_BUCKET", "product-images") or "product-images").strip("/")
    # Default to storage root because current product-images bucket uses paths like 8060040825.jpg/png.
    prefix = (_read_secret("SUPABASE_STORAGE_PREFIX", "") or "").strip("/")

    if raw.startswith(bucket + "/"):
        raw = raw[len(bucket) + 1:]

    name = Path(raw).name
    suffix = Path(name).suffix.lower()
    if "/" not in raw:
        if not suffix:
            stem = clean_scalar(sap or raw)
            raw = f"{stem}.png"
        raw = f"{prefix}/{raw}" if prefix else raw
    elif not suffix and sap:
        raw = f"{raw}.png"
    return _storage_public_url(raw)


def clean_scalar(value: object) -> str:
    text_value = str(value if value is not None else "").strip()
    if text_value.lower() in EXCEL_ERROR_VALUES or text_value.upper() in EXCEL_ERROR_VALUES:
        return ""
    if text_value.endswith(".0") and text_value[:-2].isdigit():
        return text_value[:-2]
    return text_value


def numeric_value(value: object, default: float = 0.0) -> float:
    text_value = clean_scalar(value)
    if not text_value:
        return default
    try:
        return float(str(text_value).replace(",", ""))
    except Exception:
        return default


def _get_col(row: pd.Series, *names: str) -> object:
    for name in names:
        if name in row.index:
            return row.get(name)
    stripped_map = {str(c).strip(): c for c in row.index}
    for name in names:
        key = str(name).strip()
        if key in stripped_map:
            return row.get(stripped_map[key])
    return None



def _get_at(row: pd.Series, index: int) -> object:
    """Get a cell by zero-based position. Used for merged Excel headers such as box L/W/H."""
    try:
        if len(row.index) > index:
            return row.iloc[index]
    except Exception:
        pass
    return None


def _extract_size_from_text(text: object) -> str:
    value = clean_scalar(text)
    if not value:
        return ""
    patterns = [
        r"(dn\s*\d+(?:\s*[×xX*]\s*\d+)*)",
        r"(\d+(?:\s*[×xX*]\s*\d+)+)",
        r"DN\s*(\d+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, value, flags=re.IGNORECASE)
        if m:
            return re.sub(r"\s+", "", m.group(1)).replace("x", "×").replace("X", "×").replace("*", "×")
    return ""



def _normalize_product_type(value: object) -> str:
    """Normalize broad product groups for easier filtering: PPR / PVC / PE.

    Important: when inferring from long material descriptions, do not keep the
    whole description as a product type. Otherwise the category dropdown becomes
    polluted with thousands of one-off values from price-base sheets.
    """
    raw = clean_scalar(value).upper().replace("PP-R", "PPR").replace("PVC-U", "PVC").replace("UPVC", "PVC")
    if not raw:
        return ""
    if "PPR" in raw:
        return "PPR"
    if "PVC" in raw:
        return "PVC"
    if re.search(r"\bHDPE\b|\bPE\b|PE-RT|PERT|P\.E\.", raw):
        return "PE"
    return ""


def _infer_product_type(*values: object) -> str:
    text = " ".join(clean_scalar(v) for v in values if clean_scalar(v))
    return _normalize_product_type(text)


def _clean_subcategory(value: object) -> str:
    text_value = clean_scalar(value)
    if not text_value:
        return ""
    # Keep useful pressure/series info, only trim excessive whitespace.
    return re.sub(r"\s+", " ", text_value).strip()


def _format_category(product_type: object, subcategory: object, fallback: object = "") -> str:
    ptype = _normalize_product_type(product_type)
    sub = _clean_subcategory(subcategory)
    if ptype and sub:
        return f"{ptype} / {sub}"
    if sub:
        return sub
    if ptype:
        return ptype
    return clean_scalar(fallback)


def _guess_unit_from_text(*values: object) -> str:
    text = " ".join(clean_scalar(v).lower() for v in values if clean_scalar(v))
    if any(k in text for k in ["pipe", "直管", "管材", "5.8m", "6m", "4m", "米"]):
        return "M"
    return "PCS"


def _looks_like_simple_product_catalog(df: pd.DataFrame) -> bool:
    """Recognize a quick product-master table such as 工作簿1.xlsx.

    Required idea: one row = one SAP. Typical headers:
    SAP No. / product describe / TYPE / 白色物料描述（5.8m） / 规格（mm) / 壁厚（mm) / PRICE
    """
    cols = {str(c).strip() for c in df.columns}
    lower_cols = {c.lower() for c in cols}
    has_sap = bool({"SAP No.", "SAP No", "sap no.", "sap no", "SAP", "sap", "物料", "物料编码"} & cols) or bool({"sap no.", "sap no", "sap"} & lower_cols)
    has_type_or_desc = bool({"TYPE", "type", "Product Type", "product type", "产品大类"} & cols) or "product describe" in lower_cols
    has_price = bool({"PRICE", "price", "单价", "价格", "报价"} & cols) or "price" in lower_cols
    return has_sap and has_type_or_desc and has_price


def expand_simple_product_catalog_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert the user's simple product-master file into normalized product rows."""
    rows: list[dict[str, object]] = []
    seen: set[str] = set()

    for _, r in df.iterrows():
        sap = clean_scalar(_get_col(r, "SAP No.", "SAP No", "sap no.", "sap no", "SAP", "sap", "物料", "物料编码"))
        if not sap or sap in seen:
            continue
        seen.add(sap)

        product_type_raw = _get_col(r, "TYPE", "type", "Product Type", "product type", "产品大类", "大类")
        subcategory = _clean_subcategory(_get_col(r, "product describe", "Product Describe", "产品小类", "小类", "子分类", "subcategory", "category"))
        material_desc = clean_scalar(_get_col(r, "白色物料描述（5.8m）", "白色物料描述(5.8m)", "白色物料描述", "物料描述", "material_description", "description"))
        size = clean_scalar(_get_col(r, "规格（mm)", "规格（mm）", "规格(mm)", "规格", "Size(mm)", "size_mm", "size")) or _extract_size_from_text(material_desc or subcategory)
        wall = clean_scalar(_get_col(r, "壁厚（mm)", "壁厚（mm）", "壁厚(mm)", "壁厚", "wall thickness"))
        price = numeric_value(_get_col(r, "PRICE", "price", "单价", "价格", "报价"))
        product_type = _normalize_product_type(product_type_raw) or _infer_product_type(subcategory, material_desc)
        category = _format_category(product_type, subcategory, fallback=material_desc)
        color = "White" if ("白色" in material_desc or "white" in material_desc.lower()) else ""
        unit = _guess_unit_from_text(subcategory, material_desc)

        desc_parts = []
        if subcategory:
            desc_parts.append(subcategory)
        if material_desc:
            desc_parts.append(material_desc)
        if size:
            desc_parts.append(f"Size: {size}")
        if wall:
            desc_parts.append(f"Wall thickness: {wall} mm")

        model_parts = [p for p in [size, f"Wall {wall}mm" if wall else ""] if p]
        model = " / ".join(model_parts) if model_parts else size

        rows.append({
            "sap": sap,
            "product_type": product_type,
            "subcategory": subcategory,
            "category": category,
            "cn_name": material_desc or subcategory,
            "en_name": "",
            "model": model,
            "description": " | ".join(desc_parts),
            "color": color,
            "size_mm": size,
            "weight": 0,
            "material_description": material_desc,
            "price": price,
            "price_cny": 0,
            "stock": 0,
            "packing_volume": 0,
            "qty_per_ctn": 0,
            "package_length": 0,
            "package_width": 0,
            "package_height": 0,
            "unit": unit,
            "currency": "USD",
            "package_info": "",
            "image_url": "",
            "active": 1,
        })
    return pd.DataFrame(rows)


def _looks_like_pvcu_drainage_wide(df: pd.DataFrame) -> bool:
    cols = {str(c).strip() for c in df.columns}
    return (
        "英文描述" in cols
        and "产品名称" in cols
        and (
            "白色物料编码4m/条" in cols
            or "白色物料编码 6m/条" in cols
            or "白色物料编码 5.8m/条" in cols
        )
    )


def _looks_like_usd_price_base(df: pd.DataFrame) -> bool:
    cols = {str(c).strip() for c in df.columns}
    return {"物料", "物料描述", "单价"}.issubset(cols)


def expand_pvcu_drainage_wide_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert PVC-U drainage quotation wide format into one row per SAP SKU.

    The source may place three pipe SAP codes in one row: 4m / 6m / 5.8m.
    Fittings usually use the first SAP column only. This function normalizes them
    into the products table format used by the quotation system.
    """
    rows: list[dict[str, object]] = []
    seen: set[str] = set()

    for _, r in df.iterrows():
        en_desc = clean_scalar(_get_col(r, "英文描述", "Description", "description"))
        cn_name = clean_scalar(_get_col(r, "产品名称", "品名", "中文品名"))
        size = clean_scalar(_get_col(r, "规格（mm)", "规格(mm)", "规格", "size_mm", "Size(mm)", "Size"))
        wall_4m = clean_scalar(_get_col(r, "壁厚（mm）", "壁厚(mm)", "壁厚"))
        wall_6m = clean_scalar(_get_col(r, "壁厚（mm）.1"))
        price = numeric_value(_get_col(r, "单价（USD/M）\n单价(USD/只)", "单价（USD/M）", "单价(USD/只)", "price", "价格"))
        weight = numeric_value(_get_col(r, "米重（kg/m)\n单重\n（kg/只）", "米重（kg/m)", "单重（kg/只）", "weight"))
        price_unit_multiple = numeric_value(_get_col(r, "每"))
        qty_per_ctn = numeric_value(_get_col(r, "包装\n（只/箱）", "包装（只/箱）", "Qty/CTN", "Pcs/Carton"))
        length = numeric_value(_get_col(r, "箱子尺寸 (m)", "L", "length")) or numeric_value(_get_at(r, 17))
        width = numeric_value(_get_col(r, "Unnamed: 18", "W", "width")) or numeric_value(_get_at(r, 18))
        height = numeric_value(_get_col(r, "Unnamed: 19", "H", "height")) or numeric_value(_get_at(r, 19))
        cbm = numeric_value(_get_col(r, "体积CBM", "CBM", "cbm", "packing_volume")) or numeric_value(_get_at(r, 20))
        if cbm == 0 and length and width and height:
            cbm = round(length * width * height, 6)

        if not en_desc and not cn_name:
            continue

        is_pipe = "pipe" in en_desc.lower() or "管材" in cn_name or "排水管" in cn_name
        unit = "M" if is_pipe else "PCS"
        category = en_desc or cn_name or "PVC-U Drainage"
        image_url = clean_scalar(_get_col(r, "图片", "图片.1", "Picture", "image_url"))

        variants = [
            {
                "sap": _get_col(r, "白色物料编码4m/条"),
                "length_label": "4m" if is_pipe else "",
                "wall": wall_4m,
                "material_description": _get_col(r, "物料描述4m/条"),
            },
            {
                "sap": _get_col(r, "白色物料编码 6m/条"),
                "length_label": "6m" if is_pipe else "",
                "wall": wall_6m or wall_4m,
                "material_description": _get_col(r, "物料描述6m/条"),
            },
            {
                "sap": _get_col(r, "白色物料编码 5.8m/条"),
                "length_label": "5.8m" if is_pipe else "",
                "wall": wall_6m or wall_4m,
                "material_description": _get_col(r, "物料描述5.8m/条"),
            },
        ]

        for v in variants:
            sap = clean_scalar(v.get("sap"))
            material_desc = clean_scalar(v.get("material_description"))
            if not sap or sap in seen:
                continue
            if not is_pipe and v.get("length_label"):
                continue
            # Fittings usually only have the first SAP column. If later columns are empty, they are skipped.
            seen.add(sap)

            length_label = clean_scalar(v.get("length_label"))
            wall = clean_scalar(v.get("wall"))
            model_parts = [p for p in [size, length_label] if p]
            model = " / ".join(model_parts) if model_parts else size

            desc_parts = []
            if en_desc:
                desc_parts.append(en_desc)
            if cn_name:
                desc_parts.append(cn_name)
            if size:
                desc_parts.append(f"Size: {size}")
            if wall:
                desc_parts.append(f"Wall thickness: {wall} mm")
            if length_label:
                desc_parts.append(f"Length: {length_label}")
            if material_desc:
                desc_parts.append(material_desc)

            package_parts = []
            if qty_per_ctn:
                package_parts.append(f"{qty_per_ctn:g} pcs/carton")
            if length or width or height:
                package_parts.append(f"L{length:g}×W{width:g}×H{height:g} m")
            if price_unit_multiple:
                package_parts.append(f"price base: per {price_unit_multiple:g}")

            rows.append({
                "sap": sap,
                "product_type": "PVC",
                "subcategory": category,
                "category": _format_category("PVC", category),
                "cn_name": cn_name or material_desc,
                "en_name": en_desc,
                "model": model,
                "description": " | ".join(desc_parts),
                "color": "White",
                "size_mm": size,
                "weight": weight,
                "material_description": material_desc,
                "price": price,
                "price_cny": 0,
                "stock": 0,
                "packing_volume": cbm,
                "qty_per_ctn": qty_per_ctn,
                "package_length": length,
                "package_width": width,
                "package_height": height,
                "unit": unit,
                "currency": "USD",
                "package_info": "; ".join(package_parts),
                "image_url": image_url,
                "active": 1,
            })
    return pd.DataFrame(rows)


def expand_usd_price_base_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert USD base price table with columns 物料 / 物料描述 / 单价 into products rows."""
    rows: list[dict[str, object]] = []
    seen: set[str] = set()

    for _, r in df.iterrows():
        sap = clean_scalar(_get_col(r, "物料", "物料编码", "sap"))
        if not sap or sap in seen:
            continue
        seen.add(sap)

        material_desc = clean_scalar(_get_col(r, "物料描述", "material_description"))
        price = numeric_value(_get_col(r, "单价", "price"))
        amount = numeric_value(_get_col(r, "金额"))
        per = numeric_value(_get_col(r, "每"))
        if price == 0 and amount and per:
            price = amount / per

        currency = clean_scalar(_get_col(r, "币别", "currency")) or "USD"
        unit = clean_scalar(_get_col(r, "计量单位", "unit")) or "PC"
        category = clean_scalar(_get_col(r, "仓位描述", "定价小组分类", "category")) or "Price Base"
        size = _extract_size_from_text(material_desc)

        rows.append({
            "sap": sap,
            "product_type": _infer_product_type(category, material_desc),
            "subcategory": category,
            "category": _format_category(_infer_product_type(category, material_desc), category),
            "cn_name": material_desc,
            "en_name": "",
            "model": size,
            "description": material_desc,
            "color": "White" if "白色" in material_desc else "",
            "size_mm": size,
            "weight": 0,
            "material_description": material_desc,
            "price": price,
            "price_cny": 0,
            "stock": 0,
            "packing_volume": 0,
            "qty_per_ctn": 0,
            "package_length": 0,
            "package_width": 0,
            "package_height": 0,
            "unit": unit,
            "currency": currency,
            "package_info": "",
            "image_url": "",
            "active": 1,
        })
    return pd.DataFrame(rows)


def _first_non_empty(values: pd.Series) -> object:
    for v in values:
        text_value = clean_scalar(v)
        if text_value:
            return v
    return ""


def _last_non_empty(values: pd.Series) -> object:
    for v in reversed(list(values)):
        text_value = clean_scalar(v)
        if text_value:
            return v
    return ""


def _last_non_zero(values: pd.Series) -> float:
    for v in reversed(list(values)):
        n = numeric_value(v)
        if n != 0:
            return n
    return 0


def combine_imported_product_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Merge normalized products from multiple sheets.

    When the same SAP appears in a detailed product sheet and a base-price sheet,
    keep rich text/spec/packing fields from the first detailed occurrence, but use
    the latest non-zero price/currency/unit if later sheets provide them.
    """
    frames = [f for f in frames if f is not None and not f.empty]
    if not frames:
        return pd.DataFrame(columns=PRODUCT_COLUMNS)

    combined = pd.concat(frames, ignore_index=True)
    for col in PRODUCT_COLUMNS:
        if col not in combined.columns:
            combined[col] = None
    combined = combined[PRODUCT_COLUMNS].copy()
    combined["sap"] = combined["sap"].apply(clean_scalar)
    combined = combined[combined["sap"] != ""]
    if combined.empty:
        return combined

    rows: list[dict[str, object]] = []
    for sap, grp in combined.groupby("sap", sort=False):
        out: dict[str, object] = {"sap": sap}
        for col in PRODUCT_COLUMNS:
            if col == "sap":
                continue
            if col in {"price", "price_cny", "stock", "packing_volume", "qty_per_ctn", "package_length", "package_width", "package_height", "weight"}:
                # Later sheets can update price; detailed sheets usually provide packing/weight.
                out[col] = _last_non_zero(grp[col]) if col in {"price", "price_cny", "stock"} else (_first_non_empty(grp[col]) or 0)
            elif col in {"unit", "currency", "active"}:
                out[col] = _last_non_empty(grp[col]) or ("USD" if col == "currency" else "PC" if col == "unit" else 1)
            else:
                out[col] = _first_non_empty(grp[col])
        rows.append(out)

    return pd.DataFrame(rows, columns=PRODUCT_COLUMNS)


def _merge_price_base_into_details(detailed: pd.DataFrame, price_base: pd.DataFrame) -> pd.DataFrame:
    """Use a USD base-price sheet only as a supplement for already-detected products.

    Many company price-base sheets contain tens of thousands of unrelated SAP codes.
    When a workbook also has a detailed product sheet, importing every base row will
    pollute the website database and dropdown filters. Therefore, only matching SAPs
    are used to refresh price/currency and fill empty descriptions.
    """
    if detailed is None or detailed.empty or price_base is None or price_base.empty:
        return detailed

    result = detailed.copy()
    base = price_base.copy()
    base["sap"] = base["sap"].apply(clean_scalar)
    base = base[base["sap"] != ""].drop_duplicates(subset=["sap"], keep="last").set_index("sap")

    for idx, row in result.iterrows():
        sap = clean_scalar(row.get("sap"))
        if not sap or sap not in base.index:
            continue
        b = base.loc[sap]
        # Price-base sheets are trusted for price, but should not overwrite the
        # richer product structure, unit, packing, model or category from detail sheets.
        for col in ["price", "price_cny"]:
            val = numeric_value(b.get(col))
            if val:
                result.at[idx, col] = val
        currency = clean_scalar(b.get("currency"))
        if currency:
            result.at[idx, "currency"] = currency
        for col in ["material_description", "cn_name", "description"]:
            if not clean_scalar(result.at[idx, col]):
                fill = clean_scalar(b.get(col))
                if fill:
                    result.at[idx, col] = fill

    return result


def read_product_import_file(file) -> pd.DataFrame:
    """Read all supported product import formats from Excel/CSV.

    Supported formats:
    1. Standard products table with SAP column aliases.
    2. Quick product-master table: SAP No. / product describe / TYPE / PRICE.
    3. PPR wide quotation table with Grey SAP No. / Green SAP No.
    4. PVC-U drainage wide quotation table with 4m / 6m / 5.8m SAP columns.
    5. USD base price table with 物料 / 物料描述 / 单价.

    Excel files are scanned sheet by sheet. If a workbook contains a detailed
    product sheet plus a large USD base sheet, the base sheet is used only to
    supplement matching SAPs instead of importing all unrelated rows.
    """
    filename = str(getattr(file, "name", "") or "").lower()
    if hasattr(file, "seek"):
        file.seek(0)

    if filename.endswith(".csv"):
        raw = pd.read_csv(file).dropna(how="all").dropna(axis=1, how="all")
        return clean_products_df(raw)

    if not filename.endswith((".xlsx", ".xls")):
        raise ValueError("仅支持 xlsx / xls / csv 产品表。")

    excel = pd.ExcelFile(file)
    detail_frames: list[pd.DataFrame] = []
    base_price_frames: list[pd.DataFrame] = []
    errors: list[str] = []

    for sheet_name in excel.sheet_names:
        try:
            raw = pd.read_excel(excel, sheet_name=sheet_name)
            raw = raw.dropna(how="all").dropna(axis=1, how="all")
            if raw.empty:
                continue

            # Base-price sheets are handled separately to avoid importing an
            # entire all-company price book when the workbook also contains a
            # small detailed product sheet.
            if _looks_like_usd_price_base(raw) and not (
                _looks_like_ppr_wide(raw) or _looks_like_pvcu_drainage_wide(raw) or _looks_like_simple_product_catalog(raw)
            ):
                normalized_base = clean_products_df(raw)
                if not normalized_base.empty:
                    base_price_frames.append(normalized_base)
                continue

            normalized = clean_products_df(raw)
            if not normalized.empty:
                detail_frames.append(normalized)
        except Exception as exc:
            errors.append(f"{sheet_name}: {exc}")

    if detail_frames:
        detailed = combine_imported_product_frames(detail_frames)
        if base_price_frames:
            base_prices = combine_imported_product_frames(base_price_frames)
            detailed = _merge_price_base_into_details(detailed, base_prices)
        return detailed

    if base_price_frames:
        return combine_imported_product_frames(base_price_frames)

    detail = "；".join(errors[:5])
    raise ValueError(
        "未识别到可导入的产品数据。支持：标准字段、快速产品大全表、PPR宽表、PVC-U排水宽表、USD基准表。"
        + (f" 工作表检查结果：{detail}" if detail else "")
    )


def _looks_like_ppr_wide(df: pd.DataFrame) -> bool:
    cols = {str(c).strip() for c in df.columns}
    return "Grey SAP No." in cols or "Green SAP No." in cols


def expand_ppr_wide_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert PPR quotation wide format into one row per SAP SKU.

    Source columns include Grey SAP No. and Green SAP No.; each source row may become
    two normalized products. If grey and green SAP are the same, only one row is kept.
    """
    rows: list[dict[str, object]] = []
    seen: set[str] = set()

    for _, r in df.iterrows():
        en_desc = clean_scalar(_get_col(r, "Description"))
        cn_name = clean_scalar(_get_col(r, "产品名称"))
        size = clean_scalar(_get_col(r, "Size(mm)", "Size", "size_mm"))
        weight = numeric_value(_get_col(r, "Weight\n(kg/m or pc)", "Weight(kg/m or pc)", "weight"))
        qty_per_ctn = numeric_value(_get_col(r, "Pcs/Carton", "Qty/CTN"))
        length = numeric_value(_get_col(r, "L"))
        width = numeric_value(_get_col(r, "W"))
        height = numeric_value(_get_col(r, "H"))
        cbm = numeric_value(_get_col(r, "CBM"))
        if cbm == 0 and length and width and height:
            cbm = round(length * width * height, 6)

        package_parts = []
        if qty_per_ctn:
            package_parts.append(f"{qty_per_ctn:g} pcs/carton")
        if length or width or height:
            package_parts.append(f"L{length:g}×W{width:g}×H{height:g} m")
        package_info = "; ".join(package_parts)
        is_pipe = "pipe" in en_desc.lower() and "fitting" not in en_desc.lower()
        unit = "M" if is_pipe else "PCS"
        category = en_desc or "PPR"

        grey_sap = clean_scalar(_get_col(r, "Grey SAP No."))
        green_sap = clean_scalar(_get_col(r, "Green SAP No."))
        variants = [
            {
                "sap": grey_sap,
                "color": "Grey" if grey_sap != green_sap else "Universal",
                "image_url": clean_scalar(_get_col(r, "灰色图片")),
                "material_description": clean_scalar(_get_col(r, "物料描述（灰色）")),
                "price": numeric_value(_get_col(r, "灰色基准价格（USD）/m(pcs)")),
                "price_cny": numeric_value(_get_col(r, "灰色基准价格（CNY）/m(pcs)")),
            },
            {
                "sap": green_sap,
                "color": "Green" if grey_sap != green_sap else "Universal",
                "image_url": clean_scalar(_get_col(r, "Picture ", "Picture")),
                "material_description": clean_scalar(_get_col(r, "物料描述（绿色）")),
                "price": numeric_value(_get_col(r, "绿色基准价格（USD）/m(pcs)")),
                "price_cny": numeric_value(_get_col(r, "绿色基准价格（CNY）/m(pcs)")),
            },
        ]
        for v in variants:
            sap = clean_scalar(v["sap"])
            if not sap or sap in seen:
                continue
            seen.add(sap)
            material_desc = clean_scalar(v.get("material_description"))
            color = clean_scalar(v.get("color"))
            model_parts = [p for p in [size, color] if p]
            model = " / ".join(model_parts) if model_parts else size
            desc_parts = []
            if en_desc:
                desc_parts.append(en_desc)
            if cn_name:
                desc_parts.append(cn_name)
            if size:
                desc_parts.append(f"Size: {size}")
            if color:
                desc_parts.append(f"Color: {color}")
            if weight:
                desc_parts.append(f"Weight: {weight:g} kg/m or pc")
            if material_desc:
                desc_parts.append(material_desc)
            rows.append({
                "sap": sap,
                "product_type": "PPR",
                "subcategory": category,
                "category": _format_category("PPR", category),
                "cn_name": cn_name,
                "en_name": en_desc,
                "model": model,
                "description": " | ".join(desc_parts),
                "color": color,
                "size_mm": size,
                "weight": weight,
                "material_description": material_desc,
                "price": float(v.get("price") or 0),
                "price_cny": float(v.get("price_cny") or 0),
                "stock": 0,
                "packing_volume": cbm,
                "qty_per_ctn": qty_per_ctn,
                "package_length": length,
                "package_width": width,
                "package_height": height,
                "unit": unit,
                "currency": "USD",
                "package_info": package_info,
                "image_url": clean_scalar(v.get("image_url")),
                "active": 1,
            })
    return pd.DataFrame(rows)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        raw = str(col).strip()
        lower = raw.lower()
        mapped = COLUMN_ALIASES.get(raw) or COLUMN_ALIASES.get(lower)
        if mapped:
            rename_map[col] = mapped
    return df.rename(columns=rename_map)


def clean_products_df(df: pd.DataFrame) -> pd.DataFrame:
    if _looks_like_ppr_wide(df):
        df = expand_ppr_wide_df(df)
    elif _looks_like_pvcu_drainage_wide(df):
        df = expand_pvcu_drainage_wide_df(df)
    elif _looks_like_simple_product_catalog(df):
        df = expand_simple_product_catalog_df(df)
    elif _looks_like_usd_price_base(df):
        df = expand_usd_price_base_df(df)
    else:
        df = normalize_columns(df)
        if "material_description" in df.columns and "cn_name" not in df.columns:
            df["cn_name"] = df["material_description"]
        if "size_mm" in df.columns and "model" not in df.columns:
            df["model"] = df["size_mm"]

    if "sap" not in df.columns:
        raise ValueError("导入文件必须包含 SAP 号列，或包含 PPR 宽表字段 Grey SAP No. / Green SAP No.，或包含快速产品大全字段 SAP No. / product describe / TYPE / PRICE，或包含 PVC-U 排水宽表字段，或包含 USD 基准表字段 物料 / 物料描述 / 单价。")

    for col in PRODUCT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df = df[PRODUCT_COLUMNS].copy()
    df["sap"] = df["sap"].apply(clean_scalar)
    df = df[df["sap"].notna() & (df["sap"] != "")]

    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col].apply(clean_scalar), errors="coerce").fillna(0)
    df["stock"] = df["stock"].astype(int)
    df["active"] = df["active"].replace({0: 0}).fillna(1).astype(int)

    df["unit"] = df["unit"].fillna("PC").astype(str).str.strip().replace({"": "PC", "nan": "PC", "None": "PC"})
    df["currency"] = df["currency"].fillna("USD").astype(str).str.strip().replace({"": "USD", "nan": "USD", "None": "USD"})

    for col in TEXT_COLUMNS:
        df[col] = df[col].fillna("").apply(clean_scalar)

    # Build a two-level category system for searching: product_type = PPR/PVC/PE, subcategory = product family.
    df["product_type"] = df.apply(
        lambda r: _normalize_product_type(r.get("product_type")) or _infer_product_type(
            r.get("category"), r.get("subcategory"), r.get("cn_name"), r.get("en_name"), r.get("description"), r.get("material_description")
        ),
        axis=1,
    )
    df["subcategory"] = df.apply(
        lambda r: _clean_subcategory(r.get("subcategory")) or _clean_subcategory(r.get("category")) or _clean_subcategory(r.get("en_name")) or _clean_subcategory(r.get("cn_name")),
        axis=1,
    )
    df["category"] = df.apply(
        lambda r: _format_category(r.get("product_type"), r.get("subcategory"), fallback=r.get("category")),
        axis=1,
    )

    # Keep model aligned with size when model is empty.
    df.loc[df["model"] == "", "model"] = df.loc[df["model"] == "", "size_mm"]

    # Convert file names like 8110010814.png into full public URLs. Empty values stay empty.
    df["image_url"] = df.apply(
        lambda r: normalize_image_url(r.get("image_url"), r.get("sap")) if clean_scalar(r.get("image_url")) else "",
        axis=1,
    )
    return df.drop_duplicates(subset=["sap"], keep="last")


def upsert_products(df: pd.DataFrame) -> int:
    init_db()
    df = clean_products_df(df)
    rows = df.to_dict("records")
    if not rows:
        return 0
    cols = PRODUCT_COLUMNS
    insert_cols = ", ".join(cols)
    values_cols = ", ".join([f":{c}" for c in cols])
    update_cols = ",\n            ".join([
        f"{c}=excluded.{c}" if c != "image_url" else "image_url=COALESCE(NULLIF(excluded.image_url, ''), products.image_url)"
        for c in cols if c != "sap"
    ])
    sql = text(f"""
        INSERT INTO products ({insert_cols})
        VALUES ({values_cols})
        ON CONFLICT(sap) DO UPDATE SET
            {update_cols},
            updated_at=CURRENT_TIMESTAMP
    """)
    with get_engine().begin() as conn:
        for row in rows:
            conn.execute(sql, row)
    return len(rows)


def _like_clause_for_term(term_index: int, term: str, params: dict[str, object]) -> str:
    key = f"kw{term_index}"
    params[key] = f"%{term.lower()}%"
    fields = ["sap", "product_type", "subcategory", "category", "cn_name", "en_name", "model", "description", "color", "size_mm", "material_description"]
    return "(" + " OR ".join([f"LOWER(COALESCE(CAST({field} AS TEXT), '')) LIKE :{key}" for field in fields]) + ")"


def load_products(
    keyword: str = "",
    sap_list: Optional[list[str]] = None,
    category: str = "全部",
    product_type: str = "全部",
    subcategory: str = "全部",
    min_price: Optional[float] = None,
    max_price: Optional[float] = None,
    min_stock: Optional[int] = None,
    has_image: bool = False,
    only_active: bool = True,
    limit: Optional[int] = 200,
) -> pd.DataFrame:
    init_db()
    where: list[str] = []
    params: dict[str, object] = {}
    if only_active:
        where.append("active = 1")

    sap_list = [clean_scalar(x) for x in (sap_list or []) if clean_scalar(x)]
    if sap_list:
        placeholders = []
        for i, sap in enumerate(sap_list[:300]):
            key = f"sap{i}"
            params[key] = sap.lower()
            placeholders.append(f":{key}")
        where.append(f"LOWER(TRIM(CAST(sap AS TEXT))) IN ({', '.join(placeholders)})")
    elif keyword:
        terms = [t for t in re.split(r"\s+", keyword.strip()) if t][:6]
        for i, term in enumerate(terms):
            where.append(_like_clause_for_term(i, term, params))

    if product_type and product_type != "全部":
        where.append("product_type = :product_type")
        params["product_type"] = product_type
    if subcategory and subcategory != "全部":
        where.append("subcategory = :subcategory")
        params["subcategory"] = subcategory
    if category and category != "全部":
        where.append("category = :category")
        params["category"] = category
    if min_price is not None:
        where.append("price >= :min_price")
        params["min_price"] = float(min_price)
    if max_price is not None:
        where.append("price <= :max_price")
        params["max_price"] = float(max_price)
    if min_stock is not None:
        where.append("stock >= :min_stock")
        params["min_stock"] = int(min_stock)
    if has_image:
        where.append("image_url IS NOT NULL AND image_url <> ''")

    select_cols = ", ".join(PRODUCT_COLUMNS)
    sql = f"SELECT {select_cols} FROM products"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY product_type, subcategory, category, sap"
    if limit is not None:
        sql += " LIMIT :limit"
        params["limit"] = int(limit)

    with get_engine().connect() as conn:
        df = pd.read_sql_query(text(sql), conn, params=params)
    if not df.empty:
        df["image_url"] = df.apply(lambda r: normalize_image_url(r.get("image_url"), r.get("sap")) if clean_scalar(r.get("image_url")) else "", axis=1)
    return df


def get_product_types() -> list[str]:
    init_db()
    with get_engine().connect() as conn:
        rows = conn.execute(text("SELECT DISTINCT product_type FROM products WHERE product_type IS NOT NULL AND product_type != '' ORDER BY product_type")).fetchall()
    preferred = ["PPR", "PVC", "PE"]
    values = [row[0] for row in rows]
    ordered = [x for x in preferred if x in values] + [x for x in values if x not in preferred]
    return ["全部"] + ordered


def get_subcategories(product_type: str = "全部") -> list[str]:
    init_db()
    where = "subcategory IS NOT NULL AND subcategory != ''"
    params: dict[str, object] = {}
    if product_type and product_type != "全部":
        where += " AND product_type = :product_type"
        params["product_type"] = product_type
    with get_engine().connect() as conn:
        rows = conn.execute(text(f"SELECT DISTINCT subcategory FROM products WHERE {where} ORDER BY subcategory"), params).fetchall()
    return ["全部"] + [row[0] for row in rows]


def get_categories() -> list[str]:
    init_db()
    with get_engine().connect() as conn:
        rows = conn.execute(text("SELECT DISTINCT category FROM products WHERE category IS NOT NULL AND category != '' ORDER BY category")).fetchall()
    return ["全部"] + [row[0] for row in rows]


def delete_product(sap: str) -> bool:
    init_db()
    with get_engine().begin() as conn:
        result = conn.execute(text("DELETE FROM products WHERE sap = :sap"), {"sap": sap})
    return result.rowcount > 0


def _clean_sap_from_image_filename(filename: str) -> str:
    stem = Path(filename).stem.strip()
    stem = stem.replace("（", "(").replace("）", ")")
    stem = re.sub(r"\s*\(\d+\)$", "", stem).strip()
    stem = re.sub(r"[\s_\-]*(副本|copy)$", "", stem, flags=re.IGNORECASE).strip()
    return stem


def _upload_image_to_supabase_storage(file, sap: str) -> str:
    supabase_url = (_read_secret("SUPABASE_URL") or "").rstrip("/")
    service_key = _read_secret("SUPABASE_SERVICE_ROLE_KEY") or ""
    bucket = _read_secret("SUPABASE_STORAGE_BUCKET", "product-images") or "product-images"
    prefix = (_read_secret("SUPABASE_STORAGE_PREFIX", "") or "").strip("/")

    if not supabase_url or not service_key:
        raise RuntimeError("缺少 SUPABASE_URL 或 SUPABASE_SERVICE_ROLE_KEY。请在 Streamlit Secrets 中配置。")

    suffix = Path(file.name).suffix.lower()
    if suffix not in [".jpg", ".jpeg", ".png", ".webp"]:
        raise ValueError(f"不支持的图片格式：{suffix}")

    sap = str(sap).strip()
    if not sap:
        raise ValueError("图片文件名必须包含 SAP 号，例如 8110022978.png 或 8110022978_副本.png。")

    raw_storage_path = f"{prefix}/{sap}{suffix}" if prefix else f"{sap}{suffix}"
    storage_path = quote(raw_storage_path, safe="/")
    upload_url = f"{supabase_url}/storage/v1/object/{bucket}/{storage_path}"
    public_url = f"{supabase_url}/storage/v1/object/public/{bucket}/{storage_path}"

    content_type_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
    headers = {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Content-Type": content_type_map.get(suffix, "application/octet-stream"),
        "Cache-Control": "3600",
        "x-upsert": "true",
    }
    data = file.getvalue()
    response = requests.post(upload_url, headers=headers, data=data, timeout=45)
    if response.status_code in (400, 409):
        response = requests.put(upload_url, headers=headers, data=data, timeout=45)
    if not response.ok:
        detail = response.text[:500] if response.text else ""
        raise RuntimeError(f"Storage 上传失败：HTTP {response.status_code}，bucket={bucket}，path={raw_storage_path}，详情：{detail}")
    return public_url


def save_uploaded_images(files: Iterable) -> int:
    init_db()
    count = 0
    missing_products: list[str] = []
    with get_engine().begin() as conn:
        for file in files:
            suffix = Path(file.name).suffix.lower()
            if suffix not in [".jpg", ".jpeg", ".png", ".webp"]:
                continue
            sap = _clean_sap_from_image_filename(file.name)
            image_url = _upload_image_to_supabase_storage(file, sap)
            result = conn.execute(
                text("UPDATE products SET image_url=:image_url, updated_at=CURRENT_TIMESTAMP WHERE sap=:sap"),
                {"image_url": image_url, "sap": sap},
            )
            if result.rowcount and result.rowcount > 0:
                count += 1
            else:
                missing_products.append(sap)
    if missing_products:
        preview = ", ".join(missing_products[:10])
        more = "..." if len(missing_products) > 10 else ""
        raise RuntimeError(f"图片已上传，但以下 SAP 在 products 表中不存在，未能绑定：{preview}{more}")
    return count


def save_quote_history(quote_no: str, customer: str, total_amount: float, created_by: str = "") -> None:
    init_db()
    with get_engine().begin() as conn:
        conn.execute(
            text("INSERT INTO quote_history (quote_no, customer, total_amount, created_by) VALUES (:quote_no, :customer, :total_amount, :created_by)"),
            {"quote_no": quote_no, "customer": customer, "total_amount": float(total_amount or 0), "created_by": created_by},
        )


def get_db_status() -> str:
    if DATABASE_URL.startswith("sqlite"):
        return f"本地 SQLite: {DB_PATH}"
    safe = DATABASE_URL.split("@")[-1] if "@" in DATABASE_URL else DATABASE_URL
    return f"云端数据库: {safe}"


# -------- 用户与权限 --------

def get_user(username: str) -> dict | None:
    init_db()
    with get_engine().connect() as conn:
        row = conn.execute(text("SELECT username, password_hash, role, active FROM users WHERE username=:username"), {"username": username}).mappings().fetchone()
    return dict(row) if row else None


def list_users() -> pd.DataFrame:
    init_db()
    with get_engine().connect() as conn:
        return pd.read_sql_query(text("SELECT username, role, active, created_at FROM users ORDER BY created_at DESC"), conn)


def create_or_update_user(username: str, password_hash: str, role: str = "user", active: int = 1) -> None:
    init_db()
    with get_engine().begin() as conn:
        conn.execute(text("""
            INSERT INTO users (username, password_hash, role, active)
            VALUES (:username, :password_hash, :role, :active)
            ON CONFLICT(username) DO UPDATE SET
                password_hash=excluded.password_hash,
                role=excluded.role,
                active=excluded.active,
                updated_at=CURRENT_TIMESTAMP
        """), {"username": username, "password_hash": password_hash, "role": role, "active": int(active)})


def set_user_active(username: str, active: int) -> None:
    init_db()
    with get_engine().begin() as conn:
        conn.execute(text("UPDATE users SET active=:active, updated_at=CURRENT_TIMESTAMP WHERE username=:username"), {"username": username, "active": int(active)})


def has_any_user() -> bool:
    init_db()
    with get_engine().connect() as conn:
        count = conn.execute(text("SELECT COUNT(*) FROM users")).scalar_one()
    return int(count) > 0