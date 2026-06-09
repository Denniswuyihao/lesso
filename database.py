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
    "sap", "category", "cn_name", "en_name", "model", "description",
    "color", "size_mm", "weight", "material_description",
    "price", "price_cny", "stock", "packing_volume", "qty_per_ctn",
    "package_length", "package_width", "package_height",
    "unit", "currency", "package_info", "image_url", "active",
]

TEXT_COLUMNS = [
    "sap", "category", "cn_name", "en_name", "model", "description", "color", "size_mm",
    "material_description", "unit", "currency", "package_info", "image_url",
]
NUMERIC_COLUMNS = [
    "price", "price_cny", "stock", "packing_volume", "qty_per_ctn",
    "package_length", "package_width", "package_height", "weight", "active",
]

EXCEL_ERROR_VALUES = {"#N/A", "#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "nan", "none", "null", ""}

COLUMN_ALIASES = {
    "sap号": "sap", "sap": "sap", "sap code": "sap", "SAP": "sap", "SAP号": "sap",
    "分类": "category", "类别": "category", "category": "category",
    "品名": "cn_name", "中文名": "cn_name", "中文品名": "cn_name", "产品名称": "cn_name", "name": "cn_name",
    "英文名": "en_name", "英文品名": "en_name", "英文产品名": "en_name", "en_name": "en_name", "english name": "en_name",
    "型号": "model", "model": "model", "Size(mm)": "size_mm", "size(mm)": "size_mm", "size": "size_mm", "规格": "size_mm",
    "描述": "description", "产品描述": "description", "description": "description",
    "颜色": "color", "color": "color", "colour": "color",
    "重量": "weight", "weight": "weight", "Weight\n(kg/m or pc)": "weight", "weight(kg/m or pc)": "weight", "kg/m": "weight",
    "物料描述": "material_description", "material_description": "material_description", "物料描述（灰色）": "material_description", "物料描述（绿色）": "material_description",
    "价格": "price", "报价": "price", "fob": "price", "FOB价": "price", "FOB价（USD/PC）": "price", "fob usd": "price",
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
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_sap_lookup ON products(sap)"))
    _DB_INIT_DONE = True


def migrate_products_table() -> None:
    required = {
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
                "category": category,
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
    else:
        df = normalize_columns(df)
        if "size_mm" in df.columns and "model" not in df.columns:
            df["model"] = df["size_mm"]

    if "sap" not in df.columns:
        raise ValueError("导入文件必须包含 SAP 号列，或包含 PPR 宽表字段 Grey SAP No. / Green SAP No.。")

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
    fields = ["sap", "cn_name", "en_name", "model", "description", "category", "color", "size_mm", "material_description"]
    return "(" + " OR ".join([f"LOWER(COALESCE(CAST({field} AS TEXT), '')) LIKE :{key}" for field in fields]) + ")"


def load_products(
    keyword: str = "",
    sap_list: Optional[list[str]] = None,
    category: str = "全部",
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
    sql += " ORDER BY category, sap"
    if limit is not None:
        sql += " LIMIT :limit"
        params["limit"] = int(limit)

    with get_engine().connect() as conn:
        df = pd.read_sql_query(text(sql), conn, params=params)
    if not df.empty:
        df["image_url"] = df.apply(lambda r: normalize_image_url(r.get("image_url"), r.get("sap")) if clean_scalar(r.get("image_url")) else "", axis=1)
    return df


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
