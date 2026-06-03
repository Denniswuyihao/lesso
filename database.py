from __future__ import annotations
import streamlit as st
import os
import re
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import quote

import pandas as pd
import requests
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


def _has_streamlit_secrets_file() -> bool:
    """Only touch st.secrets when a secrets.toml file actually exists.
    This avoids Streamlit's "No secrets files found" error during local use.
    """
    candidates = [
        Path.home() / ".streamlit" / "secrets.toml",
        Path.cwd() / ".streamlit" / "secrets.toml",
        BASE_DIR / ".streamlit" / "secrets.toml",
    ]
    return any(path.exists() for path in candidates)


def _read_secret(name: str, default: str | None = None) -> str | None:
    # 1) Prefer normal environment variables / .env
    value = os.getenv(name)
    if value:
        return value

    # 2) Use Streamlit secrets only when the file exists locally.
    #    Without this check, Streamlit may stop the app with:
    #    "No secrets files found. Valid paths are..."
    if _has_streamlit_secrets_file():
        try:
            import streamlit as st
            value = st.secrets.get(name)
            if value:
                return str(value)
        except Exception:
            pass

    # 3) Fall back to defaults, usually local SQLite.
    return default


def get_database_url() -> str:
    """
    优先读取 Streamlit Cloud Secrets 里的 DATABASE_URL。
    如果没有，再读取本地环境变量。
    如果都没有，才回退到本地 SQLite。
    """

    url = None

    # 1. 优先读取 Streamlit Cloud Secrets
    try:
        url = st.secrets.get("DATABASE_URL")
    except Exception:
        url = None

    # 2. 如果 Streamlit Secrets 没有，再读取系统环境变量
    if not url:
        url = os.getenv("DATABASE_URL")

    # 3. 如果云端数据库地址仍然没有，才使用本地 SQLite
    if not url:
        return f"sqlite:///{DB_PATH}"

    url = str(url).strip()

    # 4. 兼容 postgres:// 写法
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)

    # 5. 如果是 postgresql://，但没有指定 psycopg2，也自动补上
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

_ENGINE: Engine | None = None
_DB_INIT_DONE = False

PRODUCT_COLUMNS = [
    "sap", "category", "cn_name", "en_name", "model", "description", "price", "stock",
    "packing_volume", "unit", "currency", "package_info", "image_url", "active"
]

COLUMN_ALIASES = {
    "sap号": "sap", "sap": "sap", "sap code": "sap", "SAP": "sap", "SAP号": "sap",
    "分类": "category", "类别": "category", "category": "category",
    "品名": "cn_name", "中文名": "cn_name", "中文品名": "cn_name", "产品名称": "cn_name", "name": "cn_name",
    "英文名": "en_name", "英文品名": "en_name", "英文产品名": "en_name", "en_name": "en_name", "english name": "en_name",
    "型号": "model", "model": "model",
    "描述": "description", "产品描述": "description", "description": "description",
    "价格": "price", "报价": "price", "fob": "price", "FOB价": "price", "FOB价（USD/PC）": "price", "fob usd": "price",
    "库存": "stock", "stock": "stock",
    "包装体积": "packing_volume", "体积": "packing_volume", "cbm": "packing_volume", "CBM": "packing_volume",
    "单位": "unit", "unit": "unit",
    "币种": "currency", "currency": "currency",
    "包装": "package_info", "包装信息": "package_info", "package": "package_info", "package_info": "package_info",
    "图片": "image_url", "图片链接": "image_url", "图片URL": "image_url", "image": "image_url", "image_url": "image_url", "image link": "image_url",
    "状态": "active", "active": "active", "是否启用": "active",
}


def get_engine() -> Engine:
    global _ENGINE
    if _ENGINE is None:
        connect_args = {}
        if DATABASE_URL.startswith("sqlite"):
            connect_args = {"check_same_thread": False}
        _ENGINE = create_engine(DATABASE_URL, future=True, pool_pre_ping=True, connect_args=connect_args)
    return _ENGINE


def _dialect() -> str:
    return get_engine().dialect.name


def init_db() -> None:
    """Initialize schema only once per running app process.

    Streamlit reruns the script frequently; without this guard every product query
    will also execute CREATE TABLE / schema checks, which slows cloud deployments.
    """
    global _DB_INIT_DONE
    if _DB_INIT_DONE:
        return

    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS products (
                sap VARCHAR(80) PRIMARY KEY,
                category TEXT,
                cn_name TEXT,
                en_name TEXT,
                model TEXT,
                description TEXT,
                price REAL DEFAULT 0,
                stock INTEGER DEFAULT 0,
                packing_volume REAL DEFAULT 0,
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
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_active_category_sap ON products(active, category, sap)"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_category_sap ON products(category, sap)"))
        conn.execute(text("CREATE INDEX IF NOT EXISTS idx_products_sap_lookup ON products(sap)"))
    migrate_products_table()
    _DB_INIT_DONE = True


def migrate_products_table() -> None:
    """兼容旧数据库：给旧表补英文名等新字段。"""
    engine = get_engine()
    required = {
        "en_name": "TEXT DEFAULT ''",
        "image_url": "TEXT DEFAULT ''",
        "active": "INTEGER DEFAULT 1",
    }
    with engine.begin() as conn:
        if _dialect() == "sqlite":
            rows = conn.execute(text("PRAGMA table_info(products)")).fetchall()
            existing = {row[1] for row in rows}
        elif _dialect() == "postgresql":
            rows = conn.execute(text("""
                SELECT column_name FROM information_schema.columns
                WHERE table_name = 'products' AND table_schema = current_schema()
            """)).fetchall()
            existing = {row[0] for row in rows}
        else:
            existing = set()
        for col, definition in required.items():
            if col not in existing:
                conn.execute(text(f"ALTER TABLE products ADD COLUMN {col} {definition}"))


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
    df = normalize_columns(df)
    if "sap" not in df.columns:
        raise ValueError("导入文件必须包含 SAP 号列，可命名为 sap、SAP、sap号、SAP号。")

    for col in PRODUCT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df = df[PRODUCT_COLUMNS].copy()
    df["sap"] = df["sap"].astype(str).str.strip()
    df = df[df["sap"].notna() & (df["sap"] != "") & (df["sap"].str.lower() != "nan")]

    df["price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0)
    df["stock"] = pd.to_numeric(df["stock"], errors="coerce").fillna(0).astype(int)
    df["packing_volume"] = pd.to_numeric(df["packing_volume"], errors="coerce").fillna(0)
    df["active"] = pd.to_numeric(df["active"], errors="coerce").fillna(1).astype(int)
    df["unit"] = df["unit"].fillna("PC").astype(str).str.strip().replace({"": "PC", "nan": "PC"})
    df["currency"] = df["currency"].fillna("USD").astype(str).str.strip().replace({"": "USD", "nan": "USD"})

    text_cols = ["category", "cn_name", "en_name", "model", "description", "package_info", "image_url"]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str).str.strip().replace({"nan": ""})

    return df.drop_duplicates(subset=["sap"], keep="last")


def upsert_products(df: pd.DataFrame) -> int:
    init_db()
    df = clean_products_df(df)
    rows = df.to_dict("records")
    sql = text("""
        INSERT INTO products (
            sap, category, cn_name, en_name, model, description, price, stock,
            packing_volume, unit, currency, package_info, image_url, active
        ) VALUES (
            :sap, :category, :cn_name, :en_name, :model, :description, :price, :stock,
            :packing_volume, :unit, :currency, :package_info, :image_url, :active
        )
        ON CONFLICT(sap) DO UPDATE SET
            category=excluded.category,
            cn_name=excluded.cn_name,
            en_name=excluded.en_name,
            model=excluded.model,
            description=excluded.description,
            price=excluded.price,
            stock=excluded.stock,
            packing_volume=excluded.packing_volume,
            unit=excluded.unit,
            currency=excluded.currency,
            package_info=excluded.package_info,
            image_url=COALESCE(NULLIF(excluded.image_url, ''), products.image_url),
            active=excluded.active,
            updated_at=CURRENT_TIMESTAMP
    """)
    with get_engine().begin() as conn:
        for row in rows:
            conn.execute(sql, row)
    return len(rows)


def load_products(
    keyword: str = "",
    category: str = "全部",
    min_price: Optional[float] = None,
    max_price: Optional[float] = None,
    min_stock: Optional[int] = None,
    only_active: bool = True,
    limit: Optional[int] = 200,
) -> pd.DataFrame:
    init_db()
    where: list[str] = []
    params: dict[str, object] = {}
    if only_active:
        where.append("active = 1")
    if keyword:
        where.append("(sap LIKE :kw OR cn_name LIKE :kw OR en_name LIKE :kw OR model LIKE :kw OR description LIKE :kw)")
        params["kw"] = f"%{keyword}%"
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

    sql = """
        SELECT
            sap, category, cn_name, en_name, model, description, price, stock,
            packing_volume, unit, currency, package_info, image_url, active
        FROM products
    """
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY category, sap"
    if limit is not None:
        sql += " LIMIT :limit"
        params["limit"] = int(limit)

    with get_engine().connect() as conn:
        return pd.read_sql_query(text(sql), conn, params=params)


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


def _read_storage_secret(name: str, default: str | None = None) -> str | None:
    """Read Supabase Storage config from Streamlit Secrets or environment variables."""
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    value = os.getenv(name)
    if value:
        return value
    return default


def _clean_sap_from_image_filename(filename: str) -> str:
    """Extract a clean SAP code from image filename.

    Examples:
    8110022978.png -> 8110022978
    8060050433_副本.png -> 8060050433
    8060050433 copy.png -> 8060050433
    8060050433(1).png -> 8060050433
    """
    stem = Path(filename).stem.strip()
    stem = stem.replace("（", "(").replace("）", ")")
    stem = re.sub(r"\s*\(\d+\)$", "", stem).strip()
    stem = re.sub(r"[\s_\-]*(副本|copy)$", "", stem, flags=re.IGNORECASE).strip()
    return stem


def _upload_image_to_supabase_storage(file, sap: str) -> str:
    """Upload image to Supabase Storage and return its public URL.

    Default cloud image location:
        bucket: product-images
        path: products/<SAP>.<ext>

    Configure in Streamlit Secrets:
        SUPABASE_STORAGE_BUCKET = "product-images"
        SUPABASE_STORAGE_PREFIX = "products"
    """
    supabase_url = (_read_storage_secret("SUPABASE_URL") or "").rstrip("/")
    service_key = _read_storage_secret("SUPABASE_SERVICE_ROLE_KEY") or ""
    bucket = _read_storage_secret("SUPABASE_STORAGE_BUCKET", "product-images") or "product-images"
    prefix = (_read_storage_secret("SUPABASE_STORAGE_PREFIX", "products") or "products").strip("/")

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

    content_type_map = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
    }
    headers = {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Content-Type": content_type_map.get(suffix, "application/octet-stream"),
        "Cache-Control": "3600",
        "x-upsert": "true",
    }

    data = file.getvalue()
    response = requests.post(upload_url, headers=headers, data=data, timeout=45)

    # Some Supabase Storage versions return 400 for existing object even when upsert handling differs.
    # Try PUT once before failing, then expose response text so debugging is precise.
    if response.status_code in (400, 409):
        response = requests.put(upload_url, headers=headers, data=data, timeout=45)

    if not response.ok:
        detail = response.text[:500] if response.text else ""
        raise RuntimeError(f"Storage 上传失败：HTTP {response.status_code}，bucket={bucket}，path={raw_storage_path}，详情：{detail}")

    return public_url


def save_uploaded_images(files: Iterable) -> int:
    """Upload images to Supabase Storage and bind public URLs to products.image_url.

    文件名建议是 SAP号.png；如果是 SAP号_副本.png，也会自动识别为 SAP号。
    返回成功绑定到 products 表的图片数量。
    """
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
    safe = DATABASE_URL.split("@")[1] if "@" in DATABASE_URL else DATABASE_URL
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
