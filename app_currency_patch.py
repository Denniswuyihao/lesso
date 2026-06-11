from __future__ import annotations

from datetime import datetime
import hashlib
import html
import re
from typing import Any

import pandas as pd
import streamlit as st

from auth import authenticate, ensure_default_admin, hash_password
from database import (
    create_or_update_user,
    delete_product,
    get_categories,
    get_product_types,
    get_subcategories,
    get_db_status,
    get_database_backend_name,
    init_db,
    list_users,
    load_products,
    save_quote_history,
    save_uploaded_images,
    set_user_active,
    upsert_products,
    read_product_import_file,
    read_quote_request_file,
    build_quote_request_preview,
)
from exporters import export_quote_excel, export_quote_pdf, get_pdf_font_status, resolve_image_url
from quote_engine import PriceRule, build_quote_items, quote_summary


# -----------------------------------------------------------------------------
# Init
# -----------------------------------------------------------------------------
st.set_page_config(page_title="LESSO 报价系统", page_icon="📦", layout="wide")

try:
    init_db()
    ensure_default_admin()

    # ===== 应急管理员账号：登录成功后请删除这段 =====
    emergency_user = st.secrets.get("EMERGENCY_ADMIN_USERNAME", "admin")
    emergency_password = st.secrets.get("EMERGENCY_ADMIN_PASSWORD", "")

    if emergency_password:
        create_or_update_user(
            emergency_user,
            hash_password(str(emergency_password)),
            role="admin",
            active=1
        )

except Exception as exc:
    st.error("系统初始化失败。请检查 Supabase DATABASE_URL、products 表结构和网络连接。")
    st.exception(exc)
    st.stop()

except Exception as exc:
    st.error("系统初始化失败。请检查 Supabase DATABASE_URL、products 表结构和网络连接。")
    st.exception(exc)
    st.stop()
except Exception as exc:
    st.error("系统初始化失败。请检查 Supabase DATABASE_URL、products 表结构和网络连接。")
    st.exception(exc)
    st.stop()


# -----------------------------------------------------------------------------
# UI style
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main .block-container { padding-top: 2.2rem !important; padding-bottom: 2rem; max-width: 1360px; }
    header[data-testid="stHeader"] { background: rgba(255,255,255,0.96); box-shadow: none; }
    .topbar { border:1px solid #E5E7EB; border-radius:14px; padding:14px 16px; background:#FFFFFF; margin-bottom:12px; }
    .title { font-size:24px; font-weight:780; color:#111827; line-height:1.3; }
    .sub { font-size:13px; color:#6B7280; margin-top:4px; }
    .section-title { font-size:16px; font-weight:750; color:#111827; margin: 2px 0 2px 0; }
    .section-sub { font-size:12.5px; color:#6B7280; margin-bottom:8px; }
    .note { color:#6B7280; font-size:12.5px; }
    .card { border:1px solid #E5E7EB; border-radius:14px; padding:13px 14px; background:#FFFFFF; margin-bottom:12px; }
    .label { color:#6B7280; font-size:12px; margin-bottom:2px; }
    .value { color:#111827; font-size:14px; margin-bottom:8px; }
    .safe-img-wrap { width:100%; display:flex; align-items:center; justify-content:center; overflow:hidden; border:1px solid #EEF0F3; border-radius:10px; background:#FAFAFA; }
    .safe-img-wrap img { object-fit:contain; display:block; }
    .safe-img-caption { color:#6B7280; font-size:11px; text-align:center; margin-top:3px; }
    div[data-testid="metric-container"] { background:#FFFFFF; border:1px solid #E5E7EB; padding:9px 10px; border-radius:12px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def h(value: Any) -> str:
    return html.escape(str(value or ""), quote=False)


def clean_text(value: Any) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"nan", "none", "null", "#n/a", "#name?", "#value!"} else text


def num_text(value: Any, digits: int = 4) -> str:
    try:
        n = float(value or 0)
        if n == 0:
            return ""
        return f"{n:.{digits}f}"
    except Exception:
        return clean_text(value)


def qty_text(value: Any) -> str:
    try:
        n = float(value or 0)
        if n == 0:
            return ""
        return str(int(n)) if n.is_integer() else f"{n:g}"
    except Exception:
        return clean_text(value)


def qty_number(value: Any, default: float = 1.0) -> float:
    try:
        n = float(value or default)
        return n if n > 0 else float(default)
    except Exception:
        return float(default)


SUPPORTED_CURRENCIES = ["USD", "CNY"]


def normalize_currency(value: Any, default: str = "USD") -> str:
    """Normalize currency text from customer files or database fields."""
    text = str(value or "").strip().upper()
    if text in {"CNY", "RMB", "CNH", "人民币", "￥", "¥"}:
        return "CNY"
    if text in {"USD", "US$", "$", "美金", "美元", "DOLLAR", "DOLLARS"}:
        return "USD"
    return default


def currency_label(currency: str) -> str:
    currency = normalize_currency(currency)
    return "USD 美金" if currency == "USD" else "CNY 人民币"


def currency_symbol(currency: str) -> str:
    return "$" if normalize_currency(currency) == "USD" else "¥"


def price_number(value: Any, default: float = 0.0) -> float:
    """Convert price-like values such as '$12.3', '￥86', '1,280.00' to float."""
    if value is None:
        return float(default)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "#n/a", "#value!", "#name?"}:
        return float(default)
    text = re.sub(r"[,$￥¥\s]", "", text)
    text = text.replace("USD", "").replace("usd", "").replace("RMB", "").replace("rmb", "").replace("CNY", "").replace("cny", "")
    try:
        return float(text)
    except Exception:
        return float(default)


def detect_currency_from_df(df: pd.DataFrame, default: str = "USD") -> str:
    """Detect whether an uploaded customer file is more likely USD or CNY.

    It checks column names first, then samples cell text. This is intentionally
    conservative: when unclear, default to USD because the original system used USD.
    """
    if df is None or df.empty:
        return default

    names = " ".join(str(c).lower() for c in df.columns)
    cny_keys = ["人民币", "rmb", "cny", "含税价", "内销价", "国内价", "¥", "￥"]
    usd_keys = ["美元", "美金", "usd", "us$", "fob", "dollar", "$"]

    cny_score = sum(3 for k in cny_keys if k.lower() in names)
    usd_score = sum(3 for k in usd_keys if k.lower() in names)

    try:
        sample = " ".join(
            df.head(80).astype(str).fillna("").to_numpy().ravel().tolist()
        ).lower()
        cny_score += sum(1 for k in cny_keys if k.lower() in sample)
        usd_score += sum(1 for k in usd_keys if k.lower() in sample)
    except Exception:
        pass

    if cny_score > usd_score:
        return "CNY"
    if usd_score > cny_score:
        return "USD"
    return default


def unit_price_for_currency(row: Any, currency: str) -> float:
    """Return the correct unit price from a product row according to quotation currency."""
    currency = normalize_currency(currency)
    if isinstance(row, dict):
        getter = row.get
    else:
        getter = row.get

    if currency == "CNY":
        primary = price_number(getter("price_cny", 0))
        fallback = price_number(getter("price", 0))
    else:
        primary = price_number(getter("price", 0))
        fallback = price_number(getter("price_cny", 0))
    return primary if primary > 0 else fallback


def apply_currency_prices(df: pd.DataFrame, currency: str) -> pd.DataFrame:
    """Add display quotation price columns without destroying original USD/CNY columns."""
    safe = df.copy()
    for col in ["price", "price_cny"]:
        if col not in safe.columns:
            safe[col] = 0
    safe["quote_currency"] = normalize_currency(currency)
    safe["quote_base_price"] = safe.apply(lambda r: unit_price_for_currency(r, currency), axis=1)
    return safe


def prepare_quote_input_for_currency(items_raw: pd.DataFrame, currency: str) -> pd.DataFrame:
    """Feed selected currency price into quote_engine, whose existing logic reads `price`."""
    safe = apply_currency_prices(items_raw, currency)
    safe["price_original_usd"] = safe.get("price", 0)
    safe["price"] = safe["quote_base_price"]
    safe["currency"] = normalize_currency(currency)
    return safe

def split_sap_text(value: str) -> list[str]:
    """Split customer SAP text into clean SAP codes.

    Supports newline, space, comma, semicolon and slash-separated lists, for example:
    8010018062 / 8010018066 / 8010018069
    """
    text = str(value or "").strip()
    if not text:
        return []

    # Convert common separators to spaces. The previous version did not handle '/',
    # so a copied list like 8010018062/8010018066 could be treated as one SAP.
    text = re.sub(r"[/\\|、，,;；]+", " ", text)
    parts = re.split(r"\s+", text)

    result: list[str] = []
    seen: set[str] = set()
    for part in parts:
        p = part.strip().strip("'\"`，,;；/\\|、")
        if not p:
            continue
        # If a token still contains extra text, extract SAP-like digit strings first.
        candidates = re.findall(r"\d{6,20}", p) or [p]
        for c in candidates:
            c = c.strip()
            if c.endswith(".0") and c[:-2].isdigit():
                c = c[:-2]
            key = c.lower()
            if c and key not in seen:
                result.append(c)
                seen.add(key)
    return result[:500]


def render_topbar() -> None:
    st.markdown(
        """
        <div class="topbar">
            <div class="title">LESSO 外贸报价系统</div>
            <div class="sub">按真实报价流程设计：查产品 → 核图片/型号/价格 → 加入报价单 → 调价 → 导出 Excel / PDF。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str, subtitle: str = "") -> None:
    st.markdown(f"<div class='section-title'>{h(title)}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='section-sub'>{h(subtitle)}</div>", unsafe_allow_html=True)


def render_safe_image(image_url: str | None, width: int = 120, height: int | None = None, caption: str = "") -> None:
    """Render image as browser HTML, so bad URLs will not crash Streamlit."""
    url = str(image_url or "").strip()
    if not url:
        st.caption("无图片")
        return
    if not (url.startswith("http://") or url.startswith("https://") or url.startswith("data:image/")):
        st.caption("图片路径异常")
        return
    h_px = int(height or width)
    safe_url = html.escape(url, quote=True)
    safe_caption = html.escape(caption or "", quote=False)
    st.markdown(
        f"""
        <div class="safe-img-wrap" style="height:{h_px}px; min-height:{h_px}px;">
            <img src="{safe_url}" style="max-width:{int(width)}px; max-height:{h_px - 8}px;" loading="lazy" referrerpolicy="no-referrer" />
        </div>
        {f'<div class="safe-img-caption">{safe_caption}</div>' if safe_caption else ''}
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(ttl=90, show_spinner=False)
def cached_categories() -> list[str]:
    try:
        return get_categories()
    except Exception as exc:
        st.warning(f"分类读取失败，请检查 Supabase 数据库连接：{exc}")
        return ["全部"]


@st.cache_data(ttl=90, show_spinner=False)
def cached_product_types() -> list[str]:
    try:
        return get_product_types()
    except Exception as exc:
        st.warning(f"产品大类读取失败，请检查 Supabase 数据库连接：{exc}")
        return ["全部"]


@st.cache_data(ttl=90, show_spinner=False)
def cached_subcategories(product_type: str = "全部") -> list[str]:
    try:
        return get_subcategories(product_type)
    except Exception as exc:
        st.warning(f"产品小类读取失败，请检查 Supabase 数据库连接：{exc}")
        return ["全部"]


@st.cache_data(ttl=90, show_spinner=False)
def cached_products(
    keyword: str,
    sap_text: str,
    product_type: str,
    subcategory: str,
    category: str,
    min_price: float | None,
    max_price: float | None,
    min_stock: int | None,
    has_image: bool,
    limit: int,
) -> pd.DataFrame:
    return load_products(
        keyword=keyword,
        sap_list=split_sap_text(sap_text),
        product_type=product_type,
        subcategory=subcategory,
        category=category,
        min_price=min_price,
        max_price=max_price,
        min_stock=min_stock,
        has_image=has_image,
        limit=limit,
    )


@st.cache_data(ttl=3600, show_spinner=False)
def cached_pdf_font_status() -> str:
    return get_pdf_font_status()


def clear_product_cache() -> None:
    cached_categories.clear()
    cached_product_types.clear()
    cached_subcategories.clear()
    cached_products.clear()


def reset_export_cache() -> None:
    st.session_state["export_files"] = {}


def make_export_signature(items: pd.DataFrame, customer: str, quote_no: str, price_note: str, quote_currency: str) -> str:
    cols = ["sap", "product_type", "subcategory", "category", "cn_name", "en_name", "model", "description", "color", "size_mm", "weight", "material_description", "quote_currency", "quote_base_price", "quote_price", "quantity", "amount", "stock", "packing_volume", "qty_per_ctn", "total_volume", "unit", "package_info", "image_url"]
    safe = items.copy()
    for col in cols:
        if col not in safe.columns:
            safe[col] = ""
    payload = "|".join([customer or "", quote_no or "", price_note or "", normalize_currency(quote_currency)]) + "\n" + safe[cols].fillna("").astype(str).to_csv(index=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


# -----------------------------------------------------------------------------
# Login
# -----------------------------------------------------------------------------
def login_page() -> None:
    render_topbar()
    c1, c2, c3 = st.columns([1, 1.05, 1])
    with c2:
        st.subheader("登录")
        with st.form("login_form"):
            username = st.text_input("用户名")
            password = st.text_input("密码", type="password")
            submitted = st.form_submit_button("登录", type="primary", use_container_width=True)
        if submitted:
            user = authenticate(username, password)
            if user:
                st.session_state["user"] = user
                st.session_state.setdefault("cart", {})
                st.rerun()
            else:
                st.error("用户名或密码错误，或账号已停用。")


if "user" not in st.session_state:
    login_page()
    st.stop()

user = st.session_state["user"]
is_admin = user.get("role") == "admin"
st.session_state.setdefault("cart", {})
st.session_state.setdefault("export_files", {})
st.session_state.setdefault("quote_currency", "USD")
st.session_state.setdefault("quote_no_default", f"QT-{datetime.now().strftime('%Y%m%d-%H%M')}")

render_topbar()


# -----------------------------------------------------------------------------
# Sidebar / admin
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"**{h(user['username'])}**｜{h(user['role'])}")
    st.caption(f"数据库：{get_database_backend_name()}")
    st.caption(get_db_status())
    st.caption(cached_pdf_font_status())

    current_currency = normalize_currency(st.session_state.get("quote_currency", "USD"))
    selected_currency = st.selectbox(
        "本次报价币种",
        SUPPORTED_CURRENCIES,
        index=SUPPORTED_CURRENCIES.index(current_currency),
        format_func=currency_label,
        help="控制产品展示价、购物车单价、报价预览、汇总金额和导出文件金额。",
    )
    if selected_currency != current_currency:
        st.session_state["quote_currency"] = selected_currency
        reset_export_cache()

    if st.button("退出登录", use_container_width=True):
        st.session_state.pop("user", None)
        st.session_state.pop("cart", None)
        st.session_state.pop("export_files", None)
        st.rerun()

    st.divider()
    if is_admin:
        admin_panel_open = st.toggle("后台管理", value=False, help="导入产品、上传图片、管理账号。平时报价建议关闭。")
    else:
        admin_panel_open = False
        st.info("普通用户只能搜索产品、生成报价。")

    if is_admin and admin_panel_open:
        st.subheader("后台管理")

        with st.expander("导入/更新产品表", expanded=True):
            st.caption("支持五种导入：1）标准字段；2）快速产品大全表（SAP No./product describe/TYPE/PRICE）；3）PPR宽表；4）PVC-U排水宽表；5）USD基准表。系统会自动读取 Excel 全部工作表，并按 PPR/PVC/PE 自动归类。")
            product_file = st.file_uploader("上传 Excel / CSV", type=["xlsx", "xls", "csv"])
            if product_file is not None:
                try:
                    df_import = read_product_import_file(product_file)
                    count = upsert_products(df_import)
                    clear_product_cache()
                    reset_export_cache()
                    st.success(f"已导入/更新 {count} 个产品。")
                except Exception as exc:
                    st.error(f"导入失败：{exc}")

        with st.expander("批量上传产品图片", expanded=False):
            st.caption("图片名建议为 SAP号.png。系统会上传到 Supabase Storage，并写入 products.image_url。")
            image_files = st.file_uploader(
                "选择图片文件",
                type=["jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
            )
            if image_files:
                try:
                    count = save_uploaded_images(image_files)
                    clear_product_cache()
                    reset_export_cache()
                    st.success(f"已上传并绑定 {count} 张图片。")
                except Exception as exc:
                    st.error(f"图片上传失败：{exc}")

        with st.expander("手工新增/修改商品", expanded=False):
            with st.form("single_product_form", clear_on_submit=False):
                sap = st.text_input("sap *")
                product_type_new = st.text_input("product_type 产品大类", placeholder="PPR / PVC / PE")
                subcategory_new = st.text_input("subcategory 产品小类", placeholder="例：PVC-U给水直管, 1.0MPa")
                category_new = st.text_input("category 分类", help="可不填；系统会按 产品大类 / 产品小类 自动生成。")
                cn_name = st.text_input("cn_name 中文品名")
                en_name = st.text_input("en_name 英文品名", value="")
                model = st.text_input("model / size 型号规格")
                description = st.text_area("description 描述", height=70)
                c1, c2, c3, c4 = st.columns(4)
                color = c1.text_input("color 颜色", value="")
                size_mm = c2.text_input("size_mm 规格", value="")
                weight = c3.number_input("weight kg/m or pc", min_value=0.0, step=0.001, format="%.4f")
                qty_per_ctn = c4.number_input("Qty/CTN", min_value=0.0, step=1.0, format="%.0f")
                p1, p2, p3, p4 = st.columns(4)
                price = p1.number_input("price USD", min_value=0.0, step=0.01, format="%.4f")
                price_cny = p2.number_input("price CNY", min_value=0.0, step=0.01, format="%.4f")
                stock = p3.number_input("stock 库存", min_value=0, step=1)
                packing_volume = p4.number_input("packing_volume CBM", min_value=0.0, step=0.0001, format="%.4f")
                d1, d2, d3 = st.columns(3)
                package_length = d1.number_input("L", min_value=0.0, step=0.01, format="%.4f")
                package_width = d2.number_input("W", min_value=0.0, step=0.01, format="%.4f")
                package_height = d3.number_input("H", min_value=0.0, step=0.01, format="%.4f")
                c4, c5, c6 = st.columns(3)
                unit = c4.text_input("unit 单位", value="PC")
                currency = c5.text_input("currency 币种", value="USD")
                active = c6.selectbox("active 状态", [1, 0], format_func=lambda x: "启用" if x == 1 else "隐藏")
                material_description = st.text_input("material_description 物料描述")
                package_info = st.text_input("package_info 包装信息")
                image_url = st.text_input("image_url 图片URL/文件名", help="可填完整 URL，也可填 8110010814.png。")
                submitted_single = st.form_submit_button("保存商品", type="primary", use_container_width=True)

            if submitted_single:
                if not sap.strip():
                    st.error("sap 不能为空。")
                else:
                    one = pd.DataFrame([
                        {
                            "sap": sap,
                            "product_type": product_type_new,
                            "subcategory": subcategory_new,
                            "category": category_new,
                            "cn_name": cn_name,
                            "en_name": en_name,
                            "model": model,
                            "description": description,
                            "color": color,
                            "size_mm": size_mm,
                            "weight": weight,
                            "material_description": material_description,
                            "price": price,
                            "price_cny": price_cny,
                            "stock": stock,
                            "packing_volume": packing_volume,
                            "qty_per_ctn": qty_per_ctn,
                            "package_length": package_length,
                            "package_width": package_width,
                            "package_height": package_height,
                            "unit": unit,
                            "currency": currency,
                            "package_info": package_info,
                            "image_url": image_url,
                            "active": active,
                        }
                    ])
                    upsert_products(one)
                    clear_product_cache()
                    reset_export_cache()
                    st.success("商品已保存。")

        with st.expander("删除商品 / 账号管理", expanded=False):
            del_sap = st.text_input("删除 SAP")
            if st.button("删除该商品", use_container_width=True):
                if del_sap.strip() and delete_product(del_sap.strip()):
                    clear_product_cache()
                    st.session_state.cart.pop(del_sap.strip(), None)
                    reset_export_cache()
                    st.success("已删除商品。")
                else:
                    st.warning("没有找到该商品。")

            st.divider()
            if st.button("显示用户列表", use_container_width=True):
                st.session_state["show_users_table"] = True
            if st.session_state.get("show_users_table"):
                st.dataframe(list_users(), hide_index=True, use_container_width=True)
            with st.form("create_user_form"):
                new_user = st.text_input("用户名")
                new_password = st.text_input("密码", type="password")
                new_role = st.selectbox("权限", ["user", "admin"], format_func=lambda x: "普通用户" if x == "user" else "管理员")
                create_user_submitted = st.form_submit_button("创建/重置账号", use_container_width=True)
            if create_user_submitted:
                if not new_user.strip() or not new_password:
                    st.error("用户名和密码不能为空。")
                else:
                    create_or_update_user(new_user.strip(), hash_password(new_password), role=new_role, active=1)
                    st.success("账号已创建/重置。")
            disable_user = st.text_input("停用/启用用户名")
            u1, u2 = st.columns(2)
            if u1.button("停用", use_container_width=True):
                if disable_user.strip():
                    set_user_active(disable_user.strip(), 0)
                    st.success("已停用。")
            if u2.button("启用", use_container_width=True):
                if disable_user.strip():
                    set_user_active(disable_user.strip(), 1)
                    st.success("已启用。")



# -----------------------------------------------------------------------------
# Quote request import / preview
# -----------------------------------------------------------------------------
section("0. 从客户文档导入报价草稿", "上传客户 Excel/CSV：系统按 SAP/型号/数量对比数据库补齐信息；重复 SAP 自动合并数量；先预览再加入购物车。")
with st.expander("上传客户询价表 / 旧报价单并预览", expanded=False):
    request_file = st.file_uploader("选择客户文档", type=["xlsx", "xls", "csv"], key="quote_request_file")
    st.checkbox("启用 AI 辅助分析未匹配项（接口预留，需要 OPENAI_API_KEY）", value=False, key="use_ai_review")
    if st.button("读取并生成预览", use_container_width=True) and request_file is not None:
        try:
            req_df = read_quote_request_file(request_file)
            detected_currency = detect_currency_from_df(req_df, default=normalize_currency(st.session_state.get("quote_currency", "USD")))
            st.session_state["quote_currency"] = detected_currency
            st.session_state["detected_request_currency"] = detected_currency
            st.session_state["quote_request_preview"] = build_quote_request_preview(req_df)
            reset_export_cache()
            st.success(f"已生成预览：{len(st.session_state['quote_request_preview'])} 行。系统识别币种：{currency_label(detected_currency)}。如识别不准，可在左侧“本次报价币种”手动切换。")
        except Exception as exc:
            st.error(f"读取客户文档失败：{exc}")
    preview_df = st.session_state.get("quote_request_preview")
    quote_currency = normalize_currency(st.session_state.get("quote_currency", "USD"))
    if isinstance(preview_df, pd.DataFrame) and not preview_df.empty:
        preview_show_df = apply_currency_prices(preview_df, quote_currency)
        show_cols=["select","match_status","match_reason","input_sap","input_model","input_description","requested_quantity","sap","product_type","subcategory","category","cn_name","model","quote_base_price","quote_currency","price","price_cny","unit","stock","packing_volume","qty_per_ctn"]
        for col in show_cols:
            if col not in preview_show_df.columns: preview_show_df[col]=""
        edited_preview=st.data_editor(
            preview_show_df[show_cols],
            hide_index=True,
            use_container_width=True,
            height=320,
            disabled=[c for c in show_cols if c not in ["select","requested_quantity"]],
            column_config={
                "select":st.column_config.CheckboxColumn("加入",width="small"),
                "requested_quantity":st.column_config.NumberColumn("数量",min_value=0.0,step=1.0,format="%.4f"),
                "quote_base_price":st.column_config.NumberColumn(f"报价单价 {quote_currency}",format="%.4f",width="small"),
                "quote_currency":st.column_config.TextColumn("币种",width="small"),
                "price":st.column_config.NumberColumn("USD价",format="%.4f",width="small"),
                "price_cny":st.column_config.NumberColumn("人民币价",format="%.4f",width="small"),
            },
            key="quote_request_preview_editor"
        )
        if st.button("把勾选且已匹配的产品加入报价购物车", type="primary", use_container_width=True):
            selected=edited_preview[(edited_preview["select"]==True)&(edited_preview["sap"].astype(str).str.strip()!="")].copy()
            if selected.empty: st.warning("没有可加入的已匹配产品。")
            else:
                for _,row in selected.iterrows():
                    sap_key=clean_text(row.get("sap")); item=preview_show_df[preview_show_df["sap"].astype(str)==sap_key].iloc[0].to_dict(); new_qty=qty_number(row.get("requested_quantity"),1.0)
                    item["quote_currency"] = quote_currency
                    item["quantity"] = new_qty
                    if sap_key in st.session_state.cart: st.session_state.cart[sap_key]["quantity"]=qty_number(st.session_state.cart[sap_key].get("quantity"),0.0)+new_qty
                    else: st.session_state.cart[sap_key]=item
                reset_export_cache(); st.success(f"已加入/合并 {len(selected)} 个 SKU 到购物车，当前币种：{currency_label(quote_currency)}。")
    elif request_file is None:
        st.info("上传后会先预览，不会直接生成报价单。")

# -----------------------------------------------------------------------------
# Search
# -----------------------------------------------------------------------------
section("1. 查找产品", "真实报价常用三种方式：按关键词搜、粘贴客户 SAP 清单、按分类浏览。")
quote_currency = normalize_currency(st.session_state.get("quote_currency", "USD"))

with st.form("search_form"):
    r1c1, r1c2, r1c3, r1c4 = st.columns([1.25, 1.1, 0.65, 1.0])
    keyword = r1c1.text_input("关键词", placeholder="例：角阀 DN15 / 水龙头 / W13101B / 806005")
    sap_text = r1c2.text_area("批量 SAP 精确查询", placeholder="客户发来一串 SAP 时粘贴到这里，空格/逗号/斜杠/换行均可", height=68)
    product_type = r1c3.selectbox("产品大类", cached_product_types(), help="系统会从导入表里的 TYPE / 描述自动识别 PPR、PVC、PE。")
    subcategory = r1c4.selectbox("产品小类", cached_subcategories(product_type), help="例如 PVC-U给水直管, 1.0MPa / Tee PPR / Socket PPR。")
    category = "全部"

    r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns([0.85, 0.85, 0.8, 0.75, 0.75])
    min_price_raw = r2c1.number_input(f"最低价 {quote_currency}", min_value=0.0, value=0.0, step=0.1)
    max_price_raw = r2c2.number_input(f"最高价 {quote_currency}", min_value=0.0, value=0.0, step=0.1, help="0 表示不限")
    min_stock_raw = r2c3.number_input("最低库存", min_value=0, value=0, step=1)
    has_image = r2c4.checkbox("只看有图", value=False)
    limit = r2c5.selectbox("显示数量", [20, 50, 100, 200], index=1)

    submitted_search = st.form_submit_button("查询产品", type="primary", use_container_width=True)

min_price = min_price_raw if min_price_raw > 0 else None
max_price = max_price_raw if max_price_raw > 0 else None
min_stock = min_stock_raw if min_stock_raw > 0 else None

# load_products 原本按 USD price 过滤；CNY 模式下先取数，再按 price_cny 在前端过滤。
db_min_price = min_price if quote_currency == "USD" else None
db_max_price = max_price if quote_currency == "USD" else None
products = cached_products(keyword, sap_text, product_type, subcategory, category, db_min_price, db_max_price, min_stock, has_image, int(limit))
products = apply_currency_prices(products, quote_currency)
if quote_currency == "CNY" and not products.empty:
    if min_price is not None:
        products = products[pd.to_numeric(products["quote_base_price"], errors="coerce").fillna(0) >= min_price]
    if max_price is not None:
        products = products[pd.to_numeric(products["quote_base_price"], errors="coerce").fillna(0) <= max_price]

st.caption(f"当前结果：{len(products)} 条｜购物车：{len(st.session_state.cart)} 个 SKU｜当前币种：{currency_label(quote_currency)}")


# -----------------------------------------------------------------------------
# Product selection
# -----------------------------------------------------------------------------
section("2. 选择产品", "先看左侧表格批量加入；不确定时看右侧“加入前确认”。")

if products.empty:
    st.info("没有匹配产品。可以减少条件，或检查产品表是否已导入。")
else:
    results_col, preview_col = st.columns([2.15, 0.85], gap="large")

    product_map = {clean_text(row.get("sap")): row.to_dict() for _, row in products.iterrows()}

    with results_col:
        edit_df = products.copy()
        edit_df["image_preview"] = edit_df.apply(lambda r: resolve_image_url(r.get("image_url"), str(r.get("sap", ""))), axis=1)
        edit_df["select"] = edit_df["sap"].astype(str).apply(lambda x: x in st.session_state.cart)
        edit_df["quantity"] = edit_df["sap"].astype(str).apply(lambda x: qty_number(st.session_state.cart.get(x, {}).get("quantity", 1)))

        display_cols = ["select", "image_preview", "sap", "product_type", "subcategory", "category", "cn_name", "model", "color", "size_mm", "description", "quote_base_price", "quote_currency", "price", "price_cny", "stock", "packing_volume", "qty_per_ctn", "unit", "quantity"]
        for col in display_cols:
            if col not in edit_df.columns:
                edit_df[col] = ""

        with st.form("product_select_form"):
            edited_products = st.data_editor(
                edit_df[display_cols],
                hide_index=True,
                use_container_width=True,
                height=430,
                disabled=["image_preview", "sap", "product_type", "subcategory", "category", "cn_name", "model", "color", "size_mm", "description", "quote_base_price", "quote_currency", "price", "price_cny", "stock", "packing_volume", "qty_per_ctn", "unit"],
                column_config={
                    "select": st.column_config.CheckboxColumn("选", width="small"),
                    "image_preview": st.column_config.ImageColumn("图片", width="small"),
                    "sap": st.column_config.TextColumn("SAP", width="medium"),
                    "product_type": st.column_config.TextColumn("大类", width="small"),
                    "subcategory": st.column_config.TextColumn("小类", width="medium"),
                    "category": st.column_config.TextColumn("分类", width="medium"),
                    "cn_name": st.column_config.TextColumn("品名", width="large"),
                    "model": st.column_config.TextColumn("型号", width="medium"),
                    "color": st.column_config.TextColumn("颜色", width="small"),
                    "size_mm": st.column_config.TextColumn("规格", width="small"),
                    "description": st.column_config.TextColumn("描述", width="large"),
                    "quote_base_price": st.column_config.NumberColumn(f"报价价 {quote_currency}", format="%.4f", width="small"),
                    "quote_currency": st.column_config.TextColumn("币种", width="small"),
                    "price": st.column_config.NumberColumn("USD价", format="%.4f", width="small"),
                    "price_cny": st.column_config.NumberColumn("人民币价", format="%.4f", width="small"),
                    "stock": st.column_config.NumberColumn("库存", width="small"),
                    "packing_volume": st.column_config.NumberColumn("CBM/件", format="%.4f", width="small"),
                    "qty_per_ctn": st.column_config.NumberColumn("Qty/CTN", format="%.0f", width="small"),
                    "unit": st.column_config.TextColumn("单位", width="small"),
                    "quantity": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0, format="%.4f", width="small"),
                },
                key="product_editor",
            )
            add_selected = st.form_submit_button("加入/更新已选产品", type="primary", use_container_width=True)

        if add_selected:
            selected = edited_products[edited_products["select"] == True].copy()
            if selected.empty:
                st.warning("请先勾选产品。")
            else:
                added = 0
                for _, row in selected.iterrows():
                    sap_key = clean_text(row.get("sap"))
                    base = product_map.get(sap_key)
                    if not base:
                        continue
                    base["quantity"] = qty_number(row.get("quantity"), 1.0)
                    base["quote_currency"] = quote_currency
                    base["quote_base_price"] = unit_price_for_currency(base, quote_currency)
                    st.session_state.cart[sap_key] = base
                    added += 1
                reset_export_cache()
                st.success(f"已加入/更新 {added} 个产品。")

    with preview_col:
        section("加入前确认", "核对图片、型号、描述和价格。")
        preview_options = products.copy()
        preview_options["preview_label"] = preview_options.apply(
            lambda r: f"{clean_text(r.get('sap'))} | {clean_text(r.get('product_type'))} | {clean_text(r.get('subcategory'))} | {clean_text(r.get('cn_name'))} | {clean_text(r.get('model'))}", axis=1
        )
        selected_preview = st.selectbox("选择产品", preview_options["preview_label"].tolist(), key="preview_select")
        if selected_preview:
            preview_sap = selected_preview.split(" | ")[0]
            preview_row = preview_options[preview_options["sap"].astype(str) == preview_sap].iloc[0]
            p_img = resolve_image_url(preview_row.get("image_url"), preview_sap)
            render_safe_image(p_img, width=260, height=220, caption=f"SAP: {preview_sap}" if p_img else "")
            st.markdown(
                f"""
                <div class="card">
                    <div class="label">产品大类 / 产品小类</div><div class="value">{h(preview_row.get('product_type'))} ｜ {h(preview_row.get('subcategory'))}</div>
                    <div class="label">分类</div><div class="value">{h(preview_row.get('category'))}</div>
                    <div class="label">品名</div><div class="value">{h(preview_row.get('cn_name'))}</div>
                    <div class="label">型号 / 规格 / 颜色</div><div class="value">{h(preview_row.get('model'))} ｜ {h(preview_row.get('size_mm'))} ｜ {h(preview_row.get('color'))}</div>
                    <div class="label">描述</div><div class="value">{h(preview_row.get('description'))}</div>
                    <div class="label">重量 / Qty/CTN</div><div class="value">{qty_text(preview_row.get('weight')) or '-'} kg/m or pc ｜ {qty_text(preview_row.get('qty_per_ctn')) or '-'}</div>
                    <div class="label">价格 / 库存 / 体积</div><div class="value">{unit_price_for_currency(preview_row, quote_currency):.4f} {h(quote_currency)} ｜ {int(float(preview_row.get('stock') or 0))} ｜ {float(preview_row.get('packing_volume') or 0):.4f} CBM</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("preview_add_form"):
                pq = st.number_input(
    "加入数量",
    min_value=1.0,
    value=float(max(qty_number(st.session_state.cart.get(preview_sap, {}).get("quantity", 1), 1.0), 1.0)),
    step=1.0,
    format="%.4f",
)
                add_one = st.form_submit_button("加入当前确认产品", use_container_width=True)
            if add_one:
                item = preview_row.to_dict()
                item["quantity"] = qty_number(pq, 1.0)
                item["quote_currency"] = quote_currency
                item["quote_base_price"] = unit_price_for_currency(item, quote_currency)
                st.session_state.cart[preview_sap] = item
                reset_export_cache()
                st.success(f"已加入：{preview_sap}")


# -----------------------------------------------------------------------------
# Cart
# -----------------------------------------------------------------------------
section("3. 报价购物车", "集中修改数量，确认后进入报价规则。")
cart = st.session_state.cart
if not cart:
    st.warning("报价单为空。请先加入产品。")
    st.stop()

cart_df = pd.DataFrame(list(cart.values()))
cart_cols = ["sap", "product_type", "subcategory", "category", "cn_name", "en_name", "model", "description", "color", "size_mm", "weight", "material_description", "price", "price_cny", "quantity", "stock", "packing_volume", "qty_per_ctn", "package_length", "package_width", "package_height", "unit", "currency", "package_info", "image_url"]
for col in cart_cols:
    if col not in cart_df.columns:
        cart_df[col] = ""
quote_currency = normalize_currency(st.session_state.get("quote_currency", "USD"))
cart_editor = apply_currency_prices(cart_df[cart_cols].copy(), quote_currency)
cart_editor["quantity"] = pd.to_numeric(cart_editor["quantity"], errors="coerce").fillna(1.0)
cart_editor.insert(0, "remove", False)

with st.form("cart_form"):
    edited_cart = st.data_editor(
        cart_editor[["remove", "sap", "product_type", "subcategory", "category", "cn_name", "model", "color", "size_mm", "quote_base_price", "quote_currency", "price", "price_cny", "quantity", "stock", "packing_volume", "qty_per_ctn", "unit", "package_info"]],
        hide_index=True,
        use_container_width=True,
        height=260,
        disabled=["sap", "product_type", "subcategory", "category", "cn_name", "model", "color", "size_mm", "quote_base_price", "quote_currency", "price", "price_cny", "stock", "packing_volume", "qty_per_ctn", "unit", "package_info"],
        column_config={
            "remove": st.column_config.CheckboxColumn("删", width="small"),
            "sap": st.column_config.TextColumn("SAP", width="medium"),
            "product_type": st.column_config.TextColumn("大类", width="small"),
            "subcategory": st.column_config.TextColumn("小类", width="medium"),
            "category": st.column_config.TextColumn("分类", width="medium"),
            "cn_name": st.column_config.TextColumn("品名", width="large"),
            "model": st.column_config.TextColumn("型号", width="medium"),
            "color": st.column_config.TextColumn("颜色", width="small"),
            "size_mm": st.column_config.TextColumn("规格", width="small"),
            "quote_base_price": st.column_config.NumberColumn(f"报价单价 {quote_currency}", format="%.4f", width="small"),
            "quote_currency": st.column_config.TextColumn("币种", width="small"),
            "price": st.column_config.NumberColumn("USD价", format="%.4f", width="small"),
            "price_cny": st.column_config.NumberColumn("人民币价", format="%.4f", width="small"),
            "quantity": st.column_config.NumberColumn("数量", min_value=0.0, step=1.0, format="%.4f", width="small"),
            "stock": st.column_config.NumberColumn("库存", width="small"),
            "packing_volume": st.column_config.NumberColumn("CBM/件", format="%.4f", width="small"),
            "qty_per_ctn": st.column_config.NumberColumn("Qty/CTN", format="%.0f", width="small"),
        },
        key="cart_editor",
    )
    c1, c2 = st.columns([1, 1])
    update_cart = c1.form_submit_button("更新购物车", type="primary", use_container_width=True)
    clear_cart = c2.form_submit_button("清空购物车", use_container_width=True)

if update_cart:
    for _, row in edited_cart.iterrows():
        sap_key = clean_text(row.get("sap"))
        if bool(row.get("remove")):
            st.session_state.cart.pop(sap_key, None)
        elif sap_key in st.session_state.cart:
            st.session_state.cart[sap_key]["quantity"] = qty_number(row.get("quantity"), 1.0)
    reset_export_cache()
    st.success("购物车已更新。")
    st.rerun()

if clear_cart:
    st.session_state.cart = {}
    reset_export_cache()
    st.rerun()


# -----------------------------------------------------------------------------
# Quote and export
# -----------------------------------------------------------------------------
section("4. 报价规则与导出", "导出 Excel 时图片会嵌入表格图片列。")
with st.form("quote_rule_form"):
    q1, q2, q3, q4 = st.columns([0.9, 0.8, 1.2, 1.2])
    mode = q1.selectbox("价格模式", ["原价", "加点", "打折"])
    percent = q2.number_input("比例 %", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
    customer = q3.text_input("客户名称", value="")
    quote_no = q4.text_input("报价单号", value=st.session_state["quote_no_default"])
    apply_rule = st.form_submit_button("更新报价预览", type="primary", use_container_width=True)

if apply_rule:
    reset_export_cache()

quote_currency = normalize_currency(st.session_state.get("quote_currency", "USD"))
rule = PriceRule(mode=mode, percent=percent)
items_raw = pd.DataFrame(list(st.session_state.cart.values()))
items_input = prepare_quote_input_for_currency(items_raw, quote_currency)
items = build_quote_items(items_input, rule)
items["quote_currency"] = quote_currency
items["currency_symbol"] = currency_symbol(quote_currency)
summary = quote_summary(items)
price_note = (f"{mode} {percent:.2f}% / Currency: {quote_currency}" if mode != "原价" else f"原价 / Original Price / Currency: {quote_currency}")

m1, m2, m3, m4 = st.columns(4)
m1.metric("SKU", summary["sku_count"])
m2.metric("总数量", summary["total_qty"])
m3.metric(f"总金额 {quote_currency}", f"{summary['total_amount']:.2f}")
m4.metric("总体积 CBM", f"{summary['total_volume']:.4f}")

preview_cols = ["sap", "product_type", "subcategory", "category", "cn_name", "model", "color", "size_mm", "weight", "quote_currency", "quote_base_price", "quote_price", "quantity", "amount", "stock", "packing_volume", "qty_per_ctn", "total_volume", "unit", "package_info"]
for col in preview_cols:
    if col not in items.columns:
        items[col] = ""
st.dataframe(items[preview_cols], use_container_width=True, hide_index=True, height=245)

if st.button("保存报价历史", use_container_width=True):
    save_quote_history(quote_no, customer, summary["total_amount"], created_by=user["username"])
    st.success("已保存报价历史。")

signature = make_export_signature(items, customer, quote_no, price_note, quote_currency)
export_state = st.session_state["export_files"]
if export_state.get("signature") != signature:
    st.session_state["export_files"] = {"signature": signature}
    export_state = st.session_state["export_files"]

e1, e2, e3, e4 = st.columns(4)
with e1:
    if st.button("生成中文 Excel", use_container_width=True):
        with st.spinner("正在生成中文 Excel..."):
            export_state["excel_zh"] = export_quote_excel(items, customer, quote_no, price_note, lang="zh")
    if "excel_zh" in export_state:
        st.download_button("下载中文 Excel", export_state["excel_zh"], file_name=f"{quote_no}_中文报价单.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with e2:
    if st.button("生成英文 Excel", use_container_width=True):
        with st.spinner("正在生成英文 Excel..."):
            export_state["excel_en"] = export_quote_excel(items, customer, quote_no, price_note, lang="en")
    if "excel_en" in export_state:
        st.download_button("下载英文 Excel", export_state["excel_en"], file_name=f"{quote_no}_English_Quotation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with e3:
    if st.button("生成中文 PDF", use_container_width=True):
        with st.spinner("正在生成中文 PDF..."):
            export_state["pdf_zh"] = export_quote_pdf(items, customer, quote_no, price_note, lang="zh")
    if "pdf_zh" in export_state:
        st.download_button("下载中文 PDF", export_state["pdf_zh"], file_name=f"{quote_no}_中文报价单.pdf", mime="application/pdf", use_container_width=True)
with e4:
    if st.button("生成英文 PDF", use_container_width=True):
        with st.spinner("正在生成英文 PDF..."):
            export_state["pdf_en"] = export_quote_pdf(items, customer, quote_no, price_note, lang="en")
    if "pdf_en" in export_state:
        st.download_button("下载英文 PDF", export_state["pdf_en"], file_name=f"{quote_no}_English_Quotation.pdf", mime="application/pdf", use_container_width=True)