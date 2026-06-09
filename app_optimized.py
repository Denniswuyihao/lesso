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


def split_sap_text(value: str) -> list[str]:
    parts = re.split(r"[\s,，;；]+", str(value or "").strip())
    result: list[str] = []
    seen: set[str] = set()
    for part in parts:
        p = part.strip()
        if not p:
            continue
        key = p.lower()
        if key not in seen:
            result.append(p)
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
    return get_categories()


@st.cache_data(ttl=90, show_spinner=False)
def cached_products(keyword: str, sap_text: str, category: str, min_price: float | None, max_price: float | None, min_stock: int | None, has_image: bool, limit: int) -> pd.DataFrame:
    return load_products(
        keyword=keyword,
        sap_list=split_sap_text(sap_text),
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
    cached_products.clear()


def reset_export_cache() -> None:
    st.session_state["export_files"] = {}


def make_export_signature(items: pd.DataFrame, customer: str, quote_no: str, price_note: str) -> str:
    cols = ["sap", "category", "cn_name", "en_name", "model", "description", "color", "size_mm", "weight", "material_description", "quote_price", "quantity", "amount", "stock", "packing_volume", "qty_per_ctn", "total_volume", "unit", "package_info", "image_url"]
    safe = items.copy()
    for col in cols:
        if col not in safe.columns:
            safe[col] = ""
    payload = "|".join([customer or "", quote_no or "", price_note or ""]) + "\n" + safe[cols].fillna("").astype(str).to_csv(index=False)
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
            st.caption("支持四种导入：1）标准字段；2）PPR宽表；3）PVC-U排水宽表（4m/6m/5.8m物料编码）；4）USD基准表（物料/物料描述/单价）。系统会自动读取 Excel 全部工作表。")
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
                category_new = st.text_input("category 分类")
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
# Search
# -----------------------------------------------------------------------------
section("1. 查找产品", "真实报价常用三种方式：按关键词搜、粘贴客户 SAP 清单、按分类浏览。")

with st.form("search_form"):
    r1c1, r1c2, r1c3 = st.columns([1.4, 1.2, 0.9])
    keyword = r1c1.text_input("关键词", placeholder="例：角阀 DN15 / 水龙头 / W13101B / 806005")
    sap_text = r1c2.text_area("批量 SAP 精确查询", placeholder="客户发来一串 SAP 时粘贴到这里，空格/逗号/换行均可", height=68)
    category = r1c3.selectbox("分类", cached_categories())

    r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns([0.85, 0.85, 0.8, 0.75, 0.75])
    min_price_raw = r2c1.number_input("最低价 USD", min_value=0.0, value=0.0, step=0.1)
    max_price_raw = r2c2.number_input("最高价 USD", min_value=0.0, value=0.0, step=0.1, help="0 表示不限")
    min_stock_raw = r2c3.number_input("最低库存", min_value=0, value=0, step=1)
    has_image = r2c4.checkbox("只看有图", value=False)
    limit = r2c5.selectbox("显示数量", [20, 50, 100, 200], index=1)

    submitted_search = st.form_submit_button("查询产品", type="primary", use_container_width=True)

min_price = min_price_raw if min_price_raw > 0 else None
max_price = max_price_raw if max_price_raw > 0 else None
min_stock = min_stock_raw if min_stock_raw > 0 else None

products = cached_products(keyword, sap_text, category, min_price, max_price, min_stock, has_image, int(limit))

st.caption(f"当前结果：{len(products)} 条｜购物车：{len(st.session_state.cart)} 个 SKU")


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
        edit_df["quantity"] = edit_df["sap"].astype(str).apply(lambda x: int(st.session_state.cart.get(x, {}).get("quantity", 1)))

        display_cols = ["select", "image_preview", "sap", "category", "cn_name", "model", "color", "size_mm", "description", "price", "stock", "packing_volume", "qty_per_ctn", "unit", "quantity"]
        for col in display_cols:
            if col not in edit_df.columns:
                edit_df[col] = ""

        with st.form("product_select_form"):
            edited_products = st.data_editor(
                edit_df[display_cols],
                hide_index=True,
                use_container_width=True,
                height=430,
                disabled=["image_preview", "sap", "category", "cn_name", "model", "color", "size_mm", "description", "price", "stock", "packing_volume", "qty_per_ctn", "unit"],
                column_config={
                    "select": st.column_config.CheckboxColumn("选", width="small"),
                    "image_preview": st.column_config.ImageColumn("图片", width="small"),
                    "sap": st.column_config.TextColumn("SAP", width="medium"),
                    "category": st.column_config.TextColumn("分类", width="medium"),
                    "cn_name": st.column_config.TextColumn("品名", width="large"),
                    "model": st.column_config.TextColumn("型号", width="medium"),
                    "color": st.column_config.TextColumn("颜色", width="small"),
                    "size_mm": st.column_config.TextColumn("规格", width="small"),
                    "description": st.column_config.TextColumn("描述", width="large"),
                    "price": st.column_config.NumberColumn("价格 USD", format="%.4f", width="small"),
                    "stock": st.column_config.NumberColumn("库存", width="small"),
                    "packing_volume": st.column_config.NumberColumn("CBM/件", format="%.4f", width="small"),
                    "qty_per_ctn": st.column_config.NumberColumn("Qty/CTN", format="%.0f", width="small"),
                    "unit": st.column_config.TextColumn("单位", width="small"),
                    "quantity": st.column_config.NumberColumn("数量", min_value=1, step=1, width="small"),
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
                    base["quantity"] = max(int(row.get("quantity") or 1), 1)
                    st.session_state.cart[sap_key] = base
                    added += 1
                reset_export_cache()
                st.success(f"已加入/更新 {added} 个产品。")

    with preview_col:
        section("加入前确认", "核对图片、型号、描述和价格。")
        preview_options = products.copy()
        preview_options["preview_label"] = preview_options.apply(
            lambda r: f"{clean_text(r.get('sap'))} | {clean_text(r.get('cn_name'))} | {clean_text(r.get('model'))} | {clean_text(r.get('color'))}", axis=1
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
                    <div class="label">分类</div><div class="value">{h(preview_row.get('category'))}</div>
                    <div class="label">品名</div><div class="value">{h(preview_row.get('cn_name'))}</div>
                    <div class="label">型号 / 规格 / 颜色</div><div class="value">{h(preview_row.get('model'))} ｜ {h(preview_row.get('size_mm'))} ｜ {h(preview_row.get('color'))}</div>
                    <div class="label">描述</div><div class="value">{h(preview_row.get('description'))}</div>
                    <div class="label">重量 / Qty/CTN</div><div class="value">{qty_text(preview_row.get('weight')) or '-'} kg/m or pc ｜ {qty_text(preview_row.get('qty_per_ctn')) or '-'}</div>
                    <div class="label">价格 / 库存 / 体积</div><div class="value">{float(preview_row.get('price') or 0):.4f} USD ｜ {int(float(preview_row.get('stock') or 0))} ｜ {float(preview_row.get('packing_volume') or 0):.4f} CBM</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("preview_add_form"):
                pq = st.number_input("加入数量", min_value=1, value=max(int(st.session_state.cart.get(preview_sap, {}).get("quantity", 1)), 1), step=1)
                add_one = st.form_submit_button("加入当前确认产品", use_container_width=True)
            if add_one:
                item = preview_row.to_dict()
                item["quantity"] = int(pq)
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
cart_cols = ["sap", "category", "cn_name", "en_name", "model", "description", "color", "size_mm", "weight", "material_description", "price", "price_cny", "quantity", "stock", "packing_volume", "qty_per_ctn", "package_length", "package_width", "package_height", "unit", "currency", "package_info", "image_url"]
for col in cart_cols:
    if col not in cart_df.columns:
        cart_df[col] = ""
cart_editor = cart_df[cart_cols].copy()
cart_editor["quantity"] = pd.to_numeric(cart_editor["quantity"], errors="coerce").fillna(1).astype(int)
cart_editor.insert(0, "remove", False)

with st.form("cart_form"):
    edited_cart = st.data_editor(
        cart_editor[["remove", "sap", "category", "cn_name", "model", "color", "size_mm", "price", "quantity", "stock", "packing_volume", "qty_per_ctn", "unit", "package_info"]],
        hide_index=True,
        use_container_width=True,
        height=260,
        disabled=["sap", "category", "cn_name", "model", "color", "size_mm", "price", "stock", "packing_volume", "qty_per_ctn", "unit", "package_info"],
        column_config={
            "remove": st.column_config.CheckboxColumn("删", width="small"),
            "sap": st.column_config.TextColumn("SAP", width="medium"),
            "category": st.column_config.TextColumn("分类", width="medium"),
            "cn_name": st.column_config.TextColumn("品名", width="large"),
            "model": st.column_config.TextColumn("型号", width="medium"),
            "color": st.column_config.TextColumn("颜色", width="small"),
            "size_mm": st.column_config.TextColumn("规格", width="small"),
            "price": st.column_config.NumberColumn("单价 USD", format="%.4f", width="small"),
            "quantity": st.column_config.NumberColumn("数量", min_value=1, step=1, width="small"),
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
            st.session_state.cart[sap_key]["quantity"] = max(int(row.get("quantity") or 1), 1)
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

rule = PriceRule(mode=mode, percent=percent)
items_raw = pd.DataFrame(list(st.session_state.cart.values()))
items = build_quote_items(items_raw, rule)
summary = quote_summary(items)
price_note = f"{mode} {percent:.2f}%" if mode != "原价" else "原价 / Original Price"

m1, m2, m3, m4 = st.columns(4)
m1.metric("SKU", summary["sku_count"])
m2.metric("总数量", summary["total_qty"])
m3.metric("总金额 USD", f"{summary['total_amount']:.2f}")
m4.metric("总体积 CBM", f"{summary['total_volume']:.4f}")

preview_cols = ["sap", "category", "cn_name", "model", "color", "size_mm", "weight", "quote_price", "quantity", "amount", "stock", "packing_volume", "qty_per_ctn", "total_volume", "unit", "package_info"]
for col in preview_cols:
    if col not in items.columns:
        items[col] = ""
st.dataframe(items[preview_cols], use_container_width=True, hide_index=True, height=245)

if st.button("保存报价历史", use_container_width=True):
    save_quote_history(quote_no, customer, summary["total_amount"], created_by=user["username"])
    st.success("已保存报价历史。")

signature = make_export_signature(items, customer, quote_no, price_note)
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