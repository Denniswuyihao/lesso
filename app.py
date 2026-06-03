from __future__ import annotations

from datetime import datetime
import hashlib

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
)
from exporters import export_quote_excel, export_quote_pdf, get_pdf_font_status
from quote_engine import PriceRule, build_quote_items, quote_summary

st.set_page_config(page_title="外贸报价表自动生成系统", layout="wide", page_icon="📦")

# -----------------------------------------------------------------------------
# One-time initialization
# -----------------------------------------------------------------------------
init_db()
ensure_default_admin()

# -----------------------------------------------------------------------------
# UI styles: conservative business look, avoid decorative overhead.
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 3.2rem !important;
        padding-bottom: 2rem;
        max-width: 1360px;
    }
    header[data-testid="stHeader"] {
        background: rgba(255,255,255,0.96);
        box-shadow: none;
    }
    .app-header {
        padding: 2px 0 12px 0;
        margin: 0 0 12px 0;
        border-bottom: 1px solid #EEF0F3;
        background: #FFFFFF;
    }
    .app-title {
        margin: 0;
        padding: 0;
        line-height: 1.35;
        font-size: 24px;
        font-weight: 700;
        color: #1F2937;
        letter-spacing: .1px;
    }
    .app-subtitle {
        margin: 5px 0 0 0;
        padding: 0;
        line-height: 1.45;
        color: #6B7280;
        font-size: 13px;
    }
    .step-title {
        font-size: 17px;
        font-weight: 700;
        color: #111827;
        margin: 0 0 6px 0;
    }
    .step-subtitle {
        font-size: 12.5px;
        color: #6B7280;
        margin-bottom: 10px;
    }
    .small-badge {
        display: inline-block;
        padding: 4px 9px;
        border-radius: 6px;
        background: #F7F9FC;
        color: #1F4E78;
        font-size: 12px;
        font-weight: 600;
        border: 1px solid #D9E2F3;
    }
    .audit-note {
        padding: 9px 11px;
        border-left: 3px solid #1F4E78;
        background: #F7F9FC;
        color: #374151;
        font-size: 12.8px;
        border-radius: 6px;
        margin: 8px 0 10px 0;
    }
    .product-row {
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 9px 10px;
        margin-bottom: 8px;
        background: #FFFFFF;
    }
    .product-title {
        font-weight: 700;
        color: #17365D;
        font-size: 14px;
        line-height: 1.35;
        margin-bottom: 4px;
    }
    .product-meta {
        color: #4B5563;
        font-size: 12px;
        line-height: 1.45;
    }
    .preview-card {
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 14px;
        background: #FFFFFF;
    }
    .preview-label {
        color: #6B7280;
        font-size: 12px;
        margin-bottom: 2px;
    }
    .preview-value {
        color: #111827;
        font-size: 14px;
        margin-bottom: 8px;
    }
    div[data-testid="metric-container"] {
        background: #FAFAFA;
        border: 1px solid #E5E7EB;
        padding: 10px;
        border-radius: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_app_header(subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="app-header">
            <div class="app-title">外贸报价表自动生成系统</div>
            <div class="app-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(ttl=60, show_spinner=False)
def cached_get_categories() -> list[str]:
    return get_categories()


@st.cache_data(ttl=120, show_spinner=False)
def cached_load_products(keyword: str, category: str, min_price: float | None, max_price: float | None, min_stock: int | None, limit: int) -> pd.DataFrame:
    return load_products(
        keyword=keyword,
        category=category,
        min_price=min_price,
        max_price=max_price,
        min_stock=min_stock,
        limit=limit,
    )


@st.cache_data(ttl=3600, show_spinner=False)
def cached_pdf_font_status() -> str:
    return get_pdf_font_status()


def clear_product_cache() -> None:
    cached_get_categories.clear()
    cached_load_products.clear()


def make_export_signature(items: pd.DataFrame, customer: str, quote_no: str, price_note: str) -> str:
    cols = [
        "sap", "category", "cn_name", "en_name", "model", "description",
        "base_price", "quote_price", "quantity", "amount", "stock",
        "packing_volume", "total_volume", "package_info", "image_url",
    ]
    safe = items.copy()
    for col in cols:
        if col not in safe.columns:
            safe[col] = ""
    payload = "|".join([customer or "", quote_no or "", price_note or ""]) + "\n" + safe[cols].fillna("").astype(str).to_csv(index=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def safe_text(value: object) -> str:
    return str(value or "").strip()


def login_page() -> None:
    render_app_header("产品数据库 · 云端图片 · 价格筛选 · 报价购物车 · 中英双版本 Excel/PDF 导出")
    c1, c2, c3 = st.columns([1, 1.1, 1])
    with c2:
        st.subheader("账号登录")
        st.caption("正式使用请创建独立管理员账号，并停用默认 admin。")
        with st.form("login_form"):
            username = st.text_input("用户名")
            password = st.text_input("密码", type="password")
            submitted = st.form_submit_button("登录", use_container_width=True)
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

render_app_header("稳定云端版 · Supabase 数据库与图片 · A4 竖版中英报价单 · 点击生成文件")

# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------
with st.sidebar:
    st.caption(f"当前数据库：{get_database_backend_name()}")
    st.markdown(f"<span class='small-badge'>当前用户：{user['username']}｜{user['role']}</span>", unsafe_allow_html=True)
    st.caption(get_db_status())
    st.caption(cached_pdf_font_status())
    if st.button("退出登录", use_container_width=True):
        st.session_state.pop("user", None)
        st.session_state.pop("cart", None)
        st.rerun()

    admin_panel_open = False
    if is_admin:
        admin_panel_open = st.toggle("显示后台管理", value=False, help="平时关闭后台管理，可减少页面渲染并提升报价速度。")

    if is_admin and admin_panel_open:
        st.divider()
        st.header("后台管理员")

        with st.expander("导入/更新产品数据库", expanded=True):
            product_file = st.file_uploader("导入产品表 Excel/CSV", type=["xlsx", "xls", "csv"])
            if product_file is not None:
                try:
                    if product_file.name.lower().endswith(".csv"):
                        df_import = pd.read_csv(product_file)
                    else:
                        df_import = pd.read_excel(product_file)
                    count = upsert_products(df_import)
                    clear_product_cache()
                    st.success(f"已导入/更新 {count} 个产品")
                except Exception as e:
                    st.error(f"导入失败：{e}")

            image_files = st.file_uploader(
                "批量上传产品图片到 Supabase Storage，文件名建议为 SAP号.png；SAP号_副本.png 可自动识别",
                type=["jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
            )
            if image_files:
                try:
                    count = save_uploaded_images(image_files)
                    clear_product_cache()
                    st.success(f"已上传并绑定 {count} 张图片到云端。")
                except Exception as e:
                    st.error(f"图片上传失败：{e}")

        with st.expander("手工新增/修改商品"):
            with st.form("single_product_form"):
                sap = st.text_input("SAP号 *")
                category_new = st.text_input("分类")
                cn_name = st.text_input("中文品名")
                en_name = st.text_input("英文品名")
                model = st.text_input("型号")
                description = st.text_area("描述", height=70)
                c1, c2 = st.columns(2)
                price = c1.number_input("价格", min_value=0.0, step=0.01, format="%.4f")
                stock = c2.number_input("库存", min_value=0, step=1)
                c3, c4 = st.columns(2)
                packing_volume = c3.number_input("包装体积 CBM/件", min_value=0.0, step=0.0001, format="%.4f")
                unit = c4.text_input("单位", value="PC")
                c5, c6 = st.columns(2)
                currency = c5.text_input("币种", value="USD")
                active = c6.selectbox("状态", [1, 0], format_func=lambda x: "启用" if x == 1 else "隐藏")
                package_info = st.text_input("包装信息")
                image_url = st.text_input("图片URL（可选，Supabase Storage public URL）")
                submitted = st.form_submit_button("保存商品", use_container_width=True)
            if submitted:
                if not sap.strip():
                    st.error("SAP号不能为空。")
                else:
                    one = pd.DataFrame([{
                        "sap": sap,
                        "category": category_new,
                        "cn_name": cn_name,
                        "en_name": en_name,
                        "model": model,
                        "description": description,
                        "price": price,
                        "stock": stock,
                        "packing_volume": packing_volume,
                        "unit": unit,
                        "currency": currency,
                        "package_info": package_info,
                        "image_url": image_url,
                        "active": active,
                    }])
                    upsert_products(one)
                    clear_product_cache()
                    st.success("商品已保存。")

        with st.expander("删除商品"):
            del_sap = st.text_input("输入要删除的 SAP 号")
            if st.button("删除该商品", use_container_width=True):
                if del_sap.strip() and delete_product(del_sap.strip()):
                    clear_product_cache()
                    st.success("已删除商品。")
                else:
                    st.warning("没有找到该商品。")

        with st.expander("账号管理"):
            users_df = list_users()
            st.dataframe(users_df, hide_index=True, use_container_width=True)
            with st.form("create_user_form"):
                new_user = st.text_input("新用户名")
                new_password = st.text_input("新密码", type="password")
                new_role = st.selectbox("权限", ["user", "admin"], format_func=lambda x: "普通用户" if x == "user" else "管理员")
                create_user_submitted = st.form_submit_button("创建/重置账号", use_container_width=True)
            if create_user_submitted:
                if not new_user.strip() or not new_password:
                    st.error("用户名和密码不能为空。")
                else:
                    create_or_update_user(new_user.strip(), hash_password(new_password), role=new_role, active=1)
                    st.success("账号已创建/重置。")
            disable_user = st.text_input("停用/启用用户名")
            dc1, dc2 = st.columns(2)
            if dc1.button("停用", use_container_width=True):
                if disable_user.strip():
                    set_user_active(disable_user.strip(), 0)
                    st.success("已停用。")
            if dc2.button("启用", use_container_width=True):
                if disable_user.strip():
                    set_user_active(disable_user.strip(), 1)
                    st.success("已启用。")
    elif not is_admin:
        st.info("普通用户权限：只能筛选产品并生成报价表，不能新增、修改或删除商品。")

# -----------------------------------------------------------------------------
# Product filter
# -----------------------------------------------------------------------------
st.markdown("<div class='step-title'>1. 产品筛选与图片确认</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='step-subtitle'>输入筛选条件后点击查询。选品和数量在表单内批量提交，避免每勾选一次就重跑页面。</div>",
    unsafe_allow_html=True,
)

if "filter_state" not in st.session_state:
    st.session_state["filter_state"] = {
        "keyword": "",
        "category": "全部",
        "result_limit": 20,
        "min_price": 0.0,
        "max_price": 0.0,
        "min_stock": 0,
        "show_images": True,
    }

category_options = cached_get_categories()
if st.session_state["filter_state"].get("category", "全部") not in category_options:
    st.session_state["filter_state"]["category"] = "全部"

with st.container(border=True):
    with st.form("product_filter_form"):
        f1, f2, f3 = st.columns([1.55, 1, 0.95])
        keyword_input = f1.text_input(
            "搜索 SAP / 中文名 / 英文名 / 型号 / 描述",
            value=st.session_state["filter_state"].get("keyword", ""),
        )
        category_input = f2.selectbox(
            "分类",
            category_options,
            index=category_options.index(st.session_state["filter_state"].get("category", "全部")),
        )
        limit_options = [10, 20, 50, 100]
        current_limit = int(st.session_state["filter_state"].get("result_limit", 20))
        if current_limit not in limit_options:
            current_limit = 20
        result_limit_input = f3.selectbox("最多显示", limit_options, index=limit_options.index(current_limit))

        f4, f5, f6, f7 = st.columns([1, 1, 1, 0.8])
        min_price_input = f4.number_input("不低于价格", min_value=0.0, value=float(st.session_state["filter_state"].get("min_price", 0.0)), step=0.1)
        max_price_input_form = f5.number_input("不高于价格，0 表示不限", min_value=0.0, value=float(st.session_state["filter_state"].get("max_price", 0.0)), step=0.1)
        min_stock_input_form = f6.number_input("最低库存", min_value=0, value=int(st.session_state["filter_state"].get("min_stock", 0)), step=1)
        show_images_input = f7.checkbox("显示图片", value=bool(st.session_state["filter_state"].get("show_images", True)))
        submitted_filter = st.form_submit_button("查询 / 刷新产品", type="primary", use_container_width=True)

    if submitted_filter:
        st.session_state["filter_state"] = {
            "keyword": keyword_input,
            "category": category_input,
            "result_limit": int(result_limit_input),
            "min_price": float(min_price_input),
            "max_price": float(max_price_input_form),
            "min_stock": int(min_stock_input_form),
            "show_images": bool(show_images_input),
        }

filter_state = st.session_state["filter_state"]
keyword = filter_state["keyword"]
category = filter_state["category"]
result_limit = int(filter_state["result_limit"])
min_price = float(filter_state["min_price"])
max_price_input = float(filter_state["max_price"])
min_stock_input = int(filter_state["min_stock"])
show_images = bool(filter_state.get("show_images", True))

products = cached_load_products(
    keyword=keyword,
    category=category,
    min_price=min_price if min_price > 0 else None,
    max_price=max_price_input if max_price_input > 0 else None,
    min_stock=min_stock_input if min_stock_input > 0 else None,
    limit=result_limit,
)

status_cols = st.columns(4)
status_cols[0].metric("当前显示", len(products))
status_cols[1].metric("购物车 SKU", len(st.session_state.cart))
status_cols[2].metric("筛选分类", category)
status_cols[3].metric("显示上限", result_limit)

st.markdown(
    "<div class='audit-note'>准确性原则：所有加入购物车的产品都按 SAP 从当前查询结果重新读取完整字段；选品表单只提交“是否加入”和“数量”，避免隐藏字段被误改。</div>",
    unsafe_allow_html=True,
)

if products.empty:
    st.info("暂无产品。管理员可以在左侧导入产品表，普通用户需要等待管理员维护数据库。")
else:
    st.markdown("<div class='step-title'>2. 勾选产品并加入报价单</div>", unsafe_allow_html=True)
    table_col, confirm_col = st.columns([2.25, 1.05], gap="large")

    with table_col:
        with st.form("product_selection_form"):
            for idx, row in products.iterrows():
                sap_key = safe_text(row.get("sap"))
                image_url = safe_text(row.get("image_url"))
                name_cn = safe_text(row.get("cn_name"))
                name_en = safe_text(row.get("en_name"))
                model = safe_text(row.get("model"))
                category_value = safe_text(row.get("category"))
                price_value = float(row.get("price") or 0)
                stock_value = int(row.get("stock") or 0)
                cbm_value = float(row.get("packing_volume") or 0)
                currency = safe_text(row.get("currency")) or "USD"
                default_qty = max(int(st.session_state.cart.get(sap_key, {}).get("quantity", 1)), 1)

                st.markdown("<div class='product-row'>", unsafe_allow_html=True)
                c_img, c_info, c_num, c_qty, c_sel = st.columns([0.75, 3.1, 1.1, 0.9, 0.75], vertical_alignment="center")
                with c_img:
                    if show_images and image_url:
                        st.image(image_url, width=92)
                    else:
                        st.caption("无图" if not image_url else "图片关闭")
                with c_info:
                    st.markdown(f"<div class='product-title'>{name_cn or name_en or sap_key}</div>", unsafe_allow_html=True)
                    st.markdown(
                        f"<div class='product-meta'>SAP：{sap_key}｜型号：{model or '-'}｜分类：{category_value or '-'}<br/>英文：{name_en or '-'}</div>",
                        unsafe_allow_html=True,
                    )
                with c_num:
                    st.markdown(f"<div class='product-meta'>价格<br/><b>{price_value:.2f} {currency}</b></div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='product-meta'>库存：{stock_value}<br/>CBM：{cbm_value:.4f}</div>", unsafe_allow_html=True)
                with c_qty:
                    st.number_input("数量", min_value=1, value=default_qty, step=1, key=f"qty_select_{sap_key}", label_visibility="collapsed")
                with c_sel:
                    st.checkbox("加入", value=sap_key in st.session_state.cart, key=f"sel_select_{sap_key}")
                st.markdown("</div>", unsafe_allow_html=True)

            submitted_add = st.form_submit_button("加入/更新已勾选产品", type="primary", use_container_width=True)

        if submitted_add:
            added = 0
            for _, row in products.iterrows():
                sap_key = safe_text(row.get("sap"))
                if not sap_key:
                    continue
                if bool(st.session_state.get(f"sel_select_{sap_key}", False)):
                    item = row.to_dict()
                    item["quantity"] = max(int(st.session_state.get(f"qty_select_{sap_key}", 1) or 1), 1)
                    st.session_state.cart[sap_key] = item
                    added += 1
            if added:
                st.success(f"已加入/更新 {added} 个产品。")
            else:
                st.warning("请先勾选产品。")

    with confirm_col:
        st.markdown("<div class='step-title'>加入前确认</div>", unsafe_allow_html=True)
        st.markdown("<div class='step-subtitle'>用于核对图片、品名、型号、价格、库存和体积。</div>", unsafe_allow_html=True)
        preview_options = products.copy()
        preview_options["preview_label"] = preview_options.apply(
            lambda r: f"{r.get('sap', '')} | {r.get('cn_name', '') or r.get('en_name', '')} | {r.get('model', '')}",
            axis=1,
        )
        selected_preview = st.selectbox("选择产品", preview_options["preview_label"].tolist(), key="preview_product_select")
        if selected_preview:
            preview_sap = selected_preview.split(" | ")[0]
            preview_row = preview_options[preview_options["sap"].astype(str) == preview_sap].iloc[0]
            p_img = safe_text(preview_row.get("image_url"))
            st.markdown("<div class='preview-label'>产品图片</div>", unsafe_allow_html=True)
            if p_img:
                st.image(p_img, caption=f"SAP: {preview_row.get('sap', '')}", width=250)
            else:
                st.info("该产品暂未上传图片。")

            st.markdown("<div class='preview-card'>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>中文品名</div><div class='preview-value'>{preview_row.get('cn_name', '')}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>英文品名</div><div class='preview-value'>{preview_row.get('en_name', '')}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>型号</div><div class='preview-value'>{preview_row.get('model', '')}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>基础价格</div><div class='preview-value'>{preview_row.get('price', 0)} {preview_row.get('currency', 'USD')}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>库存</div><div class='preview-value'>{preview_row.get('stock', 0)}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='preview-label'>包装体积</div><div class='preview-value'>{preview_row.get('packing_volume', 0)} CBM/件</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            preview_qty = st.number_input(
                "加入数量",
                min_value=1,
                value=max(int(st.session_state.cart.get(safe_text(preview_row.get("sap")), {}).get("quantity", 1)), 1),
                step=1,
                key="preview_qty",
            )
            if st.button("加入当前确认产品", use_container_width=True):
                item = preview_row.to_dict()
                sap_key = safe_text(item.get("sap"))
                item["quantity"] = int(preview_qty)
                st.session_state.cart[sap_key] = item
                st.success(f"已加入：{sap_key}")

# -----------------------------------------------------------------------------
# Cart
# -----------------------------------------------------------------------------
st.subheader("3. 报价单购物车")
cart = st.session_state.cart
if not cart:
    st.warning("报价单为空。请先从上方筛选结果中勾选产品并加入报价单。")
    st.stop()

cart_df = pd.DataFrame(list(cart.values()))
cart_view_cols = ["sap", "category", "cn_name", "en_name", "model", "price", "quantity", "stock", "packing_volume", "unit", "currency", "package_info", "image_url"]
for c in cart_view_cols:
    if c not in cart_df.columns:
        cart_df[c] = ""

with st.form("cart_update_form"):
    for _, row in cart_df.iterrows():
        sap_key = safe_text(row.get("sap"))
        c0, c1, c2, c3, c4, c5 = st.columns([0.8, 2.6, 1.2, 0.9, 0.9, 0.7], vertical_alignment="center")
        with c0:
            image_url = safe_text(row.get("image_url"))
            if image_url:
                st.image(image_url, width=64)
            else:
                st.caption("无图")
        with c1:
            st.markdown(f"**{safe_text(row.get('cn_name')) or safe_text(row.get('en_name')) or sap_key}**")
            st.caption(f"SAP: {sap_key}｜型号: {safe_text(row.get('model')) or '-'}｜分类: {safe_text(row.get('category')) or '-'}")
        with c2:
            st.write(f"基础价：{float(row.get('price') or 0):.2f} {safe_text(row.get('currency')) or 'USD'}")
            st.caption(f"库存：{int(row.get('stock') or 0)}")
        with c3:
            st.number_input("数量", min_value=1, value=max(int(row.get("quantity") or 1), 1), step=1, key=f"cart_qty_{sap_key}", label_visibility="collapsed")
        with c4:
            st.caption(f"CBM/件：{float(row.get('packing_volume') or 0):.4f}")
        with c5:
            st.checkbox("删除", value=False, key=f"cart_remove_{sap_key}")
    cart_submitted = st.form_submit_button("更新购物车数量/删除", use_container_width=True)

if cart_submitted:
    for sap_key in list(st.session_state.cart.keys()):
        if bool(st.session_state.get(f"cart_remove_{sap_key}", False)):
            st.session_state.cart.pop(sap_key, None)
        elif sap_key in st.session_state.cart:
            st.session_state.cart[sap_key]["quantity"] = max(int(st.session_state.get(f"cart_qty_{sap_key}", 1) or 1), 1)
    st.success("报价单已更新。")
    st.rerun()

b1, b2, b3 = st.columns([1, 1, 4])
if b1.button("清空报价单", use_container_width=True):
    st.session_state.cart = {}
    st.rerun()

# -----------------------------------------------------------------------------
# Quote rules and export
# -----------------------------------------------------------------------------
st.subheader("4. 报价规则与导出")
with st.container(border=True):
    c1, c2, c3, c4 = st.columns([1, 1, 1.2, 1.2])
    mode = c1.selectbox("价格模式", ["原价", "加点", "打折"])
    percent = c2.number_input("比例 %", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
    customer = c3.text_input("客户名称", value="")
    quote_no = c4.text_input("报价单号", value=st.session_state["quote_no_default"], key="quote_no_input")

rule = PriceRule(mode=mode, percent=percent)
items_raw = pd.DataFrame(list(st.session_state.cart.values()))
items = build_quote_items(items_raw, rule)
summary = quote_summary(items)
price_note = f"{mode} {percent:.2f}%" if mode != "原价" else "原价 / Original Price"

m1, m2, m3, m4 = st.columns(4)
m1.metric("SKU数量", summary["sku_count"])
m2.metric("总数量", summary["total_qty"])
m3.metric("总金额(USD)", f"{summary['total_amount']:.2f}")
m4.metric("总体积 CBM", f"{summary['total_volume']:.4f}")

st.subheader("5. 报价预览")
st.caption("英文版 Excel/PDF 的产品名称列会优先引用英文品名 en_name；如果该字段为空，才会回退到中文品名。")
preview_cols = ["sap", "category", "cn_name", "en_name", "model", "base_price", "quote_price", "quantity", "amount", "stock", "packing_volume", "total_volume"]
for col in preview_cols:
    if col not in items.columns:
        items[col] = ""
st.dataframe(items[preview_cols], use_container_width=True, hide_index=True)

current_export_signature = make_export_signature(items, customer, quote_no, price_note)
if st.session_state.get("export_signature") != current_export_signature:
    st.session_state["export_signature"] = current_export_signature
    st.session_state["export_files"] = {}

if st.button("保存报价历史", use_container_width=True):
    save_quote_history(quote_no, customer, summary["total_amount"], created_by=user["username"])
    st.success("已保存报价历史。")

st.markdown("#### 6. 生成下载文件")
st.caption("系统不会在每次页面刷新时自动生成 PDF/Excel。需要哪个文件，就点击对应按钮生成。")
export_files = st.session_state.setdefault("export_files", {})

g1, g2, g3, g4 = st.columns(4)
with g1:
    if st.button("生成中文 Excel", use_container_width=True):
        with st.spinner("正在生成中文 Excel..."):
            export_files["excel_zh"] = export_quote_excel(items, customer, quote_no, price_note, lang="zh")
with g2:
    if st.button("生成英文 Excel", use_container_width=True):
        with st.spinner("正在生成英文 Excel..."):
            export_files["excel_en"] = export_quote_excel(items, customer, quote_no, price_note, lang="en")
with g3:
    if st.button("生成中文 PDF", use_container_width=True):
        with st.spinner("正在生成中文 PDF..."):
            export_files["pdf_zh"] = export_quote_pdf(items, customer, quote_no, price_note, lang="zh")
with g4:
    if st.button("生成英文 PDF", use_container_width=True):
        with st.spinner("正在生成英文 PDF..."):
            export_files["pdf_en"] = export_quote_pdf(items, customer, quote_no, price_note, lang="en")

st.session_state["export_files"] = export_files

d1, d2, d3, d4 = st.columns(4)
with d1:
    if "excel_zh" in export_files:
        st.download_button(
            "下载中文 Excel",
            data=export_files["excel_zh"],
            file_name=f"{quote_no}_中文报价单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with d2:
    if "excel_en" in export_files:
        st.download_button(
            "下载英文 Excel",
            data=export_files["excel_en"],
            file_name=f"{quote_no}_English_Quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with d3:
    if "pdf_zh" in export_files:
        st.download_button(
            "下载中文 PDF",
            data=export_files["pdf_zh"],
            file_name=f"{quote_no}_中文报价单.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
with d4:
    if "pdf_en" in export_files:
        st.download_button(
            "下载英文 PDF",
            data=export_files["pdf_en"],
            file_name=f"{quote_no}_English_Quotation.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
