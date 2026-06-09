# LESSO 外贸报价系统 - PPR 字段结构版 v4.0

本版在原报价系统基础上，增加对 `250710 PPR给水管报价.xlsx` 的直接导入支持。

## 支持的工作流

```text
后台上传 Excel → 自动写入 Supabase products 表 → 前台搜索产品 → 加入购物车 → 导出 Excel / PDF
```

## 1. 支持两类产品表

### A. 原水暖卫浴标准表

```text
sap, category, cn_name, model, description, price, stock, packing_volume, Qty/CTN, unit, currency, package_info, image_url, active
```

### B. PPR 给水管宽表

```text
Description
产品名称
灰色图片
Picture
Size(mm)
Grey SAP No.
Green SAP No.
Weight (kg/m or pc)
物料描述（灰色）
物料描述（绿色）
灰色基准价格（USD）/m(pcs)
灰色基准价格（CNY）/m(pcs)
绿色基准价格（USD）/m(pcs)
绿色基准价格（CNY）/m(pcs)
Pcs/Carton
L
W
H
CBM
```

PPR 表会自动拆分：

- `Grey SAP No.` → 一条灰色产品
- `Green SAP No.` → 一条绿色产品

如果两列 SAP 相同，则只保留一条，颜色为 `Universal`。

## 2. Supabase products 表结构

```sql
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
);
```

程序会自动创建或迁移这些字段。

## 3. 导出字段

Excel/PDF 导出字段固定为：

```text
SAP号 / 分类 / 品名型号 / 图片 / 描述 / 单位 / 单价 / 数量 / 金额 / Qty/CTN / CBM
```

图片会嵌入 Excel 的“图片”单元格中，并尽量占满单元格。

## 4. 图片路径

当前默认 Supabase Storage 路径为 bucket 根目录：

```text
https://oevlzhvdgojgzacnbfka.supabase.co/storage/v1/object/public/product-images/8060040825.png
```

如需改回 `products/xxx.png`，在 Streamlit Secrets 中设置：

```toml
SUPABASE_STORAGE_PREFIX = "products"
```

## 5. Streamlit Secrets

```toml
DATABASE_URL = "postgresql+psycopg2://..."
DEFAULT_ADMIN_USER = "karl"
DEFAULT_ADMIN_PASSWORD = "your-password"
SUPABASE_URL = "https://oevlzhvdgojgzacnbfka.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "your-service-role-key"
SUPABASE_STORAGE_BUCKET = "product-images"
SUPABASE_STORAGE_PREFIX = ""
```

## 6. 部署

1. 上传覆盖 GitHub。
2. Streamlit Cloud 里 Reboot app。
3. 后台上传新的 PPR 报价 Excel。
4. 前台搜索 SAP、中文品名、英文类别、规格、颜色即可。
