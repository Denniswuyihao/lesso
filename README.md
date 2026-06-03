# 外贸报价表自动生成系统 Final Stable

面向外贸业务的云端报价系统：产品数据库、云端图片、报价购物车、中英文 Excel/PDF 导出、管理员权限。

## 本版重点

- 使用 Supabase PostgreSQL 作为云端产品数据库。
- 使用 Supabase Storage 作为唯一稳定图片来源，字段为 `products.image_url`。
- 不再依赖 GitHub 或 Streamlit 本地目录保存产品图片。
- 产品筛选采用表单提交，避免每次输入、勾选、改数量都立即重跑。
- 购物车数量修改也采用表单提交，减少卡顿。
- PDF / Excel 仅在点击生成按钮时生成，不在页面刷新时自动生成。
- PDF 字体和排版重新优化，顶部 LESSO、顶部描述、底部说明避免中英文混排导致字距异常。
- Excel / PDF 金额列统一标注 USD。

## 必须配置的 Streamlit Secrets

```toml
DATABASE_URL = "postgresql+psycopg2://postgres.xxxxx:你的数据库密码@aws-1-ap-northeast-2.pooler.supabase.com:6543/postgres?sslmode=require"

DEFAULT_ADMIN_USER = "admin"
DEFAULT_ADMIN_PASSWORD = "你的后台密码"

SUPABASE_URL = "https://oevlzhvdgojgzacnbfka.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "你的 service_role key"
SUPABASE_STORAGE_BUCKET = "product-images"
SUPABASE_STORAGE_PREFIX = "products"
```

`SUPABASE_SERVICE_ROLE_KEY` 只能放在 Streamlit Secrets，不能上传 GitHub。

## Storage 图片路径规则

推荐统一：

```text
bucket: product-images
folder: products
filename: SAP号.png
```

最终 URL：

```text
https://oevlzhvdgojgzacnbfka.supabase.co/storage/v1/object/public/product-images/products/SAP号.png
```

## 产品表字段

推荐 Excel / CSV 字段：

```text
sap, category, cn_name, en_name, model, description, price, stock, packing_volume, unit, currency, package_info, image_url, active
```

也兼容中文列名，例如：SAP号、分类、中文品名、英文品名、型号、描述、价格、库存、包装体积、包装、图片URL、状态。

## 运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署

上传到 GitHub 后，在 Streamlit Cloud 中选择：

```text
Repository: Tamagotchi02/foreign-trade-quote-app
Branch: main
Main file path: app.py
```

保存 Secrets 后 Reboot app。

## 速度建议

- 默认每次筛选显示 20 条，建议不要一次显示 100 条以上带图产品。
- 产品图建议压缩到 100KB-500KB。
- 平时关闭左侧“显示后台管理”。
- PDF/Excel 只在最终确认后点击生成。
