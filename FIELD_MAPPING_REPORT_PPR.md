# PPR 给水管报价表字段结构调整报告

## 读取到的 Excel 字段

文件：`250710 PPR给水管报价.xlsx`，Sheet1，范围约 A1:S495。

表头为：

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

## 核心变化

原产品表是一行一个 SAP。这个 PPR 报价表是一行两个 SAP：

- `Grey SAP No.`：灰色产品 SAP
- `Green SAP No.`：绿色产品 SAP

因此代码已经改成：导入时自动把一行拆成两条产品记录，形成标准的一行一个 SAP。若灰色 SAP 和绿色 SAP 相同，则只保留一条，颜色标记为 `Universal`。

## 新数据库字段结构

```text
sap
category
cn_name
en_name
model
description
color
size_mm
weight
material_description
price
price_cny
stock
packing_volume
qty_per_ctn
package_length
package_width
package_height
unit
currency
package_info
image_url
active
```

## PPR Excel → 数据库字段映射

| PPR Excel 字段 | 数据库字段 | 说明 |
|---|---|---|
| Grey SAP No. | sap | 灰色产品拆成一条记录 |
| Green SAP No. | sap | 绿色产品拆成一条记录 |
| Description | category / en_name / description | 英文产品类别和描述 |
| 产品名称 | cn_name | 中文品名 |
| Size(mm) | size_mm / model | 规格型号 |
| Grey / Green | color | 自动生成颜色字段 |
| Weight (kg/m or pc) | weight | 重量 |
| 物料描述（灰色） | material_description | 灰色物料描述 |
| 物料描述（绿色） | material_description | 绿色物料描述 |
| 灰色基准价格（USD）/m(pcs) | price | 灰色 USD 价格 |
| 绿色基准价格（USD）/m(pcs) | price | 绿色 USD 价格 |
| 灰色基准价格（CNY）/m(pcs) | price_cny | 灰色人民币价格 |
| 绿色基准价格（CNY）/m(pcs) | price_cny | 绿色人民币价格 |
| Pcs/Carton | qty_per_ctn | 每箱数量 |
| L | package_length | 外箱长 |
| W | package_width | 外箱宽 |
| H | package_height | 外箱高 |
| CBM | packing_volume | 单箱体积 / 报价体积字段 |
| 灰色图片 / Picture | image_url | 如有文件名或 URL 会自动处理 |

## 同时保持兼容

旧的水暖卫浴产品库仍可导入，字段如：

```text
sap, category, cn_name, model, description, price, stock, packing_volume, Qty/CTN, unit, currency, package_info, image_url, active
```

缺失的新字段会自动补空或补 0，不影响旧报价单流程。

## 导出字段

Excel/PDF 报价单仍按用户要求保持简洁字段：

```text
SAP号、分类、品名型号、图片、描述、单位、单价、数量、金额、Qty/CTN、CBM
```

其中：

- `品名型号` 会自动合并品名、规格、颜色。
- `描述` 会尽量包含英文描述、中文品名、规格、颜色、重量、包装信息。
- `CBM` 使用 `total_volume`，即 `packing_volume × 数量`。

## 部署注意

1. 覆盖 GitHub 中的代码文件。
2. Streamlit Cloud Reboot。
3. 进入后台重新上传 `250710 PPR给水管报价.xlsx`。
4. 若 Supabase 旧表缺字段，程序会自动 `ALTER TABLE` 补字段。
5. 如果手动删除过 `products` 表，本版也会自动重新创建。
