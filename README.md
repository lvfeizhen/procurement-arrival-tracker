# 采购到位率查询系统

上传 PR 报表、PurchaseOrder 报表和采购承诺导出数据，自动匹配并计算采购到位率，一键导出分析结果。

## 快速开始

### 方式一：有 Python 环境

```bash
pip install -r requirements.txt
python app.py
```

浏览器打开 `http://127.0.0.1:5000`

### 方式二：无 Python 环境（Windows）

下载 [Releases](../../releases) 中的 `采购到位率查询系统.exe`，双击运行即可。

> 如需自行打包：`pip install pyinstaller && pyinstaller --onefile --name "采购到位率查询系统" --add-data "templates;templates" app.py`

## 使用方法

1. 打开系统页面
2. 上传 3 个 Excel 文件：
   - **PR 报表**（文件名含 `PR`）
   - **PurchaseOrder 报表**（文件名含 `PurchaseOrder`）
   - **采购承诺导出数据**
3. 点击「开始分析」
4. 查看汇总数据，下载 Excel 报表

## 数据匹配逻辑

```
第一步：PurchaseOrder.来源单据号 ∈ PR.单据编号  → 筛选

第二步：采购承诺 (U9采购单号 + U9采购单行号)
         ↔
        PurchaseOrder (单据编号 + 行号)
         → 关联匹配

第三步：计算到位状态
  - 累计实收 == 0           → 未到位
  - 0 < 累计实收 < 采购数量  → 部分未到位
  - 累计实收 >= 采购数量     → 已到位
```

## 输出报表（5 个 Sheet）

| Sheet | 说明 |
|-------|------|
| 到位率明细 | 全部匹配数据，红/黄/绿标记到位状态 |
| 汇总统计 | 按未到位/部分未到位/已到位汇总 |
| 匹配失败数据 | 无法匹配到 PO 的承诺记录 |
| 到期履约情况 | 承诺收货日期 ≤ 今天的全部记录 |
| 到期履约汇总 | 到期数据按状态汇总 |

## 技术栈

- Python 3 + Flask
- Pandas（数据处理）
- OpenPyXL（Excel 格式化）
- PyInstaller（打包 exe）
