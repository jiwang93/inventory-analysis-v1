# Inventory Insight (库存透视分析)

> Enterprise Warehouse Inventory Analytics, Multi-Dimensional Auditing & Intelligent Relocation System.  
> 企业级仓储库位库存多维分析、混载混档异常诊断与智能调拨纠偏系统。

---

## 📖 Overview (项目概述)

**Inventory Insight** is a high-performance, client-side warehouse inventory analytics platform designed for manufacturing and distribution facilities. It automates inventory data parsing from Excel workbooks, enforces strict reconciliation rules, identifies mixed-power and mixed-customer storage anomalies, isolates non-A grade items, and generates actionable warehouse relocation plans with professional styled Excel exports.

**Inventory Insight** 是一款面向制造业与仓储物流场景的高性能库位库存分析平台。支持直接上传原始库存 Excel 表格，自动执行严格的数据闭环核验（条码排除、品级隔离）、混档与混客户检测、智能调拨纠偏建议，并支持一键导出格式美观、带专业配色的 Excel 报告。

---

## ✨ Key Features (核心功能)

1. **Automated Excel Parsing & Predefined Location Layout (智能解析与标准库位映射)**
   - Pre-configured with standard warehouse racks (`A01` to `A40`, `B01` to `B32`), capacity limits (30 boxes standard, 24 boxes for B31-B32), and dynamic detection of unlisted locations.
   - 预置标准仓位拓扑结构，自动核算各库位满载率与可用容量。

2. **Strict Audit & Closed-Loop Reconciliation (严格数据闭环与勾稽诊断)**
   - **Column S (Barcode Count) Rule**: Rows with Column S = 0 or empty are automatically filtered out (scrap / non-scanned items).
   - **Column K (Quality Grade) Rule**: Strictly separates Grade-A products from Non-A (Grade B, F, R, etc.) inventory.
   - **8-Box Discrepancy Diagnostics**: Interactive audit modal reconciling `Raw Rows (7,380) - Barcode=0 (23) = Valid Boxes (7,357)` against distinct counts and duplicate box numbers.

3. **Order-to-Customer Mapping (工单-客户关系对照表)**
   - Upload custom mapping tables (Col B: Order/Sales No. -> Col C: Customer Name) to enrich raw inventory data.
   - Provides a one-click template download (`Order_Customer_Template.xlsx`).

4. **Anomaly Alerts & Risk Detection (仓储异常与混载预警)**
   - **Mixed Power Alert (功率混档预警)**: Highlights bins containing multiple power classes.
   - **Mixed Customer Alert (客户混载预警)**: Identifies bins storing items for different customers simultaneously.

5. **Intelligent Relocation & Consolidation (智能调拨与移库纠偏)**
   - Calculates dominant customer/power for each location.
   - Generates step-by-step relocation advice (target bin recommendation, box lists, and consolidation logic).

6. **Professional Excel Exports (高质量精简 Excel 导出)**
   - Designed with primary English headers and concise Chinese subheadings (`English (中文)`).
   - Styled using `ExcelJS` with corporate header colors, borders, center alignments, and auto-width columns.
   - Exports include:
     - `Location_Analysis_YYYY-MM-DD.xlsx` (Location Summary)
     - `Mixed_Power_Alerts_YYYY-MM-DD.xlsx` (Mixed Power Anomalies)
     - `Mixed_Customer_Alerts_YYYY-MM-DD.xlsx` (Mixed Customer Anomalies)
     - `Customer_Inventory_Stats_YYYY-MM-DD.xlsx` (Customer Grouped Inventory)
     - `Relocation_Suggestions_YYYY-MM-DD.xlsx` (Actionable Relocation Plan)
     - `Non_A_Grade_Report_YYYY-MM-DD.xlsx` (Non-A Grade Summary & Details)
     - `Zero_Barcode_Filtered_YYYY-MM-DD.xlsx` (Filtered Zero-Barcode Rows)
     - `Audit_Diagnostics_Report_YYYY-MM-DD.xlsx` (Deep Reconciliation Audit)

---

## 📊 Data Specifications (数据规范)

| Column (列) | Name (字段) | Requirement & Processing Rule (规则与处理逻辑) |
|---|---|---|
| **Col C** | 库位名称 (Location) | Rack location name (e.g. `A01-01-01`, `B12-02-01`). |
| **Col E** | 功率档 (Power Class) | Power rating (e.g. `450W`, `550W`). |
| **Col G** | 工单号 (Work Order) | Work order / batch identifier. Matches against mapping table. |
| **Col I** | 箱号 (Box No.) | Unique box identifier. Used for duplicate scanning check. |
| **Col K** | 箱等级 (Grade) | Quality rating. Starts with `A` is Grade-A; others are routed to Non-A pool. |
| **Col S** | 箱条码数 (Barcode Qty) | Must be > 0. If `0` or empty, automatically excluded from inventory stats. |

---

## 🛠 Tech Stack (技术架构)

- **Framework**: React 19 (TypeScript) + Vite
- **Styling**: Tailwind CSS
- **Excel Processing**: ExcelJS, XLSX (SheetJS), File-Saver
- **Animations & Icons**: Framer Motion, Lucide React

---

## 🚀 Getting Started (快速上手)

### Prerequisites (环境要求)
- Node.js 18.x or later
- npm or yarn or pnpm

### Installation (安装依赖)
```bash
# Clone the repository
git clone https://github.com/your-username/inventory-insight.git

# Enter project directory
cd inventory-insight

# Install dependencies
npm install
```

### Development (开发运行)
```bash
# Start local development server
npm run dev
```
Open your browser at `http://localhost:3000`.

### Production Build (编译打包)
```bash
# Type check and build
npm run build

# Preview build locally
npm run preview
```

---

## 📄 License (开源许可)
MIT License.
