# Power BI Dashboard — E-Commerce Business Performance Overview
### Complete Step-by-Step Guide | Deneb Vega Edition

---

## 📋 Table of Contents
1. [Prerequisites](#prerequisites)
2. [Data Overview & Missing Values](#data-overview)
3. [Step 1 — Import Data into Power BI](#step-1)
4. [Step 2 — M Code: Main Fact Table (Power Query)](#step-2)
5. [Step 3 — M Code: Date Master Table (Power Query)](#step-3)
6. [Step 4 — Data Model & Relationships](#step-4)
7. [Step 5 — All DAX Measures](#step-5)
8. [Step 6 — Dashboard Canvas Setup](#step-6)
9. [Step 7 — KPI Cards (Deneb Vega)](#step-7)
10. [Step 8 — Nett Revenue Trend Line Chart (Deneb Vega)](#step-8)
11. [Step 9 — Order Status Donut Chart (Deneb Vega)](#step-9)
12. [Step 10 — Monthly Revenue Bar Chart (Deneb Vega)](#step-10)
13. [Step 11 — Top 9 Category Horizontal Bar (Deneb Vega)](#step-11)
14. [Step 12 — Slicers & Filters](#step-12)
15. [Step 13 — Key Insight Text Box](#step-13)
16. [Final Checklist & Tips](#final-tips)

---

## 1. Prerequisites <a name="prerequisites"></a>

Install these before starting:
- **Power BI Desktop** (latest version, June 2024+)
- **Deneb Custom Visual** → AppSource → search "Deneb" → Add
- **Dataset**: `DA_Recruitment_Test_-_Dataset_for_Analysis___Visualization.xlsx`

> ⚠️ Deneb must be installed from AppSource BEFORE importing the file. Vega version used: **v5**.

---

## 2. Data Overview & Missing Values <a name="data-overview"></a>

**Sheet**: `e-commerce dataset`  
**Rows**: 267,408 (order detail level — one row per product in each order)  
**Date Range**: July 2016 → August 2018

### Column Summary
| Column | Type | Notes |
|---|---|---|
| order_id | Text | Order identifier (use DISTINCTCOUNT for Total Orders) |
| order_detail_id | Text | Line-item identifier |
| order_date | Date | Transaction date |
| year_order / month_order | Int | Derived year/month |
| customer_id | Text | **5 NULL rows** → fill with "UNKNOWN" |
| customer_name | Text | **5 NULL rows** → fill with "Unknown Customer" |
| customer_joined_my | Text | **5 NULL rows** → fill with "Unknown" |
| status_customer | Text | **5 NULL rows** → fill with "inactive" |
| order_status_name | Text | **14 NULL rows** → fill with "unknown" |
| order_status_id | Text | **14 NULL rows** → fill with "OS-UNKNOWN" |
| provinsi | Text | 5 unique provinces |
| category_name | Text | 9 categories (house, women, beauty, etc.) |
| gross_revenue / nett_revenue | Decimal | Revenue metrics |
| gross_profit_margin / nett_profit_margin | Decimal | Profit metrics |
| quantity / unit_price / discount | Decimal | Transaction details |

### Status Group Mapping (for Order Status Breakdown)
| Source Status Values | Mapped Group |
|---|---|
| complete, closed | Completed |
| canceled, fraud | Canceled |
| received, cod | Received |
| processing, holded, payment_review, paid, pending, pending_paypal | Process |
| order_refunded, refund, exchange, unknown | Other |

---

## 3. Step 1 — Import Data into Power BI <a name="step-1"></a>

1. Open Power BI Desktop
2. **Home → Get Data → Excel Workbook**
3. Browse to your `.xlsx` file → click **Open**
4. In Navigator, check `e-commerce dataset` → click **Transform Data**
5. Power Query Editor opens → proceed to Step 2

---

## 4. Step 2 — M Code: Main Fact Table <a name="step-2"></a>

In Power Query Editor, click the query `e-commerce dataset` in the left panel.  
Go to **Home → Advanced Editor** → **Delete ALL existing code** → paste the code below:

```m
let
    // ── 1. Load Source ──────────────────────────────────────────────────
    Source = Excel.Workbook(
        File.Contents(
            // UPDATE THIS PATH to your actual file location
            "C:\YourFolder\DA_Recruitment_Test_-_Dataset_for_Analysis___Visualization.xlsx"
        ),
        null,
        true
    ),
    RawSheet    = Source{[Item="e-commerce dataset", Kind="Sheet"]}[Data],
    PromHeaders = Table.PromoteHeaders(RawSheet, [PromoteAllScalars=true]),

    // ── 2. Set Data Types ────────────────────────────────────────────────
    TypedTable = Table.TransformColumnTypes(PromHeaders, {
        {"order_id",             type text},
        {"order_detail_id",      type text},
        {"order_date",           type date},
        {"year_order",           Int64.Type},
        {"month_order",          Int64.Type},
        {"my",                   Int64.Type},
        {"customer_id",          type text},
        {"customer_name",        type text},
        {"customer_joined_my",   type text},
        {"status_customer",      type text},
        {"payment_id",           type text},
        {"payment_method",       type text},
        {"order_status_id",      type text},
        {"order_status_name",    type text},
        {"location_id",          type text},
        {"id_provinsi",          Int64.Type},
        {"provinsi",             type text},
        {"category_id",          type text},
        {"category_name",        type text},
        {"margin_ratio",         type number},
        {"status_category",      type text},
        {"product_id",           type text},
        {"product_name",         type text},
        {"unit_price",           type number},
        {"status_product",       type text},
        {"quantity",             Int64.Type},
        {"discount",             type number},
        {"gross_revenue",        type number},
        {"nett_revenue",         type number},
        {"gross_profit_margin",  type number},
        {"nett_profit_margin",   type number}
    }),

    // ── 3. Handle 5 Missing Customer Rows ───────────────────────────────
    // These 5 rows have null customer_id, customer_name,
    // customer_joined_my, and status_customer
    Fix_CustomerID = Table.TransformColumns(
        TypedTable,
        {{"customer_id", each if _ = null then "UNKNOWN" else _, type text}}
    ),
    Fix_CustomerName = Table.TransformColumns(
        Fix_CustomerID,
        {{"customer_name", each if _ = null then "Unknown Customer" else _, type text}}
    ),
    Fix_CustomerJoined = Table.TransformColumns(
        Fix_CustomerName,
        {{"customer_joined_my", each if _ = null then "Unknown" else _, type text}}
    ),
    Fix_StatusCustomer = Table.TransformColumns(
        Fix_CustomerJoined,
        {{"status_customer", each if _ = null then "inactive" else _, type text}}
    ),

    // ── 4. Handle 14 Missing Order Status Rows ──────────────────────────
    Fix_OrderStatusID = Table.TransformColumns(
        Fix_StatusCustomer,
        {{"order_status_id", each if _ = null then "OS-UNKNOWN" else _, type text}}
    ),
    Fix_OrderStatusName = Table.TransformColumns(
        Fix_OrderStatusID,
        {{"order_status_name", each if _ = null then "unknown" else _, type text}}
    ),

    // ── 5. Add Status Group Column (for Order Status Breakdown visual) ───
    AddStatusGroup = Table.AddColumn(
        Fix_OrderStatusName,
        "status_group",
        each
            if   List.Contains({"complete","closed"}, [order_status_name])
            then "Completed"
            else if List.Contains({"canceled","fraud"}, [order_status_name])
            then "Canceled"
            else if List.Contains({"received","cod"}, [order_status_name])
            then "Received"
            else if List.Contains({"processing","holded","payment_review","paid","pending","pending_paypal"}, [order_status_name])
            then "Process"
            else "Other",
        type text
    ),

    // ── 6. Add Month-Year Label for Charts ──────────────────────────────
    AddMonthYearLabel = Table.AddColumn(
        AddStatusGroup,
        "month_year_label",
        each
            Text.Start(
                Date.ToText([order_date], "MMM"),
                3
            ) & "-" &
            Text.End(Text.From(Date.Year([order_date])), 2),
        type text
    ),

    // ── 7. Add Month Sort Key (YYYYMM) ──────────────────────────────────
    AddSortKey = Table.AddColumn(
        AddMonthYearLabel,
        "month_sort_key",
        each [year_order] * 100 + [month_order],
        Int64.Type
    ),

    // ── 8. Final Sort ────────────────────────────────────────────────────
    SortedResult = Table.Sort(AddSortKey, {{"order_date", Order.Ascending}})

in
    SortedResult
```

> **Rename this query** to `FactOrders` in the query properties panel (right-click → Rename).

---

## 5. Step 3 — M Code: Date Master Table <a name="step-3"></a>

In Power Query Editor: **Home → New Source → Blank Query** → Advanced Editor → paste:

```m
let
    // ── Covers full 2-year dataset range: Jan 2016 → Dec 2018 ───────────
    StartDate = #date(2016, 1, 1),
    EndDate   = #date(2018, 12, 31),

    DayCount  = Duration.Days(EndDate - StartDate) + 1,
    DateList  = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),

    // ── Build base table ─────────────────────────────────────────────────
    BaseTable    = Table.FromList(DateList, Splitter.SplitByNothing()),
    RenameDate   = Table.RenameColumns(BaseTable, {{"Column1", "Date"}}),
    TypeDate     = Table.TransformColumnTypes(RenameDate, {{"Date", type date}}),

    // ── Add derived date columns ─────────────────────────────────────────
    AddYear       = Table.AddColumn(TypeDate,   "Year",          each Date.Year([Date]),                       Int64.Type),
    AddMonthNum   = Table.AddColumn(AddYear,    "MonthNum",      each Date.Month([Date]),                      Int64.Type),
    AddMonthName  = Table.AddColumn(AddMonthNum,"MonthName",     each Date.ToText([Date], "MMMM"),             type text),
    AddMonthShort = Table.AddColumn(AddMonthName,"MonthShort",   each Date.ToText([Date], "MMM"),              type text),
    AddMonthLabel = Table.AddColumn(AddMonthShort,"MonthYearLabel",
                        each Date.ToText([Date], "MMM") & "-" & Text.End(Text.From(Date.Year([Date])), 2),
                        type text),
    AddSortKey    = Table.AddColumn(AddMonthLabel, "MonthSortKey",
                        each Date.Year([Date]) * 100 + Date.Month([Date]),
                        Int64.Type),
    AddQuarter    = Table.AddColumn(AddSortKey, "Quarter",
                        each "Q" & Text.From(Date.QuarterOfYear([Date])),
                        type text),
    AddQtrYear    = Table.AddColumn(AddQuarter, "QuarterYear",
                        each "Q" & Text.From(Date.QuarterOfYear([Date])) & " " & Text.From(Date.Year([Date])),
                        type text),
    AddWeekday    = Table.AddColumn(AddQtrYear, "Weekday",
                        each Date.DayOfWeekName([Date]),
                        type text),
    AddIsWeekend  = Table.AddColumn(AddWeekday, "IsWeekend",
                        each Date.DayOfWeek([Date]) >= 5,
                        type logical),

    // ── Final sort ───────────────────────────────────────────────────────
    Sorted = Table.Sort(AddIsWeekend, {{"Date", Order.Ascending}})
in
    Sorted
```

> **Rename this query** to `DimDate`. Click **Close & Apply**.

---

## 6. Step 4 — Data Model & Relationships <a name="step-4"></a>

Go to **Model View** (left sidebar icon).

### Create Relationship
| From Table | From Column | To Table | To Column | Cardinality | Direction |
|---|---|---|---|---|---|
| FactOrders | order_date | DimDate | Date | Many-to-One (M:1) | Single |

1. Drag `order_date` from **FactOrders** onto `Date` in **DimDate**
2. Confirm: Cardinality = **Many to one**, Cross filter = **Single**
3. Click **OK**

> This relationship enables all time-intelligence DAX functions to work correctly.

---

## 7. Step 5 — All DAX Measures <a name="step-5"></a>

In **Report View**, go to **Home → New Measure**. Create a dedicated measure table:
- Right-click in Fields panel → **New Table** → type: `_Measures = {1}` → rename it to `_Measures`

Create all measures inside `_Measures` table by right-clicking it → **New Measure**.

---

### 📌 SECTION A — Core Measures

```dax
Total Orders =
DISTINCTCOUNT( FactOrders[order_id] )
```

```dax
Total Customers =
DISTINCTCOUNT( FactOrders[customer_id] )
```

```dax
Quantity Sold =
SUM( FactOrders[quantity] )
```

```dax
Gross Revenue =
SUM( FactOrders[gross_revenue] )
```

```dax
Nett Revenue =
SUM( FactOrders[nett_revenue] )
```

```dax
Gross Profit =
SUM( FactOrders[gross_profit_margin] )
```

```dax
Nett Profit =
SUM( FactOrders[nett_profit_margin] )
```

```dax
Nett Profit Margin % =
DIVIDE( [Nett Profit], [Gross Revenue], 0 )
```

```dax
Avg Order Value =
DIVIDE( [Nett Revenue], [Total Orders], 0 )
```

```dax
Cancellation Rate =
DIVIDE(
    CALCULATE(
        DISTINCTCOUNT( FactOrders[order_id] ),
        FactOrders[status_group] = "Canceled"
    ),
    [Total Orders],
    0
)
```

---

### 📌 SECTION B — Previous Month Measures (Time Intelligence)

```dax
PM Total Orders =
CALCULATE(
    [Total Orders],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Total Customers =
CALCULATE(
    [Total Customers],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Quantity Sold =
CALCULATE(
    [Quantity Sold],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Gross Revenue =
CALCULATE(
    [Gross Revenue],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Nett Revenue =
CALCULATE(
    [Nett Revenue],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Gross Profit =
CALCULATE(
    [Gross Profit],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Nett Profit =
CALCULATE(
    [Nett Profit],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Nett Profit Margin % =
CALCULATE(
    [Nett Profit Margin %],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Avg Order Value =
CALCULATE(
    [Avg Order Value],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

```dax
PM Cancellation Rate =
CALCULATE(
    [Cancellation Rate],
    DATEADD( DimDate[Date], -1, MONTH )
)
```

---

### 📌 SECTION C — MoM % Change Measures

```dax
Orders MoM % =
DIVIDE( [Total Orders] - [PM Total Orders], [PM Total Orders], 0 )
```

```dax
Customers MoM % =
DIVIDE( [Total Customers] - [PM Total Customers], [PM Total Customers], 0 )
```

```dax
Qty MoM % =
DIVIDE( [Quantity Sold] - [PM Quantity Sold], [PM Quantity Sold], 0 )
```

```dax
Gross Rev MoM % =
DIVIDE( [Gross Revenue] - [PM Gross Revenue], [PM Gross Revenue], 0 )
```

```dax
Nett Rev MoM % =
DIVIDE( [Nett Revenue] - [PM Nett Revenue], [PM Nett Revenue], 0 )
```

```dax
Gross Profit MoM % =
DIVIDE( [Gross Profit] - [PM Gross Profit], [PM Gross Profit], 0 )
```

```dax
Nett Profit MoM % =
DIVIDE( [Nett Profit] - [PM Nett Profit], [PM Nett Profit], 0 )
```

```dax
NPM MoM ppts =
( [Nett Profit Margin %] - [PM Nett Profit Margin %] ) * 100
```

```dax
AOV MoM % =
DIVIDE( [Avg Order Value] - [PM Avg Order Value], [PM Avg Order Value], 0 )
```

```dax
Cancel Rate MoM ppts =
( [Cancellation Rate] - [PM Cancellation Rate] ) * 100
```

---

### 📌 SECTION D — Display/Label Measures (used in slicers/titles)

```dax
Selected Month Label =
IF(
    HASONEVALUE( DimDate[MonthYearLabel] ),
    VALUES( DimDate[MonthYearLabel] ),
    "All Months"
)
```

```dax
Selected Year =
IF(
    HASONEVALUE( DimDate[Year] ),
    FORMAT( VALUES( DimDate[Year] ), "0" ),
    "All Years"
)
```

---

## 8. Step 6 — Dashboard Canvas Setup <a name="step-6"></a>

1. **View → Page Size** → Custom → Width: **1440**, Height: **900**
2. **View → Wallpaper** → Color: `#F0F4FA` (light blue-gray)
3. **Format → Canvas Background** → Color: `#FFFFFF`, Transparency: 0%

### Layout Zones (recreate as rectangles)
Insert a rectangle for each zone via **Insert → Shapes → Rectangle**:

| Zone | Position | Color | Notes |
|---|---|---|---|
| Header | Top full width | `#1B2A4A` | Title area |
| KPI Row 1 | Below header, left 5 cards | White | Cards row 1 |
| KPI Row 2 | Below KPI Row 1, 5 cards | White | Cards row 2 |
| Bottom Left | Charts left side | White | Revenue charts |
| Bottom Right | Charts right side | White | Status & category |
| Footer | Very bottom | `#1B2A4A` | Key insight bar |

> **Tip**: Group each zone's shapes by selecting them → right-click → Group.

---

## 9. Step 7 — KPI Cards (Deneb Vega) <a name="step-7"></a>

### Setup for Each KPI Card

1. Insert **Deneb** visual (from Visualizations pane)
2. Add the required measures to the **Values** field well (see each card below)
3. Click **Edit** in the Deneb visual
4. Choose **Vega** (not Vega-Lite) in the schema dropdown
5. Paste the JSON spec below
6. Click **Apply**

> 📐 **Recommended size per KPI card**: Width ~250px, Height ~100px

---

### 🃏 KPI Card — TOTAL ORDER

**Values to add**: `Total Orders`, `Orders MoM %`

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 230,
  "height": 95,
  "padding": 0,
  "background": "transparent",
  "data": [{ "name": "dataset" }],
  "marks": [
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0},
          "y": {"value": 0},
          "width": {"signal": "width"},
          "height": {"signal": "height"},
          "fill": {"value": "#FFFFFF"},
          "cornerRadius": {"value": 8},
          "stroke": {"value": "#DDE3F0"},
          "strokeWidth": {"value": 1.5}
        }
      }
    },
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0},
          "y": {"value": 0},
          "width": {"value": 4},
          "height": {"signal": "height"},
          "fill": {"value": "#2563EB"},
          "cornerRadiusTopLeft": {"value": 8},
          "cornerRadiusBottomLeft": {"value": 8}
        }
      }
    },
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 10},
          "y": {"value": 15},
          "width": {"value": 30},
          "height": {"value": 30},
          "fill": {"value": "#EFF6FF"},
          "cornerRadius": {"value": 15}
        }
      }
    },
    {
      "type": "text",
      "encode": {
        "enter": {
          "text": {"value": "🛒"},
          "x": {"value": 25},
          "y": {"value": 37},
          "align": {"value": "center"},
          "fontSize": {"value": 14}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {"value": "TOTAL ORDER"},
          "x": {"value": 50},
          "y": {"value": 24},
          "fontSize": {"value": 9},
          "fill": {"value": "#94A3B8"},
          "fontWeight": {"value": "700"},
          "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {"signal": "format(datum['Total Orders'], ',')"},
          "x": {"value": 50},
          "y": {"value": 52},
          "fontSize": {"value": 22},
          "fill": {"value": "#1E293B"},
          "fontWeight": {"value": "bold"},
          "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {
            "signal": "(datum['Orders MoM %'] >= 0 ? '▲ ' : '▼ ') + format(abs(datum['Orders MoM %']), '.2%') + '  vs prev mo'"
          },
          "x": {"value": 50},
          "y": {"value": 75},
          "fontSize": {"value": 10},
          "fill": {
            "signal": "datum['Orders MoM %'] >= 0 ? '#16A34A' : '#DC2626'"
          },
          "font": {"value": "Segoe UI"}
        }
      }
    }
  ]
}
```

---

### 🃏 KPI Card — TOTAL CUSTOMER

**Values to add**: `Total Customers`, `Customers MoM %`

Copy the Total Order spec above and change these lines only:

```json
// Line 1 — Change icon:
"text": {"value": "👥"},

// Line 2 — Change label:
"text": {"value": "TOTAL CUSTOMER"},

// Line 3 — Change value signal:
"text": {"signal": "format(datum['Total Customers'], ',')"},

// Line 4 — Change MoM signal:
"text": {
  "signal": "(datum['Customers MoM %'] >= 0 ? '▲ ' : '▼ ') + format(abs(datum['Customers MoM %']), '.2%') + '  vs prev mo'"
},
"fill": {
  "signal": "datum['Customers MoM %'] >= 0 ? '#16A34A' : '#DC2626'"
}
```

---

### 🃏 KPI Card — QUANTITY SOLD

**Values to add**: `Quantity Sold`, `Qty MoM %`

Same template — change to:
- Icon: `📦`
- Label: `QUANTITY SOLD`
- Value: `format(datum['Quantity Sold'], ',')`
- MoM: use `Qty MoM %`

---

### 🃏 KPI Card — GROSS REVENUE (Currency format)

**Values to add**: `Gross Revenue`, `Gross Rev MoM %`

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 230,
  "height": 95,
  "padding": 0,
  "background": "transparent",
  "data": [{ "name": "dataset" }],
  "marks": [
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0}, "y": {"value": 0},
          "width": {"signal": "width"}, "height": {"signal": "height"},
          "fill": {"value": "#FFFFFF"}, "cornerRadius": {"value": 8},
          "stroke": {"value": "#DDE3F0"}, "strokeWidth": {"value": 1.5}
        }
      }
    },
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0}, "y": {"value": 0},
          "width": {"value": 4}, "height": {"signal": "height"},
          "fill": {"value": "#2563EB"},
          "cornerRadiusTopLeft": {"value": 8},
          "cornerRadiusBottomLeft": {"value": 8}
        }
      }
    },
    {
      "type": "text",
      "encode": {
        "enter": {
          "text": {"value": "💰"},
          "x": {"value": 25}, "y": {"value": 37},
          "align": {"value": "center"}, "fontSize": {"value": 14}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {"value": "GROSS REVENUE"},
          "x": {"value": 50}, "y": {"value": 24},
          "fontSize": {"value": 9}, "fill": {"value": "#94A3B8"},
          "fontWeight": {"value": "700"}, "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {
            "signal": "datum['Gross Revenue'] >= 1000000000 ? 'Rp ' + format(datum['Gross Revenue']/1000000000, '.2f') + ' B' : datum['Gross Revenue'] >= 1000000 ? 'Rp ' + format(datum['Gross Revenue']/1000000, '.2f') + ' M' : 'Rp ' + format(datum['Gross Revenue']/1000, '.1f') + ' K'"
          },
          "x": {"value": 50}, "y": {"value": 52},
          "fontSize": {"value": 18}, "fill": {"value": "#1E293B"},
          "fontWeight": {"value": "bold"}, "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {
            "signal": "(datum['Gross Rev MoM %'] >= 0 ? '▲ ' : '▼ ') + format(abs(datum['Gross Rev MoM %']), '.2%') + '  vs prev mo'"
          },
          "x": {"value": 50}, "y": {"value": 75},
          "fontSize": {"value": 10},
          "fill": {"signal": "datum['Gross Rev MoM %'] >= 0 ? '#16A34A' : '#DC2626'"},
          "font": {"value": "Segoe UI"}
        }
      }
    }
  ]
}
```

> **For Nett Revenue, Gross Profit, Nett Profit**: Copy this spec, change label text and field references to `Nett Revenue`/`Nett Rev MoM %`, `Gross Profit`/`Gross Profit MoM %`, `Nett Profit`/`Nett Profit MoM %`.

---

### 🃏 KPI Card — NETT PROFIT MARGIN % (Percentage + ppts format)

**Values to add**: `Nett Profit Margin %`, `NPM MoM ppts`

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 230,
  "height": 95,
  "padding": 0,
  "background": "transparent",
  "data": [{ "name": "dataset" }],
  "marks": [
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0}, "y": {"value": 0},
          "width": {"signal": "width"}, "height": {"signal": "height"},
          "fill": {"value": "#FFFFFF"}, "cornerRadius": {"value": 8},
          "stroke": {"value": "#DDE3F0"}, "strokeWidth": {"value": 1.5}
        }
      }
    },
    {
      "type": "rect",
      "encode": {
        "enter": {
          "x": {"value": 0}, "y": {"value": 0},
          "width": {"value": 4}, "height": {"signal": "height"},
          "fill": {"value": "#7C3AED"},
          "cornerRadiusTopLeft": {"value": 8},
          "cornerRadiusBottomLeft": {"value": 8}
        }
      }
    },
    {
      "type": "text",
      "encode": {
        "enter": {
          "text": {"value": "📊"},
          "x": {"value": 25}, "y": {"value": 37},
          "align": {"value": "center"}, "fontSize": {"value": 14}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {"value": "NETT PROFIT MARGIN"},
          "x": {"value": 50}, "y": {"value": 24},
          "fontSize": {"value": 9}, "fill": {"value": "#94A3B8"},
          "fontWeight": {"value": "700"}, "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {"signal": "format(datum['Nett Profit Margin %'], '.1%')"},
          "x": {"value": 50}, "y": {"value": 52},
          "fontSize": {"value": 24}, "fill": {"value": "#1E293B"},
          "fontWeight": {"value": "bold"}, "font": {"value": "Segoe UI"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "text": {
            "signal": "(datum['NPM MoM ppts'] >= 0 ? '▲ ' : '▼ ') + format(abs(datum['NPM MoM ppts']), '.2f') + ' ppts  vs prev mo'"
          },
          "x": {"value": 50}, "y": {"value": 75},
          "fontSize": {"value": 10},
          "fill": {"signal": "datum['NPM MoM ppts'] >= 0 ? '#16A34A' : '#DC2626'"},
          "font": {"value": "Segoe UI"}
        }
      }
    }
  ]
}
```

> **For Cancellation Rate**: Same spec pattern — use `Cancellation Rate` + `Cancel Rate MoM ppts`, change accent color to `#EF4444` (red), format: `format(datum['Cancellation Rate'], '.2%')`.

---

### 🃏 KPI Card — AVERAGE ORDER VALUE

**Values to add**: `Avg Order Value`, `AOV MoM %`

Same as Gross Revenue card — change:
- Label: `AVG ORDER VALUE`
- Value signal: `'Rp ' + format(datum['Avg Order Value']/1000, '.1f') + ' K'`
- MoM: `AOV MoM %`

---

## 10. Step 8 — Nett Revenue Trend Line Chart <a name="step-8"></a>

### Setup
1. Insert a new **Deneb** visual
2. In **Values** field well, add:
   - `DimDate[MonthYearLabel]` (dimension — this creates one row per month)
   - `DimDate[MonthSortKey]` (for ordering)
   - `Nett Revenue` (measure)
3. Size the visual: ~**600 × 280 px**
4. Edit → Vega → paste spec below

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 560,
  "height": 220,
  "padding": {"top": 30, "left": 55, "right": 20, "bottom": 40},
  "background": "transparent",
  "config": {"font": "Segoe UI"},
  "data": [
    {
      "name": "dataset",
      "transform": [
        {
          "type": "collect",
          "sort": {"field": "MonthSortKey", "order": "ascending"}
        }
      ]
    }
  ],
  "scales": [
    {
      "name": "xscale",
      "type": "band",
      "domain": {"data": "dataset", "field": "MonthYearLabel"},
      "range": "width",
      "padding": 0.3
    },
    {
      "name": "yscale",
      "type": "linear",
      "domain": {"data": "dataset", "field": "Nett Revenue"},
      "range": "height",
      "zero": true,
      "nice": true
    }
  ],
  "axes": [
    {
      "orient": "bottom",
      "scale": "xscale",
      "tickColor": "#E2E8F0",
      "domainColor": "#E2E8F0",
      "labelColor": "#64748B",
      "labelFontSize": 10,
      "labelAngle": 0
    },
    {
      "orient": "left",
      "scale": "yscale",
      "grid": true,
      "gridColor": "#F1F5F9",
      "gridOpacity": 0.8,
      "domainColor": "#E2E8F0",
      "tickColor": "#E2E8F0",
      "labelColor": "#64748B",
      "labelFontSize": 10,
      "format": ".2s",
      "tickCount": 5
    }
  ],
  "marks": [
    {
      "type": "area",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel", "band": 0.5},
          "y": {"scale": "yscale", "field": "Nett Revenue"},
          "y2": {"scale": "yscale", "value": 0},
          "fill": {
            "gradient": "linear",
            "stops": [
              {"offset": 0, "color": "#3B82F6", "opacity": 0.25},
              {"offset": 1, "color": "#3B82F6", "opacity": 0.02}
            ],
            "x1": 0, "y1": 0, "x2": 0, "y2": 1
          }
        }
      }
    },
    {
      "type": "line",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel", "band": 0.5},
          "y": {"scale": "yscale", "field": "Nett Revenue"},
          "stroke": {"value": "#2563EB"},
          "strokeWidth": {"value": 2.5},
          "interpolate": {"value": "monotone"}
        }
      }
    },
    {
      "type": "symbol",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel", "band": 0.5},
          "y": {"scale": "yscale", "field": "Nett Revenue"},
          "size": {"value": 55},
          "fill": {"value": "#2563EB"},
          "stroke": {"value": "#FFFFFF"},
          "strokeWidth": {"value": 2}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel", "band": 0.5},
          "y": {"scale": "yscale", "field": "Nett Revenue", "offset": -14},
          "align": {"value": "center"},
          "baseline": {"value": "bottom"},
          "text": {
            "signal": "datum['Nett Revenue'] >= 1000000000 ? format(datum['Nett Revenue']/1000000000,',.1f') + 'B' : format(datum['Nett Revenue']/1000000,',.1f') + 'M'"
          },
          "fontSize": {"value": 10},
          "fill": {"value": "#374151"},
          "fontWeight": {"value": "600"}
        }
      }
    }
  ],
  "title": {
    "text": "NETT REVENUE TREND",
    "anchor": "start",
    "offset": 10,
    "fontSize": 11,
    "fontWeight": "700",
    "color": "#1E293B",
    "font": "Segoe UI"
  }
}
```

---

## 11. Step 9 — Order Status Donut Chart <a name="step-9"></a>

### Setup
1. Insert a new **Deneb** visual
2. In **Values** field well, add:
   - `FactOrders[status_group]` (dimension)
   - `Total Orders` (measure)
3. Size: ~**400 × 320 px**
4. Edit → Vega → paste:

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 340,
  "height": 280,
  "padding": 5,
  "background": "transparent",
  "config": {"font": "Segoe UI"},
  "data": [
    {
      "name": "dataset",
      "transform": [
        {
          "type": "pie",
          "field": "Total Orders",
          "sort": true
        }
      ]
    },
    {
      "name": "totals",
      "source": "dataset",
      "transform": [
        {
          "type": "aggregate",
          "ops": ["sum"],
          "fields": ["Total Orders"],
          "as": ["grand_total"]
        }
      ]
    }
  ],
  "scales": [
    {
      "name": "colorScale",
      "type": "ordinal",
      "domain": ["Completed", "Received", "Canceled", "Process", "Other"],
      "range": ["#2563EB", "#F59E0B", "#EF4444", "#10B981", "#94A3B8"]
    }
  ],
  "marks": [
    {
      "type": "arc",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"signal": "width / 2 - 60"},
          "y": {"signal": "height / 2"},
          "startAngle": {"field": "startAngle"},
          "endAngle": {"field": "endAngle"},
          "innerRadius": {"value": 68},
          "outerRadius": {"value": 110},
          "fill": {"scale": "colorScale", "field": "status_group"},
          "stroke": {"value": "#FFFFFF"},
          "strokeWidth": {"value": 2.5}
        }
      }
    },
    {
      "type": "text",
      "encode": {
        "enter": {
          "x": {"signal": "width / 2 - 60"},
          "y": {"signal": "height / 2 - 10"},
          "align": {"value": "center"},
          "baseline": {"value": "middle"},
          "text": {"value": "Total Order"},
          "fontSize": {"value": 10},
          "fill": {"value": "#64748B"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "totals"},
      "encode": {
        "enter": {
          "x": {"signal": "width / 2 - 60"},
          "y": {"signal": "height / 2 + 12"},
          "align": {"value": "center"},
          "baseline": {"value": "middle"},
          "text": {"signal": "format(datum.grand_total, ',')"},
          "fontSize": {"value": 18},
          "fontWeight": {"value": "bold"},
          "fill": {"value": "#1E293B"}
        }
      }
    },
    {
      "type": "symbol",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"signal": "width - 85"},
          "y": {"signal": "30 + indexof(['Completed','Received','Canceled','Process','Other'], datum['status_group']) * 28"},
          "size": {"value": 80},
          "shape": {"value": "circle"},
          "fill": {"scale": "colorScale", "field": "status_group"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"signal": "width - 70"},
          "y": {"signal": "30 + indexof(['Completed','Received','Canceled','Process','Other'], datum['status_group']) * 28 + 5"},
          "text": {
            "signal": "datum['status_group'] + '  ' + format(datum['Total Orders'], ',') + ' (' + format(datum['Total Orders']/datum['grand_total_for_pct'], '.1%') + ')'"
          },
          "align": {"value": "left"},
          "fontSize": {"value": 10},
          "fill": {"value": "#374151"}
        }
      }
    }
  ],
  "title": {
    "text": "ORDER STATUS BREAKDOWN",
    "anchor": "start",
    "offset": 8,
    "fontSize": 11,
    "fontWeight": "700",
    "color": "#1E293B"
  }
}
```

> ⚠️ **Note on legend percentages**: The `grand_total_for_pct` field won't exist by default. Add this to the `dataset` transform:
> Add a `"lookup"` transform joining `totals` → `grand_total` field and alias it, OR simplify by calculating pct directly in the pie transform using a formula mark.

**Simplified legend approach** — replace the last text mark with:

```json
{
  "type": "text",
  "from": {"data": "dataset"},
  "encode": {
    "enter": {
      "x": {"signal": "width - 70"},
      "y": {
        "signal": "34 + indexof(['Completed','Received','Canceled','Process','Other'], datum['status_group']) * 28"
      },
      "text": {
        "signal": "datum['status_group'] + '  ' + format(datum['Total Orders'], ',') + ' (' + format(datum['endAngle'] - datum['startAngle'], '.1%').replace('0.', '').replace('%','') + '%)'"
      },
      "align": {"value": "left"},
      "fontSize": {"value": 10},
      "fill": {"value": "#374151"}
    }
  }
}
```

> The arc angle span `(endAngle - startAngle) / (2 * PI)` equals the percentage. Use this signal: `format((datum['endAngle'] - datum['startAngle']) / (2 * 3.14159265), '.1%')`.

---

## 12. Step 10 — Monthly Revenue Bar Chart <a name="step-10"></a>

### Setup
1. Insert a new **Deneb** visual
2. **Values**: `DimDate[MonthYearLabel]`, `DimDate[MonthSortKey]`, `Nett Revenue`
3. Size: ~**600 × 280 px**

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 550,
  "height": 200,
  "padding": {"top": 30, "left": 50, "right": 20, "bottom": 40},
  "background": "transparent",
  "config": {"font": "Segoe UI"},
  "data": [
    {
      "name": "dataset",
      "transform": [
        {
          "type": "collect",
          "sort": {"field": "MonthSortKey", "order": "ascending"}
        },
        {
          "type": "window",
          "ops": ["rank"],
          "as": ["rank"]
        }
      ]
    }
  ],
  "scales": [
    {
      "name": "xscale",
      "type": "band",
      "domain": {"data": "dataset", "field": "MonthYearLabel"},
      "range": "width",
      "padding": 0.25
    },
    {
      "name": "yscale",
      "type": "linear",
      "domain": {"data": "dataset", "field": "Nett Revenue"},
      "range": "height",
      "zero": true,
      "nice": true
    }
  ],
  "axes": [
    {
      "orient": "bottom",
      "scale": "xscale",
      "labelColor": "#64748B",
      "labelFontSize": 10,
      "domainColor": "#E2E8F0",
      "tickColor": "#E2E8F0"
    },
    {
      "orient": "left",
      "scale": "yscale",
      "grid": true,
      "gridColor": "#F1F5F9",
      "labelColor": "#64748B",
      "labelFontSize": 10,
      "format": ".2s",
      "tickCount": 5,
      "domainColor": "#E2E8F0",
      "tickColor": "#E2E8F0"
    }
  ],
  "marks": [
    {
      "type": "rect",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel"},
          "width": {"scale": "xscale", "band": 1},
          "y": {"scale": "yscale", "field": "Nett Revenue"},
          "y2": {"scale": "yscale", "value": 0},
          "cornerRadiusTopLeft": {"value": 4},
          "cornerRadiusTopRight": {"value": 4},
          "fill": {
            "signal": "datum.rank === length(data('dataset')) ? '#1D4ED8' : '#93C5FD'"
          }
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "x": {"scale": "xscale", "field": "MonthYearLabel", "band": 0.5},
          "y": {"scale": "yscale", "field": "Nett Revenue", "offset": -8},
          "align": {"value": "center"},
          "text": {
            "signal": "datum['Nett Revenue'] >= 1000000000 ? format(datum['Nett Revenue']/1000000000,',.1f') + 'B' : format(datum['Nett Revenue']/1000000,',.1f') + 'M'"
          },
          "fontSize": {"value": 9},
          "fill": {"value": "#374151"},
          "fontWeight": {"value": "600"}
        }
      }
    }
  ],
  "title": {
    "text": "MONTHLY NETT REVENUE COMPARISON",
    "anchor": "start",
    "offset": 10,
    "fontSize": 11,
    "fontWeight": "700",
    "color": "#1E293B"
  }
}
```

> The last month bar is highlighted in dark blue (`#1D4ED8`) using a `window rank` transform — the bar with `rank === length(dataset)` is the latest month.

---

## 13. Step 11 — Top Category Horizontal Bar Chart <a name="step-11"></a>

### Setup
1. Insert a new **Deneb** visual
2. **Values**: `FactOrders[category_name]`, `Nett Revenue`
3. Size: ~**500 × 320 px**

```json
{
  "$schema": "https://vega.github.io/schema/vega/v5.json",
  "width": 420,
  "height": 260,
  "padding": {"top": 30, "left": 100, "right": 80, "bottom": 20},
  "background": "transparent",
  "config": {"font": "Segoe UI"},
  "data": [
    {
      "name": "dataset",
      "transform": [
        {
          "type": "collect",
          "sort": {"field": "Nett Revenue", "order": "descending"}
        }
      ]
    }
  ],
  "scales": [
    {
      "name": "yscale",
      "type": "band",
      "domain": {"data": "dataset", "field": "category_name"},
      "range": "height",
      "padding": 0.3
    },
    {
      "name": "xscale",
      "type": "linear",
      "domain": {"data": "dataset", "field": "Nett Revenue"},
      "range": "width",
      "zero": true,
      "nice": true
    }
  ],
  "axes": [
    {
      "orient": "left",
      "scale": "yscale",
      "labelColor": "#374151",
      "labelFontSize": 11,
      "domainColor": "#E2E8F0",
      "tickColor": "#E2E8F0",
      "labelAlign": "right"
    },
    {
      "orient": "bottom",
      "scale": "xscale",
      "grid": true,
      "gridColor": "#F1F5F9",
      "labelColor": "#64748B",
      "labelFontSize": 9,
      "format": ".2s",
      "tickCount": 5,
      "domainColor": "#E2E8F0",
      "tickColor": "#E2E8F0"
    }
  ],
  "marks": [
    {
      "type": "rect",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "y": {"scale": "yscale", "field": "category_name"},
          "height": {"scale": "yscale", "band": 1},
          "x": {"scale": "xscale", "value": 0},
          "x2": {"scale": "xscale", "field": "Nett Revenue"},
          "cornerRadiusTopRight": {"value": 4},
          "cornerRadiusBottomRight": {"value": 4},
          "fill": {"value": "#3B82F6"}
        }
      }
    },
    {
      "type": "text",
      "from": {"data": "dataset"},
      "encode": {
        "enter": {
          "y": {"scale": "yscale", "field": "category_name", "band": 0.5},
          "x": {"scale": "xscale", "field": "Nett Revenue", "offset": 6},
          "align": {"value": "left"},
          "baseline": {"value": "middle"},
          "text": {
            "signal": "datum['Nett Revenue'] >= 1000000000 ? 'Rp ' + format(datum['Nett Revenue']/1000000000,',.2f') + ' B' : 'Rp ' + format(datum['Nett Revenue']/1000000,',.2f') + ' M'"
          },
          "fontSize": {"value": 10},
          "fill": {"value": "#374151"},
          "fontWeight": {"value": "600"}
        }
      }
    }
  ],
  "title": {
    "text": "TOP CATEGORY BY NETT REVENUE",
    "anchor": "start",
    "offset": 10,
    "fontSize": 11,
    "fontWeight": "700",
    "color": "#1E293B"
  }
}
```

---

## 14. Step 12 — Slicers & Filters <a name="step-12"></a>

Add standard Power BI slicers for the top filter bar (matching the reference image dropdowns):

| Slicer | Field | Style | Description |
|---|---|---|---|
| Month | `DimDate[MonthYearLabel]` | Dropdown | Sort by MonthSortKey |
| Category | `FactOrders[category_name]` | Dropdown | Multi-select |
| Province | `FactOrders[provinsi]` | Dropdown | 5 Java provinces |
| Payment Method | `FactOrders[payment_method]` | Dropdown | 18 methods |
| Order Status | `FactOrders[status_group]` | Dropdown | 5 groups |

### Slicer Formatting
For each slicer:
1. Format Pane → **Slicer settings** → Style: **Dropdown**
2. **Values** → Font: Segoe UI, Size: 10
3. **Header** → Font: Segoe UI Bold, Size: 9, Color: `#64748B`
4. **Background**: Transparent
5. **Border**: Color `#DDE3F0`, Rounded corners: 4px

> **Month slicer sort**: Go to `DimDate` table → click `MonthYearLabel` column → Column Tools → **Sort by Column** → select `MonthSortKey`. This ensures Jan-Feb-Mar... order.

---

## 15. Step 13 — Key Insight Text Box <a name="step-13"></a>

At the bottom of the dashboard, recreate the dark insight bar:

1. **Insert → Text Box**
2. Paste or type your insight text (update dynamically via a DAX card if needed)
3. Background: `#1B2A4A`, Text color: White, Font: Segoe UI 10pt
4. Add a lightbulb icon (💡) using **Insert → Icons** or emoji

**Example DAX measure for dynamic insight**:
```dax
Key Insight Text =
"Nett Revenue in " & [Selected Month Label] &
" is Rp " & FORMAT([Nett Revenue]/1000000000, "#,##0.00") & "B" &
". Top category: " &
CALCULATE(
    FIRSTNONBLANK(FactOrders[category_name], 1),
    TOPN(1,
        SUMMARIZE(FactOrders, FactOrders[category_name], "_rev", [Nett Revenue]),
        [_rev], DESC
    )
) &
". Cancellation Rate: " & FORMAT([Cancellation Rate], "0.0%") & "."
```

> Add this measure to a **Card** visual inside the dark footer box, formatted to match.

---

## 16. Final Checklist & Tips <a name="final-tips"></a>

### ✅ Pre-Publication Checklist
- [ ] All 5 null `customer_id` rows handled (filled with "UNKNOWN")
- [ ] All 14 null `order_status_name` rows handled (filled with "unknown")
- [ ] `status_group` column created in M code and visible in data
- [ ] `DimDate` table linked to `FactOrders` via `order_date` → `Date`
- [ ] All 10 DAX KPI measures + 10 PM measures + 10 MoM measures created
- [ ] `MonthYearLabel` sorted by `MonthSortKey` in `DimDate`
- [ ] All 5 slicers added and formatted
- [ ] All Deneb visuals use **Vega v5** (not Vega-Lite)
- [ ] Each Deneb visual has correct field names matching DAX measure names exactly

### 🔧 Common Issues & Fixes
| Issue | Fix |
|---|---|
| Vega spec shows blank | Check that measure names in `datum['...']` match exactly what you named them in DAX (case-sensitive) |
| Months out of order in charts | Ensure `MonthSortKey` (YYYYMM integer) is added to Deneb Values and used in `collect` transform |
| Donut percentages wrong | Use `(datum.endAngle - datum.startAngle) / (2 * PI)` for pct calculation |
| Time intelligence shows blank | Verify `DimDate` → `FactOrders` relationship is active and uses `Mark as Date Table` |
| Bar chart last bar not highlighted | Make sure `window rank` is in the transform and data is sorted ascending first |

### 📐 Recommended Visual Sizes
| Visual | Width | Height |
|---|---|---|
| KPI Card (×10) | 250px | 100px |
| Revenue Trend Line | 620px | 280px |
| Order Status Donut | 400px | 320px |
| Monthly Bar Chart | 620px | 280px |
| Top Category H-Bar | 500px | 320px |
| Slicers (×5) | 160px | 40px |

### 🎨 Color Palette Reference
| Usage | Hex |
|---|---|
| Primary Blue (bars, lines) | `#2563EB` |
| Light Blue (non-highlighted bars) | `#93C5FD` |
| Card Accent — Revenue | `#2563EB` |
| Card Accent — Margin/Profit | `#7C3AED` |
| Card Accent — Cancellation | `#EF4444` |
| Positive MoM | `#16A34A` |
| Negative MoM | `#DC2626` |
| Header/Footer Background | `#1B2A4A` |
| Canvas Background | `#F0F4FA` |
| Card Background | `#FFFFFF` |
| Card Border | `#DDE3F0` |

---

*End of Tutorial — E-Commerce XYZ Power BI Dashboard*
*Dataset: 267,408 rows | Date range: Jul 2016 – Aug 2018 | 9 categories | 5 provinces | 18 payment methods*
