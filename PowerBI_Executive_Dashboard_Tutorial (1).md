# Executive E-Commerce Dashboard in Power BI — Full Beginner Build Guide

A step-by-step tutorial that recreates the reference "Executive Dashboard" layout using your real dataset (`DA_Recruitment_Test_-_Dataset_for_Analysis___Visualization.xlsx`). Every step is explained in order so you can follow it without errors. We never rename a single source column, preprocessing is done **entirely in Power Query M**, and we handle the missing rows.

> **Reading order matters.** Do the steps top to bottom. Slicers are built before visuals (as requested), measures before charts, and the data model before measures.

---

## 0. What we are building (and how the reference maps to your data)

The reference image is a 2018 retail template. Your data is an Indonesian e-commerce order-detail table from **Jul 2016 → Aug 2018**, so the "Aug-18 vs Jul-18" comparison in the reference lines up perfectly with the last two months of your data.

| Reference element | Your data equivalent | Note |
|---|---|---|
| Period = Aug-18, Compare = Jul-18 | `my` = 201808 vs 201807 | `my` is already a YYYYMM integer |
| Net Revenue / Net Profit | `nett_revenue` / `nett_profit_margin` | `*_profit_margin` columns are **Rupiah amounts**, not ratios |
| Profit Margin % | profit ÷ revenue | computed measure |
| Total Orders | distinct `order_id` | 181,247 unique orders |
| Active Customers | distinct `customer_id` | 71,966 unique customers |
| Revenue by Category | `category_name` (9 categories) | women, house, beauty, men, jewelry, kids, shoes, accessories, bags |
| Revenue by **Channel** | Revenue by **Payment Method** (`payment_method`) | no channel column exists; payment method is the closest dimension and also drives the COD/cancellation insight |
| New vs Returning customers | derived from `customer_joined_my` vs `my` | a customer is "New" in the month they joined, else "Returning" |
| Cancellation rate | `order_status_name` = `canceled` | canceled is ~33.6% of detail rows |
| "West Java & Jakarta top contributors" | `provinsi` = Jawa Barat (40%) + DKI Jakarta (21%) | matches the reference takeaway |

**Design decisions you should know up front**

- **Realized revenue.** A third of rows are `canceled`. Headline KPIs (Net Revenue, Net Profit, AOV) count only *realized* statuses: `complete`, `received`, `paid`, `closed`, `cod`. Cancellation and refund get their own metrics. This is more honest than summing every row and is what a recruiter wants to see.
- **Period vs Compare.** We use two small **disconnected** period tables so the KPI cards can show "selected month vs comparison month" while the trend line still shows the full timeline — exactly like the reference.
- **Star-ish model.** One fact table (`Sales`) + one `Date` table + two disconnected period tables. Light, fast, and robust.

---

## 1. Dataset analysis (done — read this first)

**Grain:** one row per `order_detail_id` (line item). 267,408 rows.

**Column inventory (all kept exactly as-is):**

- **Keys / text:** `order_id`, `order_detail_id`, `customer_id`, `customer_name`, `payment_id`, `order_status_id`, `location_id`, `id_provinsi`, `category_id`, `product_id`
- **Date / time:** `order_date` (real date), `year_order`, `month_order`, `my` (YYYYMM int), `customer_joined_my` (text like `2018-4`)
- **Dimensions:** `payment_method`, `order_status_name`, `provinsi`, `category_name`, `product_name`, `status_customer`, `status_category`, `status_product`
- **Numbers / metrics:** `margin_ratio`, `unit_price`, `quantity`, `discount`, `gross_revenue`, `nett_revenue`, `gross_profit_margin`, `nett_profit_margin`

**Datatype fixes needed (only these):**

- `order_date` → **Date** (stored as Excel serial / datetime; force Date).
- `year_order`, `month_order`, `my`, `id_provinsi`, `quantity` → **Whole Number**.
- `margin_ratio`, `unit_price`, `discount`, `gross_revenue`, `nett_revenue`, `gross_profit_margin`, `nett_profit_margin` → **Decimal Number**.
- All IDs and text → **Text** (keep `customer_joined_my` as text; we parse it into a helper later).

**Missing values (the only data quality issue):**

| Column | Null rows | Fix |
|---|---|---|
| `order_status_id`, `order_status_name` | 14 | fill with `"unknown"` |
| `customer_id`, `customer_name`, `customer_joined_my`, `status_customer` | 5 | fill with placeholders (`CST-UNKNOWN`, `Unknown Customer`, etc.) |

We **fill** rather than delete (19 rows out of 267k) so revenue totals are not silently reduced. The single-line alternative to delete them instead is given in Step 2.

**Single-value columns** (`status_customer`, `status_category`, `status_product` are all `active`) — not useful as slicers; ignore them.

---

## 2. Preprocessing — full Power Query M (handles missing rows)

### 2.1 Load the file

1. Open Power BI Desktop → **Home → Get data → Excel workbook** → choose the file → in the Navigator tick the sheet **`e-commerce dataset`** → **Transform Data** (do NOT click Load yet; we want the editor).
2. In Power Query Editor, the query will be named after the sheet. Rename the **query** to `Sales` (renaming a *query/table* is allowed — only *column* names are protected).
3. **Home → Advanced Editor**, delete everything, and paste the script below.

> Renaming the query to `Sales` is the only rename we do. Every column name is preserved.

### 2.2 The M script (paste into Advanced Editor)

```m
let
    // 1. SOURCE — point at the uploaded workbook.
    //    Replace the path with your own file location.
    Source = Excel.Workbook(
        File.Contents("C:\Users\YOU\Downloads\DA_Recruitment_Test_-_Dataset_for_Analysis___Visualization.xlsx"),
        null, true
    ),

    // 2. NAVIGATE to the sheet (Kind = "Sheet" so it works even if no named table exists).
    SheetData = Source{[Item="e-commerce dataset", Kind="Sheet"]}[Data],

    // 3. PROMOTE the first row to headers.
    Promoted = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),

    // 4. SET DATA TYPES (no renaming — names stay identical to source).
    Typed = Table.TransformColumnTypes(Promoted, {
        {"order_id", type text},
        {"order_detail_id", type text},
        {"order_date", type date},
        {"year_order", Int64.Type},
        {"month_order", Int64.Type},
        {"my", Int64.Type},
        {"customer_id", type text},
        {"customer_name", type text},
        {"customer_joined_my", type text},
        {"status_customer", type text},
        {"payment_id", type text},
        {"payment_method", type text},
        {"order_status_id", type text},
        {"order_status_name", type text},
        {"location_id", type text},
        {"id_provinsi", Int64.Type},
        {"provinsi", type text},
        {"category_id", type text},
        {"category_name", type text},
        {"margin_ratio", type number},
        {"status_category", type text},
        {"product_id", type text},
        {"product_name", type text},
        {"unit_price", type number},
        {"status_product", type text},
        {"quantity", Int64.Type},
        {"discount", type number},
        {"gross_revenue", type number},
        {"nett_revenue", type number},
        {"gross_profit_margin", type number},
        {"nett_profit_margin", type number}
    }),

    // 5. DROP any fully-blank rows that Excel sometimes leaves behind.
    NoBlankRows = Table.SelectRows(Typed, each not List.IsEmpty(
        List.RemoveMatchingItems(Record.FieldValues(_), {"", null})
    )),

    // 6. HANDLE MISSING ROWS — fill nulls with safe placeholders
    //    (14 rows miss order status, 5 rows miss customer info).
    FillStatus = Table.ReplaceValue(NoBlankRows, null, "unknown",
        Replacer.ReplaceValue, {"order_status_id", "order_status_name"}),

    FillCustomer = Table.ReplaceValue(FillStatus, null, "CST-UNKNOWN",
        Replacer.ReplaceValue, {"customer_id"}),
    FillCustName = Table.ReplaceValue(FillCustomer, null, "Unknown Customer",
        Replacer.ReplaceValue, {"customer_name"}),
    FillCustStat = Table.ReplaceValue(FillCustName, null, "active",
        Replacer.ReplaceValue, {"status_customer"}),

    // 7. ADD HELPER COLUMNS (additive only — nothing is renamed/removed).
    //    join_my: turn "2018-4" into 201804 so we can compare to `my`.
    AddJoinMY = Table.AddColumn(FillCustStat, "join_my", each
        let parts = try Text.Split([customer_joined_my], "-") otherwise null
        in  if parts = null or List.Count(parts) < 2 then null
            else try Number.FromText(parts{0}) * 100 + Number.FromText(parts{1})
                 otherwise null, Int64.Type),

    //    Customer Type: New if they ordered in their join month, else Returning.
    AddCustType = Table.AddColumn(AddJoinMY, "Customer Type", each
        if [join_my] = null then "Unknown"
        else if [my] = [join_my] then "New"
        else "Returning", type text)
in
    AddCustType
```

### 2.3 What each step does (plain English)

- **Steps 1–3** open the workbook, grab the sheet, and use row 1 as headers.
- **Step 4** fixes datatypes. `order_date` becomes a real Date — this is what lets Power BI do month/year time-intelligence.
- **Step 5** removes any stray empty rows (defensive; your file had none, but this makes the query reusable).
- **Step 6** is the missing-row handling: status nulls → `"unknown"`, customer nulls → readable placeholders. Revenue is untouched.
- **Step 7** adds two *new* columns (`join_my`, `Customer Type`) used by the New-vs-Returning visual. These are additions, not edits to source columns.

**Want to delete the missing-status rows instead of filling?** Replace Step 6's `FillStatus` line with:
```m
FillStatus = Table.SelectRows(NoBlankRows, each [order_status_name] <> null),
```

Click **Home → Close & Apply**.

---

## 3. Data model (do this before measures)

You now have one table, `Sales`. Add three small tables.

### 3.1 Date table (enables the trend axis and time logic)

**Modeling → New table:**
```dax
Date =
VAR MinD = MIN ( Sales[order_date] )
VAR MaxD = MAX ( Sales[order_date] )
RETURN
ADDCOLUMNS (
    CALENDAR ( DATE ( YEAR(MinD), MONTH(MinD), 1 ), MaxD ),
    "Year", YEAR ( [Date] ),
    "Month No", MONTH ( [Date] ),
    "Month-Year", FORMAT ( [Date], "MMM-yy" ),
    "MY Key", YEAR ( [Date] ) * 100 + MONTH ( [Date] )
)
```
- **Mark as date table:** select the `Date` table → **Table tools → Mark as date table** → Date field = `Date`.
- **Relationship:** drag `Date[Date]` onto `Sales[order_date]` (1-to-many, single direction).
- **Sort the label:** select `Date[Month-Year]` → **Column tools → Sort by column → MY Key**.

### 3.2 Period table (the "Period" slicer source — disconnected)

**Modeling → New table:**
```dax
Period =
ADDCOLUMNS (
    DISTINCT ( SELECTCOLUMNS ( Sales, "PeriodKey", Sales[my] ) ),
    "Period Label", FORMAT ( DATE ( INT([PeriodKey]/100), MOD([PeriodKey],100), 1 ), "MMM-yy" )
)
```
Sort `Period[Period Label]` by `PeriodKey`. **Do NOT** relate this table to anything — it stays disconnected.

### 3.3 Compare table (the "Compare With" slicer source — disconnected)

```dax
Compare =
ADDCOLUMNS (
    DISTINCT ( SELECTCOLUMNS ( Sales, "ComparePeriodKey", Sales[my] ) ),
    "Compare Label", FORMAT ( DATE ( INT([ComparePeriodKey]/100), MOD([ComparePeriodKey],100), 1 ), "MMM-yy" )
)
```
Sort `Compare[Compare Label]` by `ComparePeriodKey`. Also disconnected.

> **Why two disconnected tables?** Because the slicers must drive *measures* (selected vs comparison month) without filtering the whole report. A normal connected slicer would hide every month except one and break the trend chart.

Your model now looks like: `Date → Sales` (related), `Period` and `Compare` floating free.

---

## 4. Slicers FIRST (build these before any chart)

We place three slicers in a top row, matching the reference proportions. Build all three now, then format, then move to visuals.

### 4.1 Period slicer

1. **Insert a Slicer** visual.
2. Field = `Period[Period Label]`.
3. Format pane → **Slicer settings → Style = Dropdown**, **Selection → Single select = On**.
4. Default selection: open the dropdown and pick **Aug-18** (your latest month).

### 4.2 Compare With slicer

1. Insert another **Slicer**.
2. Field = `Compare[Compare Label]`.
3. **Style = Dropdown**, **Single select = On**.
4. Default selection = **Jul-18**.

### 4.3 Category slicer

1. Insert another **Slicer**.
2. Field = `Sales[category_name]`.
3. **Style = Dropdown**. Leave multi-select on so "Select all" behaves like the reference "All".
4. This slicer **does** filter the visuals (it is on the fact table), exactly like the reference Category dropdown.

### 4.4 Slicer formatting (executive look)

Apply to all three:

- **Visual → Slicer header → Off** (we'll add our own text labels "Period / Compare With / Category" as Text boxes above each, like the reference).
- **General → Effects → Background = White**, **Visual border = On**, **Rounded corners = 8 px**.
- **Values → Font = Segoe UI, 10–11 px, dark grey (#3B3B3B)**.
- **Shadow = On**, subtle, to lift the cards off the page.
- Size each slicer ~ **180 × 40 px**, lay them in one horizontal row in the top-right, mirroring the reference header.

### 4.5 Interactions / sync

- **Period** and **Compare** are disconnected, so they will *not* visually filter charts — that is intended. They only feed measures.
- **Category** filters everything by default. If you want to *exclude* the trend chart from category filtering (to keep a full-category trend), select the Category slicer → **Format → Edit interactions** → set the trend chart to **None**.
- If you build multiple report pages, use **View → Sync slicers** to keep Period/Compare/Category consistent across pages.

---

## 5. Measures & dynamic Rupiah formatting

Create a measures-only table to keep things tidy: **Home → Enter data**, one empty table named `_Measures`, then add measures to it.

### 5.1 Base measures (realized = complete/received/paid/closed/cod)

```dax
Realized Revenue =
CALCULATE (
    SUM ( Sales[nett_revenue] ),
    Sales[order_status_name] IN { "complete", "received", "paid", "closed", "cod" }
)

Realized Profit =
CALCULATE (
    SUM ( Sales[nett_profit_margin] ),
    Sales[order_status_name] IN { "complete", "received", "paid", "closed", "cod" }
)

Total Orders = DISTINCTCOUNT ( Sales[order_id] )

Active Customers = DISTINCTCOUNT ( Sales[customer_id] )

Profit Margin % = DIVIDE ( [Realized Profit], [Realized Revenue] )
```

### 5.2 Period selectors (read the disconnected slicers)

```dax
Sel Period = SELECTEDVALUE ( Period[PeriodKey] )
Sel Compare = SELECTEDVALUE ( Compare[ComparePeriodKey] )
```

### 5.3 Period-aware KPI measures (selected month vs comparison month)

```dax
NR Period =
VAR p = [Sel Period]
RETURN CALCULATE ( [Realized Revenue], KEEPFILTERS ( Sales[my] = p ) )

NR Compare =
VAR c = [Sel Compare]
RETURN CALCULATE ( [Realized Revenue], KEEPFILTERS ( Sales[my] = c ) )

NR Delta % = DIVIDE ( [NR Period] - [NR Compare], [NR Compare] )
```

Clone the `NR Period / NR Compare / NR Delta %` pattern for the other KPIs by swapping the base measure:

- **Profit:** `NP Period`, `NP Compare`, `NP Delta %` → use `[Realized Profit]`.
- **Orders:** `Orders Period`, `Orders Compare`, `Orders Delta %` → use `[Total Orders]`.
- **Customers:** `Cust Period`, `Cust Compare`, `Cust Delta %` → use `[Active Customers]`.
- **Margin:** `Margin Period`, `Margin Compare`, `Margin Delta` → use `[Profit Margin %]` (for margin use a simple subtraction, not a ratio: `[Margin Period] - [Margin Compare]`).

### 5.4 Dynamic Rupiah format (Rp Jt / Rp M)

This is the reusable pattern. One formatted **text** measure per KPI value you want displayed:

```dax
NR Period (Label) =
VAR v = [NR Period]
VAR a = ABS ( v )
RETURN
SWITCH (
    TRUE (),
    a >= 1000000000, "Rp " & FORMAT ( v / 1000000000, "0.00" ) & " M",   // Miliar
    a >= 1000000,    "Rp " & FORMAT ( v / 1000000, "0" ) & " Jt",        // Juta
    "Rp " & FORMAT ( v, "#,0" )
)
```
Results look like **Rp 4.21 M**, **Rp 620 Jt**, **Rp 850,000** — exactly the spec. Clone it as `NP Period (Label)`, `NR Compare (Label)`, etc., swapping the inner measure.

**Delta as a clean text + arrow** (for KPI sublabels like "▲ 8.4% vs Jul-18"):
```dax
NR Delta (Label) =
VAR d = [NR Delta %]
VAR cmpName = SELECTEDVALUE ( Compare[Compare Label] )
RETURN
    IF ( d >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( d ), "0.0%" ) & " vs " & cmpName
```
And a direction flag for conditional colour:
```dax
NR Delta Dir = IF ( [NR Delta %] >= 0, "up", "down" )
```

> **Tip:** For native KPI cards you don't strictly need the text measures — you can set the field's display unit. But the `(Label)` measures give you pixel-perfect "Rp 4.21 M" control and are required if you render cards in Deneb.

### 5.5 Sales-summary measures (bottom-right table)

```dax
AOV = DIVIDE ( [Realized Revenue], [Total Orders] )
Orders per Customer = DIVIDE ( [Total Orders], [Active Customers] )
Revenue per Customer = DIVIDE ( [Realized Revenue], [Active Customers] )

Cancellation Rate =
DIVIDE (
    CALCULATE ( DISTINCTCOUNT ( Sales[order_id] ), Sales[order_status_name] = "canceled" ),
    [Total Orders]
)

Refund Rate =
DIVIDE (
    CALCULATE ( DISTINCTCOUNT ( Sales[order_id] ),
        Sales[order_status_name] IN { "refund", "order_refunded" } ),
    [Total Orders]
)

Repeat Purchase Rate =
VAR multi =
    COUNTROWS ( FILTER ( VALUES ( Sales[customer_id] ),
        CALCULATE ( DISTINCTCOUNT ( Sales[order_id] ) ) > 1 ) )
RETURN DIVIDE ( multi, [Active Customers] )
```

### 5.6 New vs Returning revenue (customer overview)

These need no new measure — just put `Sales[Customer Type]` on a donut/bar with `[Realized Revenue]`. The helper column from Step 2 does the work.

---

## 6. Build the visuals (layout matches the reference)

Work top to bottom. Set the page to **16:9**, light grey canvas (`#F5F6FA`).

### Row A — Header
- **Text box:** `EXECUTIVE DASHBOARD` (bold, ~28 px) + subtitle `Business Performance Analysis | Aug 2018`.
- Place the three slicers from Step 4 on the right.

### Row B — Five KPI cards
Use **native Card (new) visual** for reliability (Deneb version in Step 7):
1. **Net Revenue** → field `[NR Period (Label)]`; reference label `[NR Delta (Label)]`.
2. **Total Orders** → `[Orders Period]` (display as whole number); sublabel `[Orders Delta (Label)]`.
3. **Net Profit** → `[NP Period (Label)]`; sublabel `[NP Delta (Label)]`.
4. **Active Customers** → `[Cust Period]`; sublabel `[Cust Delta (Label)]`.
5. **Profit Margin** → `[Margin Period]` (format %); sublabel `[Margin Delta (Label)]`.

Card formatting: white background, 8 px rounded corners, soft shadow, accent icon colour per card (blue/orange/green/purple/teal like the reference). Conditional-format the delta text green/red using the `*_Delta Dir` measures (**Format → Callout value → fx → Field value**).

### Row C — three visuals
1. **Revenue, Profit & Orders Trend** — **Line chart**. X = `Date[Month-Year]`; Y = `[Realized Revenue]`, `[Realized Profit]`; secondary line `[Total Orders]`. (Set the Category slicer interaction to *None* here if you want the full-category trend.)
2. **Revenue by Category** — **Clustered bar (horizontal)**. Y = `Sales[category_name]`, X = `[Realized Revenue]`, sorted descending. Add `Sales[margin_ratio]` (avg) and `[Profit Margin %]` as extra columns if you prefer the table-style look in the reference. (Top categories will be **women, house, beauty**.)
3. **Revenue by Payment Method** — **Donut**. Legend = `Sales[payment_method]`, Values = `[Realized Revenue]`. (This replaces "Revenue by Channel". Limit to top 4–5 + "Other" via the visual's filter for a clean donut.)

### Row D — three visuals
1. **Top 5 Products by Revenue** — **Bar (horizontal)**. Y = `Sales[product_name]`, X = `[Realized Revenue]`, **Top N filter = 5** by `[Realized Revenue]`.
2. **Customer Overview** — **Donut** with `Sales[Customer Type]` + `[Realized Revenue]` (New vs Returning, ~46% / 54%). Optionally a small clustered column beside it for revenue contribution.
3. **Sales Summary (Aug-18)** — **Matrix/Table**. Rows = a small disconnected "Metric" list, or simply drop the measures `[AOV]`, `[Orders per Customer]`, `[Revenue per Customer]`, `[Repeat Purchase Rate]`, `[Cancellation Rate]`, `[Refund Rate]` into a single-column table with their `*_Period` / delta variants.

### Row E — Key Takeaways strip
Five small text/cards along the bottom. Fill with the **real findings from your data**:
- *Revenue grew **+11.4%** in Aug-18 vs Jul-18.*
- *Returning customers contribute **~54%** of realized revenue.*
- ***OVO** and **Debit Card** payments have the highest cancellation rates (>70%); **Kredivo** is lowest (~8%).*
- ***Jewelry (37%)** and **House (31%)** carry the strongest profit margins.*
- ***Jawa Barat (40%)** and **DKI Jakarta (21%)** are the top revenue provinces.*

---

## 7. Deneb visuals (limited & safe — KPI cards)

Use Deneb **only** for the executive KPI card and the small insight block. Charts stay native to avoid errors. This respects the rule: if Deneb risks failure, fall back to native.

### 7.1 KPI card — fields to add to the Deneb visual's *Values* well
Add exactly these (single-row, so measures return their single total — no filter-context collapse issue):
- `[NR Period (Label)]` → rename the role display as needed
- `[NR Delta (Label)]`
- `[NR Delta Dir]`
- a static title via the spec.

### 7.2 Vega-Lite spec (copy-paste, beginner-safe)

```json
{
  "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
  "background": "transparent",
  "config": { "view": { "stroke": null }, "font": "Segoe UI" },
  "width": 220,
  "height": 96,
  "data": { "name": "dataset" },
  "layer": [
    {
      "mark": { "type": "text", "align": "left", "fontSize": 12,
                "fontWeight": 600, "color": "#6B7280", "dx": 4, "dy": -34 },
      "encoding": { "text": { "value": "NET REVENUE" } }
    },
    {
      "mark": { "type": "text", "align": "left", "fontSize": 30,
                "fontWeight": 700, "color": "#1F2937", "dx": 4, "dy": 0 },
      "encoding": { "text": { "field": "NR Period (Label)", "type": "nominal" } }
    },
    {
      "mark": { "type": "text", "align": "left", "fontSize": 12,
                "fontWeight": 600, "dx": 4, "dy": 30 },
      "encoding": {
        "text": { "field": "NR Delta (Label)", "type": "nominal" },
        "color": {
          "field": "NR Delta Dir", "type": "nominal",
          "scale": { "domain": ["up", "down"], "range": ["#16A34A", "#DC2626"] },
          "legend": null
        }
      }
    }
  ]
}
```

**Notes that prevent the common failures:**
- `"data": { "name": "dataset" }` is the Power BI → Deneb binding. Keep it.
- The value/delta come in as **already-formatted text measures**, so there is **no aggregation inside Deneb** — this avoids the "DAX measure collapses to grand total" trap entirely.
- `"font": "Segoe UI"` is set in `config`, not inside a mark, so the vl-convert preview won't blank out.
- Duplicate this visual 5×, swapping the three field names + the title literal for each KPI.

If anything renders blank, switch that one card back to the native **Card (new)** visual — your dashboard still looks consistent.

---

## 8. Formatting settings (consistency checklist)

- **Canvas:** `#F5F6FA`; cards/visuals on white with 8 px rounded corners + soft shadow.
- **Accent palette (from the reference):** blue `#3B82F6`, orange `#F97316`, green `#22C55E`, purple `#8B5CF6`, teal `#14B8A6`. Set these in **View → Themes → Customize current theme → Data colors**.
- **Fonts:** Segoe UI throughout; KPI value 28–30 px bold, labels 11–12 px.
- **Currency:** rely on the `(Label)` measures (`Rp xx M / Rp xx Jt`) for KPIs and tooltips; in the Sales Summary table you can also set the column **Display units = Millions** with `Rp` prefix.
- **Delta colours:** green up / red down via the `*_Delta Dir` measures and conditional formatting.
- **Titles:** rename visual titles to match the reference ("REVENUE BY PAYMENT METHOD", "TOP 5 PRODUCTS BY REVENUE", etc.).

---

## 9. Interaction rules (verify before you finish)

1. **Period / Compare slicers** — confirm they change the KPI cards only (cards recalc; trend/category charts stay full-history). If a card doesn't change, check that the card uses a `*_Period` measure, not a base measure.
2. **Category slicer** — confirm it filters all charts. Use **Format → Edit interactions** to set the trend chart to **None** if you want a full-category trend.
3. **Cross-highlighting** — clicking a category bar should highlight the donut and products; leave default (highlight) interactions on for these.
4. **Default selections** — Period = Aug-18, Compare = Jul-18, Category = (all) so the report opens looking like the reference.

---

## 10. Error-prevention tips (the ones that actually bite)

- **Profit columns are amounts, not ratios.** `nett_profit_margin` is Rupiah. Don't divide it again — `Profit Margin %` already divides profit by revenue.
- **Disconnected tables must stay disconnected.** If you accidentally relate `Period`/`Compare` to `Sales`, the KPI logic breaks. Check **Model view**.
- **Mark the Date table** or time logic and the `Month-Year` sort can misbehave.
- **Sort `Month-Year` by `MY Key`**, else the trend axis goes alphabetical (Apr, Aug, Dec…).
- **`SELECTEDVALUE` returns blank** if a user multi-selects the Period slicer — keep Period/Compare on **Single select**.
- **Don't sum `margin_ratio`.** It's a fixed per-category attribute; use it as Average or just as a label.
- **Realized-status list:** if your audience wants gross figures, make a parallel `Gross Revenue` measure with `SUM(Sales[gross_revenue])` and no status filter — keep both, don't overwrite.
- **Deneb blank card?** 99% of the time it's an aggregation in the spec or a font set inside a mark. Use pre-formatted text measures (as above) and set font in `config`.
- **Performance:** 267k rows is small for Power BI; if it feels slow, turn off "Auto date/time" in Options (you already have a proper Date table).

---

### Quick build order recap
1. M preprocessing → `Sales` (Step 2)  →  2. `Date`, `Period`, `Compare` tables + relationship (Step 3)  →  3. Slicers (Step 4)  →  4. Measures + Rupiah format (Step 5)  →  5. KPI cards → charts → takeaways (Step 6)  →  6. Optional Deneb cards (Step 7)  →  7. Format + interactions + QA (Steps 8–10).

You'll end with an executive dashboard that matches the reference layout and is grounded entirely in your real numbers.
