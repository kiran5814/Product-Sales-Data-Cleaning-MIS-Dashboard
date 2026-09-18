# Product Sales Data Cleaning & MIS Dashboard

**Author:** Katta Satya Chandra Kiran | [LinkedIn](https://www.linkedin.com/in/satya-chandra-kiran-katta-b49a84286) | [GitHub](https://github.com/kiran5814)

A raw, messy sales export is cleaned, validated and turned into a live MIS dashboard using Excel formulas only — no code required — the same way a real MIS Analyst maintains a business reporting pipeline.

---

## 1. Business Problem

The sales team exports a raw transaction file every period. Before it can be used for reporting, it needs to be checked for missing values, inconsistent entries and duplicate records, and enriched with the fields Operations actually reports on (order month, pre-discount value, discount amount). This project documents that process end-to-end — from raw export to a decision-ready dashboard.

## 2. Files in this Repo

| File | Description |
|---|---|
| `Product-Sales_Raw.xlsx` | Raw sales export as received — 1,502 rows, 19 columns |
| `Product_Sales_Cleaned.xlsx` | Cleaned dataset + MIS summary sheets + dashboard — 1,500 rows, 24 columns |
| `README.md` | This file |

**Dataset:** 1,500 transactions, Jan 2023 – Jun 2025, across 5 regions, 7 products, 6 salespeople and 4 store locations.

## 3. Data Quality Issues Found in the Raw Extract

| Issue | Detail | Rows Affected |
|---|---|---|
| Duplicate records | 2 fully duplicated rows (`REG100000`, `REG100001` each appeared twice) | 2 |
| Missing values | Nulls across `Date`, `UnitPrice`, `StoreLocation`, `Discount`, `TotalPrice`, `Promotion`, `CustomerName`, `ShippingCost` | 378 cells |
| Blank `Promotion` field | No promo code applied, left blank instead of labelled | 370 |
| Inconsistent casing in `Region` | `EAST`/`NORTH`/`CENTRAL`/`SOUTH` appearing alongside proper-case `East`/`North`/`Central`/`South` | 5 |

**Total raw null cells: 378 → reduced to 8 after cleaning (97.9% reduction).**

## 4. Cleaning Steps & Formulas Used

**Step 1 — Remove exact duplicate rows**
Identified via Excel's Remove Duplicates on the full row (confirmed both copies of `REG100000` and `REG100001` were byte-for-byte identical, not just same OrderID) → 1,502 rows → 1,500 rows.

**Step 2 — Standardise the blank `Promotion` field**
```
=IF(ISBLANK([@Promotion]),"No Promotion",[@Promotion])
```
370 blank cells relabelled from *null* to an explicit `"No Promotion"` category, so it reports correctly instead of silently disappearing from pivot counts.

**Step 3 — Derive time-intelligence columns from `Date`**
```
Year       = YEAR([@Date])
Month      = TEXT([@Date],"MMM")
Year Month = [@Year] & " " & [@Month]
```
Used to build the monthly trend view without touching the source `Date` column.

**Step 4 — Rebuild the discount waterfall and cross-check `TotalPrice`**
```
TotalPriceBeforDiscount = [@Quantity] * [@UnitPrice]
DiscountAmount          = [@TotalPriceBeforDiscount] * [@Discount]
Check                   = [@TotalPriceBeforDiscount] - [@DiscountAmount]   ' should equal TotalPrice
```
Validated `Quantity × UnitPrice × (1 − Discount)` against the original `TotalPrice` column for every row — confirms the raw `TotalPrice` figures are trustworthy (0 rows off by more than a rounding cent, aside from one row differing by ₹10.93, flagged for review).

**Step 5 — Sanity checks (no formula changes needed, just validation)**
- `Quantity` range: 1–20 units — no zero or negative quantities
- `UnitPrice` range: ₹5.52–₹599.72 — no zero or negative prices
- `Discount` only takes 4 values: 0%, 5%, 10%, 15% — confirms a fixed discount-tier policy, not free-text entry
- `DeliveryDate ≥ OrderDate` for all 1,500 rows — no logistics-date errors

## 5. Known Limitations (documented, not hidden)

Flagging these instead of silently ignoring them is the point of a proper data-quality log:

- **5 rows still have inconsistent `Region` casing** (`EAST`, `NORTH`, `CENTRAL`, `SOUTH`) — not yet normalised with `=PROPER([@Region])`. Left as a follow-up item.
- **8 cells remain null** — 1 missing `Date`, 2 missing `StoreLocation`, 1 missing `CustomerName`, 1 missing `ShippingCost` — insufficient information in the source export to safely impute these; better to flag than guess.
- The `Region Analysis` summary currently totals 1,499 orders against 1,500 cleaned rows — the 1 row with a missing `Date` is excluded from date-based grouping, which is expected but worth noting for anyone auditing the totals.

## 6. Key Metrics (Post-Cleaning)

| KPI | Value |
|---|---|
| Total Sales | ₹43,79,992 |
| Total Sales Before Discount | ₹47,27,880 |
| Total Orders | 1,500 |
| Total Quantity Sold | 15,616 units |
| Average Unit Price | ₹298.83 |
| Total Discount Given | ₹3,47,899 |
| Total Shipping Cost | ₹40,444 |
| Returned Orders | 372 (24.8% return rate) |

## 7. Dashboard

`Product_Sales_Cleaned.xlsx` includes a one-page **MIS Dashboard** sheet built entirely from the cleaned data:

- KPI cards (Total Sales, Orders, Quantity, Avg Unit Price) — formula-linked to the summary sheet, not hardcoded, so they refresh automatically
- Sales by Region (column chart)
- Sales by Product, ranked (column chart)
- Salesperson performance — Sales (bars) vs Average Order Value (secondary-axis line)
- Monthly sales trend (Jan 2023 – Jun 2025, line chart)

Supporting pivot-style summary sheets: `Region Analysis`, `Product Analysis`, `Sales Person Analysis`, `Monthly Analysis` — all formula-driven (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`) so they recalculate if the underlying data changes.

## 8. Tools Used

Microsoft Excel — formulas (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `IF`, `TEXT`, `YEAR`), PivotTables, native charts, conditional data validation. No external scripts or add-ins.

## 9. What This Project Demonstrates

- Reading and auditing a raw data export before trusting it
- A documented, repeatable cleaning process (not one-off manual edits)
- Validating computed fields against source figures instead of assuming they're correct
- Honest data-quality logging — surfacing what's still imperfect rather than hiding it
- Turning cleaned data into a decision-ready MIS dashboard with live, formula-driven KPIs
