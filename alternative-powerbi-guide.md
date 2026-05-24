---
name: powerbi-household-budget
description: >
  Use this skill whenever the user wants to build, design, or improve a household budget
  dashboard or report from personal expense data. Triggers include: uploading a CSV or
  Excel file with transaction/expense data, asking to visualize spending by category,
  requesting a monthly budget dashboard, or asking how to track income vs expenses over
  time. Also use when the user mentions tracking groceries, mortgage, utilities, fast food,
  gas, or any personal spending — even if they don't say "dashboard" or "budget" explicitly.
  Works for both Excel output (default, most accessible) and Power BI guidance.
---

# Household Budget Dashboard Skill

Builds a polished, ready-to-open Excel dashboard from an uploaded expense file. Also
provides Power BI guidance if the user prefers that route.

**Default output: Excel (.xlsx)** — no software licence needed, works for everyone.
**Alternative: Power BI (.pbix guidance)** — see Step 6 if the user prefers Power BI.

---

## Step 1 — Read & Validate the Input File

Use the `file-reading` skill to read the uploaded file. The expected structure is a
multi-sheet Excel workbook with one sheet per month plus an optional Summary sheet, or a
single-sheet CSV with one row per transaction.

### Expected columns (names may vary — map intelligently)

| Expected | Common Aliases |
|----------|---------------|
| Date | Transaction Date, Date Posted |
| Description | Merchant, Payee, Note |
| Category | Type, Tag, Group |
| Withdrawals | Debits, Amount Out, Expenses |
| Deposits | Credits, Amount In, Income |

**Real-world data notes learned from this user's file:**
- Multi-sheet workbooks with one tab per month (e.g. Jan2026, Feb2026) are common
- Each month sheet has a summary block to the right of the transactions (cols G+) — ignore it
- The last row of each sheet contains withdrawal/deposit totals — skip it when reading transactions
- Dates may be stored as Excel serial numbers (e.g. 46055) — pandas handles these automatically
- Amounts are always positive; Withdrawals = money out, Deposits = money in
- A separate "Summary Sheet" tab may exist with pre-aggregated category totals — use it to
  cross-check or as the primary data source for the category breakdown sheet

### Real categories seen in this household's data
Bank Fees, Bryce, Bryce RDSP, Cell Phones, Clothing, Credit Cards, Entertainment,
Fast Food, Furniture, Gas, Gidget, Givings, Golf, Groceries, House Cleaning, House Repair,
Insurance (Auto/Home/Life), Internet, Judy, Lessons, Loan, Medication, Misc, Mortgage,
Mutual Funds, Utilities, Vehicle, and various income types (J Income, R Income, B Income).

If categories differ, map to sensible groups and confirm with the user.

---

## Step 2 — Build the Excel Dashboard (4 Sheets)

Use `openpyxl` for formatting and charts. Use `pandas` for data loading and aggregation.
Always use Excel formulas rather than hardcoded Python values where possible.

### Colour palette (use consistently)
```python
NAVY   = '2E4057'   # headers, title bars
TEAL   = '048A81'   # income, positive values
ORANGE = 'F46036'   # expenses
LGREY  = 'F5F7FA'   # alternating row background
WHITE  = 'FFFFFF'
GREEN  = '2D6A4F'   # surplus, deposits
RED    = 'C1121F'   # deficit, withdrawals
```

### Font: Arial throughout, size 10 body / 14-18 titles

---

### Sheet 1: Dashboard

**Layout:**
1. Title bar (merged, NAVY, white text) — rows 1-2
2. KPI Cards row — 4 cards across columns A-L, rows 4-6:
   - Total Income YTD (TEAL)
   - Total Expenses YTD (ORANGE)
   - Net Balance YTD (GREEN if positive, RED if negative)
   - Avg Monthly Spend (NAVY)
3. Hidden chart data table (rows 8-14) — monthly Income / Expenses / Net
4. Hidden category data (cols G-H, rows 8-16) — top 8 categories by spend
5. Hidden cumulative data (cols J-K, rows 8-14) — running total by month
6. Three charts starting at row 15:
   - **Bar chart** (A15): Clustered Income vs Expenses by month — TEAL/ORANGE
   - **Pie/Donut chart** (G15): Top 8 spending categories
   - **Line chart** (J15): Cumulative spending YTD — NAVY line

**KPI card pattern** (avoid MergedCell write errors):
```python
def merge_and_set(ws, merge_range, value, font=None, fill=None, alignment=None, num_fmt=None):
    ws.merge_cells(merge_range)
    first = merge_range.split(':')[0]
    cell = ws[first]
    cell.value = value
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if num_fmt: cell.number_format = num_fmt
```
Always use `merge_and_set()` — never assign `ws['A1'] = value` after merging that cell.

---

### Sheet 2: Monthly Breakdown

- Title bar
- Header row: Category | Jan | Feb | Mar | Apr | May | Total | % of Total
- One row per expense category, alternating LGREY/WHITE rows
- Totals row in TEAL at the bottom
- % of Total column formatted as `0.0%`
- Horizontal bar chart below the table: Top 10 categories by YTD total spend (NAVY bars)
- Exclude income/transfer categories from this sheet

**Categories to exclude from expense breakdown:**
J Income, R Income, B Income, Gift, Reimbursed, Alie, Alie Rent, Alie Table & Couch,
Income Tax, Credit Cards, Mutual Funds, Bryce RDSP, Bank Fees, Property Taxes,
Dwyer Analytics, Hugalot CC

---

### Sheet 3: Income vs Expenses

- Title bar
- Header row: Month | Income | Expenses | Net | Surplus/Deficit? | Cumulative Spend
- One row per month
- Net column: GREEN text if positive, RED text if negative
- Surplus/Deficit column: '✓ Surplus' in GREEN or '✗ Deficit' in RED
- Totals row in TEAL
- Clustered bar chart below: Income vs Expenses side by side per month

---

### Sheet 4: Transactions

- Title bar
- Header row: Month | Date | Description | Category | Withdrawals | Deposits
- All transactions from all month sheets, one row per transaction
- Withdrawals in RED text, Deposits in GREEN text
- Alternating LGREY/WHITE row background
- Column widths: Month=8, Date=10, Description=36, Category=22, Withdrawals=13, Deposits=13
- Apply thin borders throughout (`CCCCCC`)

---

## Step 3 — Data Loading Pattern

```python
import pandas as pd

month_sheets = ['Jan2026', 'Feb2026', ...]  # discover from wb.sheetnames
transactions = []

for sheet in month_sheets:
    df = pd.read_excel(filepath, sheet_name=sheet)
    df.columns = df.columns.str.strip()
    # Keep only transaction columns
    df = df[['Description','Category','Withdrawals','Deposits','Date']].copy()
    df = df.dropna(subset=['Description'])
    df = df[df['Category'].notna() & df['Date'].notna()]  # drops total row
    df['Withdrawals'] = pd.to_numeric(df['Withdrawals'], errors='coerce').fillna(0)
    df['Deposits']    = pd.to_numeric(df['Deposits'],    errors='coerce').fillna(0)
    df['Month'] = sheet[:3]  # 'Jan', 'Feb', etc.
    transactions.append(df)

all_tx = pd.concat(transactions, ignore_index=True)
```

For the Summary Sheet, read col 0 as Category and cols 1-6 as Jan/Feb/Mar/Apr/May/Total.
Strip column name spaces (note: 'Mar ', 'Apr ', 'May ' may have trailing spaces).

---

## Step 4 — Recalculate & Verify

Always run after saving:
```bash
python scripts/recalc.py output.xlsx 30
```
Check that `"status": "success"` and `"total_errors": 0` before presenting the file.

---

## Step 5 — Present the File

Use `present_files` to share the output. After sharing, highlight 2-3 observations from
the actual data (e.g. biggest category, any deficit months, notable trends).

---

## Step 6 — Power BI Alternative (if user prefers)

If the user wants Power BI instead of Excel, offer two options:

**Option A — .pbit Template approach (recommended):**
1. User builds a blank `.pbix` in Power BI Desktop with the data model, DAX measures,
   and visual layout they want
2. User saves it as File → Export → Power BI Template (.pbit)
3. Claude generates a Python script that cleans the user's CSV/Excel and outputs a
   structured Excel file matching the template's schema exactly
4. User opens the .pbit and points it at Claude's output file

**Option B — DAX measures + layout guidance:**
See `references/powerbi-dax.md` for the full set of DAX measures and Power BI layout
guidance (covers data model, measures, visual config, formatting).

Note: Generating a .pbix from scratch programmatically is not reliably possible —
Microsoft's format is partially proprietary. The .pbit + Python data-prep approach is
the most practical fully-built Power BI solution.

---

## Reference Files

- `references/sample-schema.md` — example CSV/Excel input format
- `references/alternative-powerbi-guide.md` — DAX measures and Power BI layout guide (optional path)
