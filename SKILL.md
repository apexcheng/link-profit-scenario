---
name: link-profit-scenario
description: Analyze the profit formula of e-commerce operation workbooks (xlsm/xlsx, sheet 「【预测】链接利润」 or similar) and answer "what-if" scenario questions such as "if ad spend drops x%, sales volume drops y%, how does profit change". Trigger words: 利润怎么算 / 降花费 / 降销售额 / 情景测算 / 链接利润 / 预测利润. Works with the shared workbook template used by Tmall stores.
---

# Link Profit Scenario (链接利润情景测算)

Answer questions of the form "if spend drops x%, sales/volume drop y%, how does profit change".

**Core principle: extract the real calculation rules from the workbook's own formulas, reproduce the sheet's profit figure as verification, and only then run the scenario. Never guess from generic formulas.**

> ⚠️ This skill contains the *methodology* only. All shop names, product IDs, fee rates and financial figures have been replaced with placeholders. Fill in the values extracted from your own workbook.

## Iron rule: always use the newest file the user just handed you

**Assume the file has changed on every request.** In practice, the same workbook may be re-saved multiple times a day, with formula structure and cached values changing each time — previous conclusions become invalid immediately.

1. **Read the original file directly** (`openpyxl.load_workbook(path, read_only=True)` is pure-read: it does not write to disk and always gives the newest version). Do **not** copy a working duplicate as a "safety" measure — a duplicate is a stale snapshot and is the root cause of wrong answers.
2. **Only fall back to copying** when the original is locked (`PermissionError`, file held exclusively by Excel/WPS): then `cp` a timestamped copy (`copy_YYYYMMDD_HHMM.xlsm`) just for that one analysis; never reuse it later.
3. **When the value the user sees differs from what you last read, trust the user's value** — re-read the current file first instead of defending the earlier reading.
4. Column layouts and parameter locations differ between store versions. Always locate data by the row-1 headers and the row-2 formulas, never by hard-coded column letters.
5. The only hard prohibition: never call `save()` on a workbook loaded with `data_only=False`. The analysis flow never writes back.

## The profit model (verified, shared Tmall workbook template)

### Core formula (column D, sheet 「【预测】链接利润」)

```
Profit D = C*(1-R) - G*(1-R) - [ X + AN + T1 + V + W + AK + AM + AA + AB + AC + T2 + AD ]

C  = O / 1.13                     net (tax-exclusive) sales   (O = gross sales)
O  = gross sales. The O-column formula has multiple versions — always re-read it:
     ① O = N - P                 (newer version; P = subsidy offset column)
     ② O = N - subsidy - extra subsidy + order-detail aggregation + per-unit correction
        ⚠️ correction term may be hard-coded for ONE product ID ( e.g. (price_a - price_b) * qty )
   → Re-read the O formula on every analysis; do not reuse a previously read structure.
G  = purchase cost by product ID (side summary area, split into normal / proxy orders)
R  = return amount % for that listing (from the return-rate sheet; historical, held constant in scenarios)
T1 = sunk cost = shipping cost of sales orders (from the shipping-cost sheet, ≈ per-unit × qty)
T2 = reship / exchange loss (non-sales-order amount);  V = after-sales loss
X  = O*(1-R)*store misc fee %        AN = C*(1-R)*5% (hard-coded 5% reserve)
AA = O*(1-R)*coin rebate %           AB = affiliate commission (source sheet)
AC = brand program; AD = subsidy giveaway; W = proxy-order cost  (may all be 0)
AK = AJ / 1.06   ad spend (gross ÷ 1.06 → net)
AM = (Z platform service fee % + AE + Y cross-border) ÷ 1.06   net service fee
```

### Where the parameters live
Fee rates and the reporting window sit in the dashboard sheet (`经营看板`) parameter block — e.g. Y4:Y9 for rates, and a separate pair of cells for window start/end. **Their row positions differ between workbook versions** (Y26/Y27 in one version, Y93/Y94 in another): confirm the exact cells from the formulas themselves, never hard-code them.

Typical parameters: store misc fee %, gross→net sales coefficient (e.g. 1.13), service-fee tax coefficient (e.g. 1.06), coin rebate %, platform service fee %, subsidy service fee %.

### Scenario linkage rules (how each cost item scales)

| Linkage | Cost items | Scaling factor |
|---|---|---|
| Follows sales | store misc fee, 5% reserve, coin rebate, platform service fee, cross-border fee, affiliate commission | k_sales |
| Follows volume | purchase cost, sunk cost (shipping), reship/exchange loss, after-sales loss, any per-unit correction term inside O | k_qty |
| Follows ad spend | ad spend (keyword / site-wide / audience / content) | k_ad |
| Unchanged | return rate R (historical per listing), brand program, subsidy giveaway (if fixed) | 1 |

**Key conclusion (safe to quote):** in this model *everything is variable* — there is no fixed cost. Therefore a uniform x% scale-down scales profit by exactly the same factor and leaves the profit margin unchanged. To change the margin you must either break the symmetry (e.g. spend −10% while sales only −5%) or change the per-unit cost structure (shipping / purchase cost / unit price).

## Workflow

### Step 1 — file & environment
1. Re-read the original file directly (see iron rule); copy only if locked.
2. Macro-enabled `.xlsm` files usually cannot be opened by spreadsheet MCP tools — fall back to `openpyxl`.
3. Python env: any env with `openpyxl` installed. (On managed Windows setups, remember to clear the safe-delete session env vars when creating a venv, or `rmtree` inside venv creation fails.)

### Step 2 — extraction (two loads)
- `load_workbook(data_only=True)`: cached values. The forecast sheet lists one product ID per row starting at row 2 (column A).
- `load_workbook(data_only=False)`: formulas. **Formulas come back as `ArrayFormula` objects — take `.text`.** Modern functions are stored with `_xlfn.` / `_xlpm.` prefixes (LET / ANCHORARRAY / XLOOKUP / GROUPBY / FILTER).
- Side summary area: column AX onward, header on row 2; holds non-proxy sales, purchase cost, proxy-order items, cross-border fee.

### Step 3 — verification (mandatory; do not skip)
Reproduce column D from the extracted parameters and compare with the cached value (tolerance < 0.01):
```python
D = C*(1-R) - G*(1-R) - (X+AN+T1+V+W+AK+AM+AA+AB+AC+T2+AD)
```
Also verify the O column **against the version actually present in the file**:
- `O = N - P`: check P against the subsidy figures of the source sheets.
- Aggregating versions: `O = N - Σsubsidy(window) + correction`, filtered to shipped orders in the reporting window.

If it does not reconcile, first check proxy orders (when W > 0 the purchase-cost basis changes) and whether the product ID sits in the subsidy preset list.

### Step 4 — scenario calculation
Rescale each item by k_sales / k_qty / k_ad per the linkage table. Output a comparison table: per item "original / new / delta", plus profit, margin and improvement. Reconciliation view: `net improvement = money saved (variable costs + ad spend) − revenue given up`.

### Step 5 — presentation
- Detail tables in three blocks: revenue side / cost side / result.
- Explain *why* the margin moved or did not move (whether the scaling was symmetric).
- Point out the listing's dominant cost (e.g. shipping share) and give control suggestions.
- State the model assumptions explicitly: the real elasticity of sales to ad-spend cuts must be validated with actual campaign data.

## Checklist when switching store / period
1. Fee-rate parameters may differ (different platform commissions); the window parameter cells may sit in different rows.
2. **The column layout may shift entirely** (an extra column inserted; columns re-ordered). Always locate by header + formula, not by fixed column index.
3. The D-column formula structure may differ between versions — re-read row 2.
4. Check whether the O-column correction term targets another product ID, or has been removed.
5. Only after verification passes may the linkage table be used for scenarios.

## Notes
- Large workbooks (tens of MB): use `read_only` mode and target sheets by name; full-sheet iteration takes tens of seconds.
- In some source sheets the column that looks like "amount" is actually a subsidy-giveaway amount, not revenue — confirm from the header.
- "Sunk cost" in this template = shipping-type spend (parcel fee + allocation + back-office).
- `EmptyCell` has no `.column` attribute: iterate with `values_only=True` in read_only mode.
