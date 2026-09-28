# link-profit-scenario

A WorkBuddy / Agent Skill for analyzing the profit model of e-commerce operation workbooks and answering **what-if scenario questions** such as:

> "If ad spend drops 10% and sales volume drops 5%, how does the listing's profit change?"

Built for the shared Tmall store-operation workbook template (a 50 MB macro-enabled `.xlsm` with source-data sheets, an operation dashboard, and a 「【预测】链接利润」 forecast sheet).

## What it does

1. **Extracts the real profit formula** from the workbook's own cells (dynamic-array formulas: `LET` / `XLOOKUP` / `GROUPBY` / `FILTER`), instead of assuming a textbook formula.
2. **Verifies** by reproducing the sheet's profit column — the analysis only proceeds when the reproduction matches the sheet within 0.01.
3. **Runs the scenario**: every cost item is rescaled by its true driver (sales-linked / volume-linked / ad-spend-linked / fixed) and a per-item comparison table is produced.

## Core insight it encodes

In this model every cost is **variable** — there is no fixed cost. So a uniform x% scale-down scales profit by exactly x% and leaves the **profit margin unchanged**. Margin only moves when the scaling is asymmetric (e.g. spend −10% while sales only −5%) or when the per-unit cost structure changes (shipping, purchase cost, unit price).

## Key practices

- Always read the **newest** file the user supplied; re-read the original rather than reusing any working copy (workbooks can be re-saved several times a day and the formulas change).
- Locate data by header + formula, never by hard-coded column letters (layouts differ between store versions).
- Cross-check the gross-sales column formula: it has several versions, and one of them contains a per-unit correction hard-coded for a specific product ID.

## Usage

Place `SKILL.md` in your skills directory, e.g.

```
~/.workbuddy/skills/link-profit-scenario/SKILL.md
```

The agent picks it up automatically when the user asks a profit-scenario question about such a workbook.

## Sanitization notice

All shop names, product IDs, fee rates and financial figures in this repository are **placeholders**. The skill documents the *methodology* only — no real business data is included.

## License

MIT
