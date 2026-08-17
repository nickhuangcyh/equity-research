---
name: 3-statement-model
description: Complete and populate 3-statement financial model templates (Income Statement, Balance Sheet, Cash Flow Statement). Use when asked to fill out model templates, populate financial models with data, or link integrated financial statements within an existing template structure.
---

# 3-Statement Financial Model Template Completion

Complete and populate integrated financial model templates with proper linkages between Income Statement, Balance Sheet, and Cash Flow Statement.

## ⚠️ CRITICAL PRINCIPLES — Read Before Populating Any Template

**Environment — Office JS vs Python:**
- **If running inside Excel (Office Add-in / Office JS):** Use Office JS directly. Write formulas via `range.formulas = [["=D14*(1+Assumptions!$B$5)"]]` — never `range.values` for derived cells.
- **If generating a standalone .xlsx file:** Use Python/openpyxl. Write `ws["D15"] = "=D14*(1+Assumptions!$B$5)"`, then run `recalc.py` before delivery.
- **Office JS merged cell pitfall:** Write value to top-left cell alone, then merge + format the full range.

**Formulas over hardcodes (non-negotiable):**
- Every projection cell, roll-forward, linkage, and subtotal MUST be an Excel formula
- The ONLY cells with hardcoded numbers: (1) historical actuals, (2) assumption drivers in the Assumptions tab

**Verify step-by-step with the user:**
1. After mapping the template → confirm tabs/sections identified
2. After populating historicals → confirm values/periods match source data
3. After building IS projections → run subtotal checks, confirm before moving to BS
4. After building BS → show balance check (Assets = L+E) for every period
5. After building CF → show cash tie-out (CF ending cash = BS cash), confirm before finalizing
6. **Do NOT populate the entire model end-to-end** — break at each statement

## Formatting — Professional Blue/Grey Palette

| Element | Fill | Font |
|---|---|---|
| Section headers (IS / BS / CF titles) | Dark blue `#1F4E79` | White bold |
| Column headers (FY2024A, FY2025E, etc.) | Light blue `#D9E1F2` | Black bold |
| Input cells (historicals, assumption drivers) | Light grey `#F2F2F2` or white | Blue `#0000FF` |
| Formula cells | White | Black |
| Cross-tab links | White | Green `#008000` |
| Check rows / key totals | Medium blue `#BDD7EE` | Black bold |

---

## Template Analysis Phase — Do This First

1. **Identify Input vs. Formula Cells** — look for visual cues (font color, cell shading)
2. **Map the Template's Flow** — which tabs feed into others (Assumptions → IS → BS → CF)
3. **Understand Column Structure** — historical vs projected columns, fiscal year notation
4. **Check for Named Ranges** — Revenue growth rates, cost percentages, key outputs

---

## Common Tab Names

| Common Tab Names | Contents |
|---|---|
| IS, P&L, Income Statement | Income Statement |
| BS, Balance Sheet | Balance Sheet |
| CF, CFS, Cash Flow | Cash Flow Statement |
| WC, Working Capital | Working Capital Schedule |
| DA, D&A, Depreciation | Depreciation & Amortization Schedule |
| Debt, Debt Schedule | Debt Schedule |
| Assumptions, Inputs, Drivers | Driver assumptions and inputs |
| Checks, Audit, Validation | Error-checking dashboard |

---

## Filling in Data Without Breaking Formulas

| Rule | Description |
|---|---|
| Only edit input cells | Never overwrite cells containing formulas |
| Match the template's units | Verify if template uses thousands, millions, or actual values |
| Respect sign conventions | Follow the template's existing sign convention |
| Check for circular references | Ensure "Enable Iterative Calculation" is on if model uses iterative calcs |

---

## Scenario Analysis (Base / Upside / Downside)

Use a scenario toggle (dropdown) in the Assumptions tab with CHOOSE or INDEX/MATCH formulas.

| Scenario | Description |
|---|---|
| Base Case | Management guidance or consensus estimates |
| Upside Case | Above-guidance growth, margin expansion |
| Downside Case | Below-trend growth, margin compression |

**Key Drivers to Sensitize:** Revenue growth, Gross margin, SG&A %, DSO/DIO/DPO, CapEx %, Interest rate, Tax rate

---

## Sign Convention Reference

| Statement | Item | Sign Convention |
|---|---|---|
| CFO | D&A, SBC | Positive (add-back) |
| CFO | ΔAR (increase) | Negative (use of cash) |
| CFO | ΔAP (increase) | Positive (source of cash) |
| CFI | CapEx | Negative |
| CFF | Debt issuance | Positive |
| CFF | Debt repayments | Negative |
| CFF | Dividends | Negative |

---

## Cross-Statement Integrity Checks

| Check | Formula | Expected Result |
|---|---|---|
| Balance Sheet Balance | Assets - Liabilities - Equity | = 0 |
| Cash Tie-Out | CF Ending Cash - BS Cash | = 0 |
| Net Income Link | IS Net Income - CF Starting Net Income | = 0 |
| Retained Earnings | Prior RE + NI - Dividends - BS Ending RE | = 0 |

---

## Quality Checks by Sheet

**Income Statement (IS) Quality Checks**
- Revenue figures match source data for historical periods
- All expense line items sum to reported totals
- Tax calculation logic is appropriate (handles losses correctly)
- Forecast drivers reference assumptions tab (no hardcodes)

**Balance Sheet (BS) Quality Checks**
- Assets = Liabilities + Equity for every period (primary check)
- Cash balance matches Cash Flow Statement ending cash
- Retained Earnings rolls forward correctly: Prior RE + Net Income - Dividends = Ending RE

**Cash Flow Statement (CF) Quality Checks**
- Net Income at top of CFO matches Income Statement Net Income
- Non-cash add-backs (D&A, SBC) tie to source schedules
- Working capital changes have correct signs
- Ending Cash matches Balance Sheet Cash
- Beginning Cash equals prior period Ending Cash

---

## Circular Reference Handling

Interest expense creates circularity: Interest → Net Income → Cash → Debt Balance → Interest

Enable iterative calculation in Excel: File → Options → Formulas → Enable iterative calculation (max iterations: 100, max change: 0.001).

---

## Common Formula Issues to Watch For
- Mixed absolute/relative references causing incorrect results when copied
- Broken links to external files or deleted ranges (#REF! errors)
- Division by zero in early periods before revenue ramps (#DIV/0! errors)
- Inconsistent formulas across projection columns (use Ctrl+\ to find differences)

---

## Command Workflow: /3-statement-model

Fill out a 3-statement financial model template.

If a file path is provided, use it as the template. Otherwise ask the user for their model template.

Steps:
1. Map the template structure — identify all tabs and sections
2. Confirm historical data to populate
3. Review assumption drivers
4. Build IS projections → confirm with user
5. Build BS → verify balance check
6. Build CF → verify cash tie-out
7. Run cross-statement integrity checks
8. Deliver completed model

Run `python recalc.py model.xlsx 30` before delivery. Fix ALL errors until status is "success".
