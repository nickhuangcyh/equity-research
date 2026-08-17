---
name: dcf-model
description: Build a DCF (Discounted Cash Flow) valuation model for equity valuation with WACC, sensitivity analysis, and Bear/Base/Bull scenarios in Excel. Use for intrinsic value analysis, growth projections, and terminal value calculations.
---

# DCF Model Builder

## Overview

This skill creates institutional-quality DCF models for equity valuation following investment banking standards. Each analysis produces a detailed Excel model (Bear/Base/Bull scenarios, sensitivity analysis at the bottom of the DCF sheet).

## Tools

Default to using all information provided by the user and MCP servers available for data sourcing.

## Critical Constraints — Read These First

**Environment: Office JS vs Python/openpyxl:**
- **If running inside Excel (Office JS):** Use Office JS directly. Write formulas via `range.formulas = [["=D19*(1+$B$8)"]]`. No separate recalc needed.
- **If generating a standalone .xlsx file:** Use Python/openpyxl. Write formula strings `ws["D20"] = "=D19*(1+$B$8)"`, then run `recalc.py` before delivery.

**⚠️ Office JS merged cell pitfall:** Do NOT call `.merge()` then set `.values` on the merged range. Write value to the top-left cell alone, then merge + format:
```js
// CORRECT:
ws.getRange("A7").values = [["MARKET DATA & KEY INPUTS"]];
const hdr = ws.getRange("A7:H7");
hdr.merge();
hdr.format.fill.color = "#1F4E79";
hdr.format.font.bold = true;
hdr.format.font.color = "#FFFFFF";
```

**Formulas Over Hardcodes (NON-NEGOTIABLE):**
- Every projection, margin, discount factor, PV, and sensitivity cell MUST be a live Excel formula
- Only permitted hardcodes: raw historical inputs, assumption drivers, current market data

**Verify Step-by-Step With the User:**
- After data retrieval → confirm raw inputs before projecting
- After revenue projections → confirm top line before building margin build
- After FCF build → confirm FCF schedule before computing WACC
- After WACC → confirm before discounting
- After terminal value + PV → confirm equity bridge before sensitivity tables

**Sensitivity Tables:**
- Use an **ODD number of rows and columns** (standard: 5×5) — guarantees a true center cell
- **Center cell = base case** — axis values symmetric around actual model assumptions
- **Highlight center cell** with medium-blue fill `#BDD7EE` + bold font
- Populate ALL cells (3 tables × 25 cells = 75) with full DCF recalculation formulas

**Cell Comments:**
- Add cell comments AS each hardcoded value is created
- Format: "Source: [System/Document], [Date], [Reference], [URL if applicable]"

---

## DCF Process Workflow

### Step 1: Data Retrieval and Validation

**Data Sources Priority:**
1. MCP Servers (if configured) - Daloopa, FactSet, S&P Kensho
2. User-Provided Data - Historical financials
3. Web Search/Fetch - Current prices, beta, debt when needed

**Validation Checklist:**
- Verify net debt vs net cash
- Confirm diluted shares outstanding
- Validate historical margins consistent with business model
- Cross-check revenue growth rates with industry benchmarks
- Verify tax rate is reasonable (typically 21-28%)

### Step 2: Historical Analysis (3-5 years)

Analyze: Revenue growth trends, margin progression, capital intensity (D&A and CapEx as % of revenue), working capital efficiency, return metrics (ROIC, ROE).

### Step 3: Build Revenue Projections

**Three-scenario approach:**
```
Bear Case: Conservative growth (e.g., 8-12%)
Base Case: Most likely scenario (e.g., 12-16%)
Bull Case: Optimistic growth (e.g., 16-20%)
```

Revenue(Year N) = Revenue(Year N-1) × (1 + Growth Rate)

**Growth Rate Framework:**
- Year 1-2: Higher growth (near-term visibility)
- Year 3-4: Gradual moderation toward industry average
- Year 5+: Approaching terminal growth rate

### Step 4: Operating Expense Modeling

ALL percentages based on **REVENUE**, not gross profit. Model operating leverage: % should decline as revenue scales.

- Sales & Marketing: 15-40% of revenue
- R&D: 10-30% for technology companies
- G&A: 8-15% of revenue

### Step 5: Free Cash Flow Calculation

```
EBIT
(-) Taxes (EBIT × Tax Rate)
= NOPAT
(+) D&A (non-cash, % of revenue)
(-) CapEx (% of revenue, typically 4-8%)
(-) Δ NWC (change in working capital)
= Unlevered Free Cash Flow
```

### Step 6: Cost of Capital (WACC)

**CAPM for Cost of Equity:**
```
Cost of Equity = Risk-Free Rate + Beta × Equity Risk Premium
(ERP = 5.0-6.0%)
```

**WACC Calculation:**
```
Market Cap = Current Stock Price × Shares Outstanding
Net Debt = Total Debt - Cash & Equivalents
EV = Market Cap + Net Debt
WACC = (Cost of Equity × Equity Weight) + (After-Tax Cost of Debt × Debt Weight)
```

**Typical WACC Ranges:** Large Cap: 7-9%, Growth: 9-12%, High Growth/Risk: 12-15%

### Step 7: Discount Rate Application

**Mid-Year Convention:** Discount Period: 0.5, 1.5, 2.5, 3.5, 4.5
Discount Factor = 1 / (1 + WACC)^Period

### Step 8: Terminal Value Calculation

**Perpetuity Growth Method (Preferred):**
```
Terminal FCF = Final Year FCF × (1 + Terminal Growth Rate)
Terminal Value = Terminal FCF / (WACC - Terminal Growth Rate)
```

Terminal Growth Rate: 2.0-3.5% (do not exceed risk-free rate or long-term GDP)

**Terminal Value Sanity Check:** Should represent 50-70% of Enterprise Value. If >75%, model may be over-reliant on terminal assumptions.

### Step 9: Enterprise to Equity Value Bridge

```
Sum of PV of Projected FCFs = $X million
PV of Terminal Value = $Y million
= Enterprise Value = $Z million
(-) Net Debt = $A million
= Equity Value = $B million
÷ Diluted Shares Outstanding = C million shares
= Implied Price per Share = $XX.XX
```

### Step 10: Sensitivity Analysis (Bottom of DCF Sheet)

Three sensitivity tables:
1. **WACC vs Terminal Growth** — enterprise value sensitivity
2. **Revenue Growth vs EBIT Margin** — top-line and operating leverage impact
3. **Beta vs Risk-Free Rate** — cost of equity sensitivity

Each is a 5×5 grid, NOT Excel's Data Table feature. Write formulas programmatically with openpyxl loops.

---

## Scenario Block Structure

Create **separate blocks** for Bear/Base/Bull with assumptions laid horizontally across projection years. Each block must have:
1. Section header row (merged)
2. Column header row showing years (REQUIRED)
3. Data rows

Use a **consolidation column** with INDEX formulas:
`=INDEX(B10:D10, 1, $B$6)` where $B$6 = case selector (1=Bear, 2=Base, 3=Bull)

Then reference the consolidation column in projections:
`Revenue Year 1: =D29*(1+$E$10)` — NOT scattered IF statements

---

## Excel Model Structure

**Two sheets:**
1. **DCF** — main model with sensitivity analysis at bottom
2. **WACC** — cost of capital calculation

**Color Convention:**
- Blue font: hardcoded inputs
- Black font: formulas
- Green font: cross-sheet links
- Section headers: Dark blue `#1F4E79`, white bold text
- Key outputs: Medium blue `#BDD7EE`, black bold

**Number Formats:**
- Percentages: `0.0%`
- Currency: `$#,##0` for millions; `$#,##0.00` for per-share
- Negative numbers: `(#,##0)` in parentheses

**File naming:** `[Ticker]_DCF_Model_[Date].xlsx`

---

## Command Workflow: /dcf

Build an institutional-quality DCF model that uses comparable company analysis to inform valuation ranges.

### Step 1: Gather Company Information
If company name/ticker provided, use it. Otherwise ask: "What company would you like to value?"

### Step 2: Run Comparable Company Analysis
Load `comps-analysis` skill to:
- Identify 4-6 comparable public companies
- Pull operating metrics and valuation multiples
- Calculate statistical summary (median, 25th/75th percentiles)

**Use comps to inform DCF assumptions:**
| Comps Output | DCF Input |
|---|---|
| Peer median EV/EBITDA | Terminal exit multiple range |
| Peer 25th-75th EV/EBITDA | Sensitivity analysis range |
| Peer median growth rate | Benchmark for revenue assumptions |
| Peer median EBITDA margin | Target margin in terminal year |

### Step 3: Build DCF Model
Gather historical financials, build Bear/Base/Bull scenarios, model FCF, calculate WACC, discount cash flows, calculate terminal value, bridge to equity value and implied share price.

### Step 4: Cross-Check Valuation
- Implied EV/EBITDA from DCF vs peer median
- Terminal value as % of EV (should be 50-70%)
- Implied growth vs peer growth rates

### Step 5: Deliver Output
1. **Comps analysis spreadsheet** (.xlsx)
2. **DCF model** (.xlsx) with Bear/Base/Bull, sensitivity tables, valuation summary

---

## Before Delivering Model (MANDATORY)

1. Run `python recalc.py model.xlsx 30` — fix ALL errors until status is "success"
2. Verify scenario blocks have column header rows
3. Verify consolidation column updates when case selector changes
4. Verify sensitivity tables fully populated with formulas (no placeholders)
5. Verify font colors: Blue=inputs, Black=formulas, Green=sheet links
6. Verify cell comments on ALL hardcoded inputs
7. Verify professional borders around major sections

## Common Errors to Avoid

1. **Formula row references off** → Define ALL row positions BEFORE writing formulas
2. **Missing cell comments** → Add comments AS cells are created
3. **Simplified sensitivity tables** → Populate all 75 cells with full DCF recalc formulas
4. **Scenario block references wrong** → Use consolidation column with INDEX formulas
5. **OpEx based on gross profit** → ALWAYS use revenue as denominator
6. **Terminal growth > WACC** → Creates infinite value, model is wrong
