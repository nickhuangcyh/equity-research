---
name: model-update
description: Update financial models with new data — quarterly earnings, management guidance, macro changes, or revised assumptions. Adjusts estimates, recalculates valuation, and flags material changes. Use after earnings, guidance updates, or when assumptions need refreshing.
---

# Model Update

Update financial models with new earnings, guidance, or revised assumptions.

## Workflow

### Step 1: Identify What Changed

Determine the update trigger:
- **Earnings release**: New quarterly actuals to plug in
- **Guidance change**: Company updated forward outlook
- **Estimate revision**: Analyst changing assumptions based on new data
- **Macro update**: Interest rates, FX, commodity prices changed
- **Event-driven**: M&A, restructuring, new product, management change

### Step 2: Plug New Data (After Earnings)

| Line Item | Prior Estimate | Actual | Delta | Notes |
|---|---|---|---|---|
| Revenue | | | | |
| Gross Margin | | | | |
| Operating Expenses | | | | |
| EBITDA | | | | |
| EPS | | | | |
| [Key metric 1] | | | | |

**Segment Detail** (if applicable):
- Update each segment's revenue and margin
- Note any segment mix shifts

**Balance Sheet / Cash Flow Updates:**
- Cash and debt balances
- Share count (buybacks, dilution)
- CapEx actual vs. estimate
- Working capital changes

### Step 3: Revise Forward Estimates

| | Old FY Est | New FY Est | Change | Old Next FY | New Next FY | Change |
|---|---|---|---|---|---|---|
| Revenue | | | | | | |
| EBITDA | | | | | | |
| EPS | | | | | | |

**Key Assumption Changes:**
- Revenue growth rate: old → new (reason)
- Margin assumption: old → new (reason)
- Any new items (restructuring charges, one-time gains, etc.)

### Step 4: Valuation Impact

| Valuation Method | Prior | Updated | Change |
|---|---|---|---|
| DCF fair value | | | |
| P/E (NTM EPS × target multiple) | | | |
| EV/EBITDA (NTM EBITDA × target multiple) | | | |
| **Price Target** | | | |

### Step 5: Summary & Action

**Estimate Change Summary:**
- One paragraph: what changed, why, and what it means for the stock
- Is this a thesis-changing event or noise?

**Rating / Price Target:**
- Maintain or change rating?
- New price target with methodology
- Upside/downside to current price

### Step 6: Output

- Updated Excel model (if user provides existing model)
- Estimate change summary (markdown or Word)
- Updated price target derivation

## Important Notes

- Always reconcile estimates to the company's reported figures before projecting forward
- Note any non-recurring items and whether estimates are GAAP or adjusted
- Track estimate revision history — it shows analytical progression
- Check consensus after updating — how do revised estimates compare to the Street?
- Share count matters — dilution from stock comp, converts, or buybacks can materially affect EPS

---

## Command Workflow: /model-update

Update a financial model with new data.

If a ticker is provided, use it. Otherwise ask the user which model to update and what changed.

Deliver: estimate change summary, updated price target, and updated Excel model (if provided).
