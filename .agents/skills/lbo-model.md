---
name: lbo-model
description: Build or complete LBO (Leveraged Buyout) model templates in Excel for private equity transactions. Fills in formulas, validates calculations, and ensures professional formatting. Use for PE acquisition analysis, deal materials, or IC presentations.
---

# LBO Model Builder

## TEMPLATE REQUIREMENT

**Always check for an attached template file first.**

1. **If a template file is attached/provided**: Use that template's structure exactly
2. **If no template**: Ask: *"Do you have a specific LBO template? If not, I can use the standard template which includes Sources & Uses, Operating Model, Debt Schedule, and Returns Analysis."*
3. **If using standard template**: Copy `examples/LBO_Model.xlsx` as starting point

---

## CRITICAL INSTRUCTIONS — READ FIRST

### Environment: Office JS vs Python

**Office JS (live Excel):**
- Use `Excel.run(async (context) => {...})`
- Write formulas via `range.formulas = [["=B5*B6"]]`
- **Merged cell pitfall:** Write value to top-left cell alone, THEN merge + format:
  ```js
  ws.getRange("A7").values = [["SOURCES & USES"]];
  ws.getRange("A7:F7").merge();
  ws.getRange("A7:F7").format.fill.color = "#1F4E79";
  ```

**Standalone .xlsx (Python/openpyxl):**
- Write formula strings: `ws["D20"] = "=B5*B6"`
- Run `recalc.py` before delivery

### Core Principles
- **Every calculation must be an Excel formula** — NEVER compute in Python and hardcode results
- **Use the template structure** — follow `examples/LBO_Model.xlsx` or user's template
- **Use proper cell references** — never type numbers that should come from other cells
- **Maintain sign convention consistency** throughout
- **Work section by section, verify with user at each step** — STOP after each section, show what was built, get sign-off

### Formula Color Conventions
- **Blue (0000FF)**: Hardcoded inputs
- **Black (000000)**: Formulas with calculations
- **Purple (800080)**: Links to cells on same tab
- **Green (008000)**: Links to cells on different tabs

### Fill Color Palette (Default — override with user/template preference)
- **Section headers**: Dark blue `#1F4E79`, white bold text
- **Column headers**: Light blue `#D9E1F2`, black bold text
- **Input cells**: Light grey `#F2F2F2` — blue font is the signal
- **Formula/calculated cells**: White, no fill
- **Key outputs** (IRR, MOIC, Exit Equity): Medium blue `#BDD7EE`, black bold

---

## TEMPLATE ANALYSIS PHASE — DO THIS FIRST

1. **Map the structure** — identify where each section lives
2. **Understand the timeline** — which columns are which periods
3. **Identify input vs formula cells** — respect template color coding
4. **Read existing labels carefully** — they tell you what calculation is expected
5. **Check for existing formulas** — don't overwrite working formulas
6. **Note template-specific conventions** — sign conventions, subtotal structures

---

## FILLING FORMULAS — GENERAL APPROACH

Hierarchy for each formula cell:
1. **Check the Template** — existing formula? Pattern from neighbors?
2. **Check User's Instructions** — specified calculation method?
3. **Apply Standard Practice** — LBO modeling conventions; document assumptions

---

## COMMON PROBLEM AREAS

### Balancing Sections
When two sections must equal (Sources = Uses), one item is the "plug" — calculate as the difference.

### Tax Calculations
Tax formulas should only reference relevant income line and tax rate. Consider whether losses create tax shields.

### Interest and Circular References
**Use Beginning Balance** (not average or ending) to break circular references.

### Debt Paydown / Cash Sweeps
Respect priority waterfall. Balances cannot go negative — use MAX or MIN functions.

### Returns Calculations (IRR/MOIC)
Cash flows must have correct signs: Investment = negative, Proceeds = positive.

### Sensitivity Tables
- **Use ODD dimensions** (5×5 or 7×7) — guarantees a true center cell
- **Center cell = base case** — axis values symmetric around actual model assumptions
- **Highlight center cell** — medium-blue fill `#BDD7EE` + bold font
- Center cell's IRR/MOIC MUST equal the model's actual output

---

## SECTION-BY-SECTION CHECKPOINTS

STOP and get user sign-off after:
1. **Sources & Uses** → balanced table, plug is correct
2. **Operating Model/Projections** → P&L growth rates and margins look right
3. **Debt Schedule** → beginning/ending balances and interest, waterfall logic
4. **Returns (IRR/MOIC)** → cash flow series and outputs, signs and ranges correct
5. **Sensitivity Tables** → each cell varies, base case lands where expected

---

## VERIFICATION CHECKLIST — RUN AFTER COMPLETION

```bash
python /mnt/skills/public/xlsx/recalc.py model.xlsx
```
Must return success with zero errors.

### Section Balancing
- [ ] Sections that must balance (Sources/Uses, Assets/Liabilities) balance exactly
- [ ] Plug items calculated correctly

### Income/Operating Projections
- [ ] Revenue builds correctly from drivers or growth rates
- [ ] Subtotals and totals sum correctly
- [ ] Margins and ratios are reasonable

### Balance Sheet (if applicable)
- [ ] Assets = Liabilities + Equity
- [ ] Beginning balances = prior period ending balances

### Cash Flow (if applicable)
- [ ] Ending Cash = Beginning Cash + Net Cash Flow
- [ ] Cash balances consistent across statements

### Debt/Financing Schedules (if applicable)
- [ ] Interest calculated on beginning balance
- [ ] Ending balances cannot be negative
- [ ] Paydowns respect cash availability and priority waterfall

### Returns/Output Analysis
- [ ] Cash flow signs correct (negative for investment, positive for proceeds)
- [ ] IRR/MOIC formulas reference complete ranges

### Sensitivity Tables
- [ ] Grid dimensions are ODD (5×5 or 7×7)
- [ ] Row/column axis values symmetric around base case
- [ ] Center cell output equals model's actual IRR/MOIC
- [ ] Center cell highlighted (medium-blue `#BDD7EE`, bold)
- [ ] Each data cell shows a DIFFERENT value
- [ ] Values move in expected directions (higher exit multiple → higher IRR)

### Formatting
- [ ] Hardcoded inputs: blue (0000FF)
- [ ] Calculated formulas: black (000000)
- [ ] Same-tab links: purple (800080)
- [ ] Cross-tab links: green (008000)
- [ ] No error values (#REF!, #DIV/0!, #VALUE!, #NAME?)

---

## COMMON ERRORS TO AVOID

| Error | What Goes Wrong | How to Fix |
|---|---|---|
| Hardcoding calculated values | Model doesn't update | Always use formulas referencing source cells |
| Circular reference errors | Model can't calculate | Use beginning balances for interest calcs |
| Sections don't balance | Totals that should match don't | Ensure one item is the plug |
| Negative balances where impossible | Using more than available | Use `MAX(0, ...)` or `MIN` functions |
| IRR/return errors | Wrong signs or incomplete ranges | Check CF signs, ensure formula covers all periods |
| Sensitivity table shows same value | Formula not varying | Check cell references — need mixed refs (`$A5`, `B$4`) |
| Inconsistent sign conventions | Additions become subtractions | Follow template convention consistently |

---

## Command Workflow: /lbo

Build an LBO model for a PE acquisition.

If a company name is provided, use it. Otherwise ask for the target company and deal parameters.

Key inputs to gather:
- Target company and deal size
- Entry multiple (EV/EBITDA)
- Revenue and EBITDA projections
- Leverage structure (debt tranches, rates)
- Hold period and exit assumptions
- Management rollover

Build: Sources & Uses → Operating Model → Debt Schedule → Returns Analysis → Sensitivity Tables (IRR vs Entry Multiple/Exit Multiple, MOIC vs Leverage/Growth)
