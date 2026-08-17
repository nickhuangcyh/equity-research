---
name: comps-analysis
description: Build institutional-grade comparable company analyses with operating metrics, valuation multiples, and statistical benchmarking in Excel/spreadsheet format.
---

# Comparable Company Analysis

## ⚠️ CRITICAL: Data Source Priority (READ FIRST)

**ALWAYS follow this data source hierarchy:**

1. **FIRST: Check for MCP data sources** - If S&P Kensho MCP, FactSet MCP, or Daloopa MCP are available, use them exclusively for financial and trading information
2. **DO NOT use web search** if the above MCP data sources are available
3. **ONLY if MCPs are unavailable:** Then use Bloomberg Terminal, SEC EDGAR filings, or other institutional sources
4. **NEVER use web search as a primary data source** - it lacks the accuracy, audit trails, and reliability required for institutional-grade analysis

---

## Overview
This skill teaches Claude to build institutional-grade comparable company analyses that combine operating metrics, valuation multiples, and statistical benchmarking. The output is a structured Excel/spreadsheet that enables informed investment decisions through peer comparison.

**Reference Material & Contextualization:**

When using example files in this skill directory, use them intelligently:

**DO use examples for:**
- Understanding structural hierarchy (how sections flow)
- Grasping the level of rigor expected (statistical depth, documentation standards)
- Learning principles (clear headers, transparent formulas, audit trails)

**DO NOT use examples for:**
- Exact reproduction of format or metrics
- Copying layout without considering context
- Applying the same visual style regardless of audience

**ALWAYS ask yourself first:**
1. **"Do you have a preferred format or should I adapt the template style?"**
2. **"Who is the audience?"** (Investment committee, board presentation, quick reference, detailed memo)
3. **"What's the key question?"** (Valuation, growth analysis, competitive positioning, efficiency)
4. **"What's the context?"** (M&A evaluation, investment decision, sector benchmarking, performance review)

**Core principle:** Use template principles (clear structure, statistical rigor, transparent formulas) but vary execution based on context.

## Core Philosophy
**"Build the right structure first, then let the data tell the story."**

---

## ⚠️ CRITICAL: Formulas Over Hardcodes + Step-by-Step Verification

**Environment — Office JS vs Python:**
- **If running inside Excel (Office Add-in / Office JS):** Use Office JS directly (`Excel.run(async (context) => {...})`). Write formulas via `range.formulas = [["=E7/C7"]]`, not `range.values`.
- **If generating a standalone .xlsx file:** Use Python/openpyxl. Write `cell.value = "=E7/C7"` (formula string).

**Formulas, not hardcodes:**
- Every derived value (margin, multiple, statistic) MUST be an Excel formula — never a pre-computed number
- The only hardcoded values should be raw input data (revenue, EBITDA, share price, etc.)

**Verify step-by-step with the user:**
- After setting up the structure → show header layout before filling data
- After entering raw inputs → confirm sources/periods before building formulas
- After building operating metrics → show calculated margins and sanity-check
- After building valuation multiples → show multiples and confirm reasonableness

---

## Section 1: Document Structure & Setup

### Header Block (Rows 1-3)
```
Row 1: [ANALYSIS TITLE] - COMPARABLE COMPANY ANALYSIS
Row 2: [List of Companies with Tickers]
Row 3: As of [Period] | All figures in [USD Millions/Billions] except per-share amounts and ratios
```

### Visual Convention Standards

**Default Color & Shading — Professional Blue/Grey Palette:**
- **Section headers**: Dark blue `#1F4E79` background, white bold text
- **Column headers**: Light blue `#D9E1F2` background, black bold text
- **Data rows**: White background, black text for formulas, blue text for hardcoded inputs
- **Statistics rows**: Light grey `#F2F2F2` background

**Formatting Conventions:**
- Percentages: 1 decimal (12.3%)
- Multiples: 1 decimal (13.5x)
- Dollar amounts: No decimals, thousands separator (69,632)

---

## Section 2: Operating Statistics & Financial Metrics

### Core Columns
1. **Company** - Names with consistent formatting
2. **Revenue** - Size metric (LTM, quarterly, or annual)
3. **Revenue Growth** - Year-over-year percentage change
4. **Gross Profit** - Revenue minus cost of goods sold
5. **Gross Margin** - GP/Revenue
6. **EBITDA** - Earnings before interest, tax, depreciation, amortization
7. **EBITDA Margin** - EBITDA/Revenue

### Formula Examples (Row 7 as example)
```excel
Gross Margin (F7): =E7/C7
EBITDA Margin (H7): =G7/C7
```

### Statistics Block (After company data)
```
[Leave one blank row for visual separation]
- Maximum: =MAX(B7:B9)
- 75th Percentile: =QUARTILE(B7:B9,3)
- Median: =MEDIAN(B7:B9)
- 25th Percentile: =QUARTILE(B7:B9,1)
- Minimum: =MIN(B7:B9)
```

**Columns that NEED statistics:** Revenue Growth %, Gross Margin %, EBITDA Margin %, EV/Revenue, EV/EBITDA, P/E

**Columns that DON'T need statistics:** Revenue, EBITDA, Market Cap, Enterprise Value (absolute size not comparable)

---

## Section 3: Valuation Multiples & Investment Metrics

### Core Valuation Columns
1. **Company**
2. **Market Cap** - Current market valuation
3. **Enterprise Value** - Market Cap ± Net Debt/Cash
4. **EV/Revenue** - How much market pays per dollar of sales
5. **EV/EBITDA** - How much market pays per dollar of earnings
6. **P/E Ratio** - Price relative to net earnings

### Formula Examples
```excel
EV/Revenue: =[Enterprise Value]/[LTM Revenue]
EV/EBITDA: =[Enterprise Value]/[LTM EBITDA]
P/E Ratio: =[Market Cap]/[Net Income]
```

**Cross-Reference Rule:** Valuation multiples MUST reference the operating metrics section. Never input the same raw data twice.

---

## Section 4: Notes & Methodology Documentation

**Required Components:**
- Data Sources & Quality (MCP, Bloomberg, SEC filings)
- Key Definitions (EBITDA calculation method, FCF formula)
- Valuation Methodology (how EV was calculated, growth rates used)

---

## Section 5: Choosing the Right Metrics

### Start with "What question am I answering?"

**"Which company is undervalued?"** → Focus on EV/Revenue, EV/EBITDA, P/E

**"Which company is most efficient?"** → Focus on Gross Margin, EBITDA Margin, FCF Margin

**"Which company is growing fastest?"** → Focus on Revenue Growth %, EBITDA CAGR

### Industry-Specific Metric Selection

| Industry | Must Have | Optional | Skip |
|---|---|---|---|
| Software/SaaS | Revenue Growth, Gross Margin, Rule of 40 | ARR, Net Dollar Retention | Asset Turnover |
| Manufacturing | EBITDA Margin, Asset Turnover, CapEx/Revenue | ROA, Inventory Turns | Rule of 40 |
| Financials | ROE, ROA, Efficiency Ratio, P/E | Net Interest Margin | Gross Margin, EBITDA |

### The "5-10 Rule"
**5 operating metrics** + **5 valuation metrics** = 10 total columns. If you have more than 15, edit ruthlessly.

---

## Section 6: Best Practices & Quality Checks

### Common Mistakes to Avoid
❌ Mixing market cap and enterprise value in formulas
❌ Using different time periods for numerator and denominator
❌ Hardcoding numbers into formulas instead of cell references
❌ Hard-coded inputs without cell comments citing the source
❌ Including non-comparable companies (different business models)

### Sanity Checks
- **Margin test**: Gross margin > EBITDA margin > Net margin
- **Multiple reasonableness**: EV/Revenue: 0.5-20x, EV/EBITDA: 8-25x, P/E: 10-50x
- **Growth-multiple correlation**: Higher growth usually means higher multiples

---

## Command Workflow: /comps

Build an institutional-grade comparable company analysis with operating metrics, valuation multiples, and statistical benchmarking.

### Step 1: Gather Company Information
If a company name or ticker is provided, use it. Otherwise ask: "What company would you like to analyze?"

### Step 2: Clarify the Analysis Purpose
1. "What's the key question?" (valuation, efficiency, growth comparison)
2. "Who is the audience?" (IC, board, quick reference)
3. "Do you have a preferred format or template?"

### Step 3: Identify Peer Group (4-6 companies)
- Similar business model
- Similar scale/market cap range
- Same industry/sector

### Step 4: Gather Data (prioritize MCP sources if available)
- Operating metrics: Revenue, Growth, Gross Margin, EBITDA, EBITDA Margin
- Valuation: Market Cap, Enterprise Value, EV/Revenue, EV/EBITDA, P/E
- Additional metrics based on industry

### Step 5: Build the Analysis
- Operating Statistics section with company data + statistics
- Valuation Multiples section with same statistical summary
- Notes & Methodology documentation

### Step 6: Deliver Output
1. **Excel file** (.xlsx) - the comps analysis
2. **Summary** highlighting peer group rationale, key insights (premium/discount), median multiples

## Output Checklist

Before delivering, verify:
- [ ] All companies are truly comparable
- [ ] Data is from consistent time periods
- [ ] Units are clearly labeled (millions/billions)
- [ ] Formulas reference cells, not hardcoded values
- [ ] All hard-coded input cells have comments with source citations
- [ ] Statistics include at least Max, 75th, Median, 25th, Min
- [ ] Notes section documents sources and methodology
- [ ] Date stamp is current
- [ ] Formula auditing shows no errors (#DIV/0!, #REF!, #N/A)
