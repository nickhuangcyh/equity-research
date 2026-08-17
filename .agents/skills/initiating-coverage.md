---
name: initiating-coverage
description: Create institutional-quality equity research initiation reports through a 5-task workflow — company research, financial modeling, valuation analysis, chart generation, and final report assembly. Each task produces specific deliverables and must be executed individually.
---

# Initiating Coverage

Create institutional-quality equity research initiation reports (30-50 pages) following institutional standards (JPMorgan, Goldman Sachs, Morgan Stanley format).

## ⚠️ CRITICAL: One Task at a Time

**THIS SKILL OPERATES IN SINGLE-TASK MODE ONLY.**

When user requests a full report, ask which specific task to start with. Never chain tasks automatically.

---

## Task Overview

| Task | Name | Prerequisites | Output |
|---|---|---|---|
| **1** | Company Research | Company name/ticker | 6-8K word document |
| **2** | Financial Modeling | 10-K or financials access | Excel model (6 tabs) |
| **3** | Valuation Analysis | Financial model (Task 2) | Valuation + price target |
| **4** | Chart Generation | Tasks 1, 2, 3 + external data | 25-35 PNG/JPG charts |
| **5** | Report Assembly | ALL previous tasks (1-4) | 30-50 page DOCX report |

**Deliverables Policy (NO SHORTCUTS):**
- Task 1: Research document (.md) — NOTHING ELSE
- Task 2: Financial model (.xlsx) — NOTHING ELSE
- Task 3: Valuation analysis (.md) + Excel tabs — NOTHING ELSE
- Task 4: Charts zip file (.zip) — NOTHING ELSE
- Task 5: Final report (.docx) — NOTHING ELSE

---

## Task 1: Company Research

**Prerequisites:** ✅ None — just company name or ticker

**Deliverable:** Company Research Document (6,000-8,000 words)

Contents:
- Company overview & history
- Management bios (300-400 words × 3-4 execs)
- Products & services analysis
- Industry overview
- Competitive analysis (5-10 competitors)
- TAM sizing
- Risk assessment (8-12 risks across 4 categories)

**File name:** `[Company]_Research_Document_[Date].md`

---

## Task 2: Financial Modeling

**Prerequisites:** ⚠️ Access to company financial data (10-K from SEC EDGAR, or user-provided statements)

**Deliverable:** Excel Financial Model (.xlsx) with 6 essential tabs:
1. **Revenue Model** — Product breakdown (20-30 rows) + Geography breakdown (15-20 rows)
2. **Income Statement** — Full P&L with 40-50 line items, historical (3-5 years) + projected (5 years)
3. **Cash Flow Statement** — Operating/Investing/Financing, historical + projected
4. **Balance Sheet** — Assets/Liabilities/Equity, historical + projected
5. **Scenarios** — Bull/Base/Bear comparison table
6. **DCF Inputs** — Prepared for Task 3 valuation

**File name:** `[Company]_Financial_Model_[Date].xlsx`

All formulas must be live Excel formulas — NO hardcoded computed values. Run `recalc.py` before delivery, zero errors required.

---

## Task 3: Valuation Analysis

**Prerequisites:** ⚠️ TASK 2 MUST BE COMPLETE. Do not start without the financial model.

**Deliverable:** Valuation analysis document (.md) + additional Excel tabs added to Task 2 file

Analysis includes:
- DCF valuation (WACC, terminal growth, sensitivity tables)
- Comparable company analysis (5-8 public comps)
- Precedent transaction analysis (if available)
- Price target derivation with methodology
- Bull/Base/Bear implied price targets

---

## Task 4: Chart Generation

**Prerequisites:** ⚠️ TASKS 1, 2, AND 3 MUST BE COMPLETE.

**Deliverable:** 25-35 charts (.zip file) in PNG/JPG format

Chart categories:
- Financial performance charts (revenue, EBITDA, EPS trends)
- Margin analysis charts
- Revenue by segment/geography
- Valuation comparison (trading comps, historical multiples)
- DCF sensitivity heatmaps
- Market share and competitive positioning

---

## Task 5: Report Assembly

**Prerequisites:** ⚠️ ALL PREVIOUS TASKS (1-4) MUST BE COMPLETE.

**Deliverable:** 30-50 page DOCX report

Report structure:
- Page 1: Rating, price target, investment thesis
- Pages 2-3: Executive summary
- Pages 4-9: Company overview
- Pages 10-15: Industry analysis
- Pages 16-20: Financial model summary
- Pages 21-25: Valuation analysis
- Pages 26-30: Investment risks
- Pages 31-40: Detailed financial tables
- Pages 41-50: Appendix (optional)

All sources cited with clickable hyperlinks. Times New Roman throughout.

---

## Key Differences from Earnings Update

| Aspect | Initiation Report | Earnings Update |
|---|---|---|
| Length | 30-50 pages | 8-12 pages |
| Tables | 12-20 comprehensive | 1-3 summary |
| Figures | 25-35 | 8-12 |
| Turnaround | 3-6 weeks | 1-2 days |
| Focus | Everything | What's NEW |
| XLS Model | Required | Optional |

---

## Command Workflow: /initiate

Begin the 5-task workflow to create an institutional-quality initiation report.

If a ticker is provided, use it. Otherwise ask which company to initiate on.

**First response format:**
```
I can create an equity research initiation report for [Company].
This involves 5 separate tasks:

1. Company Research — business, management, industry
2. Financial Modeling — build projection model in Excel
3. Valuation Analysis — DCF and comparable companies
4. Chart Generation — 25-35 charts
5. Report Assembly — compile final 30-50 page report

Which task would you like to start with?
```
