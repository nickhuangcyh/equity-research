---
name: earnings-analysis
description: Create professional equity research earnings update reports (8-12 pages) analyzing quarterly results for companies under coverage. Fast-turnaround format covering beat/miss analysis, key metrics, updated estimates, and revised thesis.
---

# Equity Research Earnings Update

Create professional **EARNINGS UPDATE REPORTS** analyzing quarterly results for companies already under coverage, following institutional standards (JPMorgan, Goldman Sachs, Morgan Stanley format).

**Key Characteristics:**
- **Length**: 8-12 pages, 3,000-5,000 words
- **Tables**: 1-3 summary tables (NOT comprehensive)
- **Figures**: 8-12 charts
- **Turnaround**: 1-2 days (within 24-48 hours of earnings)
- **Audience**: Clients already familiar with the company
- **Focus**: What's NEW — beat/miss, updated estimates, thesis impact

## When to Use

- "Create an earnings update for [Company] Q3 2024"
- "Analyze [Company]'s quarterly results"
- "Post-earnings report for [Company]"

**Do NOT use if:**
- User requests "initiation report" → use `initiating-coverage` skill
- Company is not already covered → need initiation first

---

## Critical Requirements

### 1. Speed & Timeliness

- Publish within 24-48 hours of earnings release
- Focus on NEW information only
- Don't rehash company background extensively

**🚨 CRITICAL: TRAINING DATA MAY BE OUTDATED 🚨**

Before starting, complete these steps in order:
1. **Check today's date**
2. **Search for latest**: "[Company] latest earnings results"
3. **Verify date**: Confirm earnings release is within last 3 months
4. **If stale (>3 months)**, search again

### 2. Beat/Miss Analysis

- Lead with whether company beat or missed estimates
- Quantify variances: "Revenue beat by $120M or 3%"
- Explain WHY results differed from expectations

### 3. Citations & Source Attribution (MANDATORY)

Include specific citations with clickable hyperlinks in every figure and table:

```
Source: Q3 2024 10-Q filed November 8, 2024; Company earnings release
        [Hyperlink "10-Q" to: https://www.sec.gov/cgi-bin/viewer?accession=...]
```

**Required Sources:**
- ✅ Earnings release (with date and URL)
- ✅ 10-Q filing (with filing date and EDGAR link)
- ✅ Earnings call transcript (with date)
- ✅ Consensus estimates source (Bloomberg/FactSet with date)

---

## High-Level Workflow

### Phase 1: Data Collection

- Search for latest earnings release (press release)
- Pull 10-Q filing from SEC EDGAR
- Pull earnings call transcript
- Pull consensus estimates (Bloomberg/FactSet)

### Phase 2: Analysis

- Beat/miss analysis for each key metric
- Segment/geographic/product breakdown
- Margin and guidance analysis
- Update financial model and estimates

### Phase 3: Chart Generation (8-12 charts)

Create charts focusing on what's new:
- Quarterly revenue progression
- Quarterly EPS progression
- Quarterly margin trends
- Revenue by segment/geography
- Beat/miss summary
- Estimate revisions

### Phase 4: Report Creation (8-12 pages)

- Page 1: Earnings summary with rating and price target
- Pages 2-3: Detailed results analysis
- Pages 4-5: Key metrics & guidance
- Pages 6-7: Updated investment thesis
- Pages 8-10: Valuation & estimates
- Pages 11-12: Appendix (optional)

### Phase 5: Quality Check

Verify all data is current, all sources are cited with clickable hyperlinks, all charts embedded.

---

## Key Differences from Initiation Report

| Aspect | Earnings Update | Initiation Report |
|---|---|---|
| Length | 8-12 pages | 30-50 pages |
| Tables | 1-3 summary | 12-20 comprehensive |
| Figures | 8-12 | 25-35 |
| Turnaround | 1-2 days | 3-6 weeks |
| Focus | What's NEW | Everything |

---

## Command Workflow: /earnings

Create an earnings update report analyzing quarterly results.

### Step 1: Gather Information
Parse input for: company name/ticker, quarter (e.g., Q3 2024, Q2 FY25)

### Step 2: Verify Timeliness
Search: "[Company] latest earnings results [current year]"
Verify earnings release is within last 3 months.

### Step 3: Collect Data
- Earnings release (press release)
- 10-Q from SEC EDGAR
- Earnings call transcript
- Consensus estimates

### Step 4: Beat/Miss Analysis
- Revenue vs consensus: Beat/Miss by $X or X%
- EPS vs consensus: Beat/Miss by $X
- Key segment performance vs expectations
- Explain WHY results differed

### Step 5: Create Report (8-12 pages)
- Page 1: Summary with rating and price target
- Pages 2-3: Detailed results analysis
- Pages 4-5: Key metrics & guidance
- Pages 6-7: Updated investment thesis
- Pages 8-10: Valuation & estimates
- Sources section with clickable hyperlinks

### Step 6: Deliver Output
1. DOCX report (8-12 pages, 3,000-5,000 words)
2. Summary: beat/miss on key metrics, guidance changes, thesis impact

## Quality Checklist

- [ ] Earnings data is from latest quarter (not stale)
- [ ] Beat/miss quantified with specific numbers
- [ ] 8-12 charts embedded
- [ ] Sources section with clickable hyperlinks
- [ ] Every figure/table has source citation
- [ ] Guidance changes clearly documented
- [ ] Rating and price target stated upfront
