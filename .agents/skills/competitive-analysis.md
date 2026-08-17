---
name: competitive-analysis
description: Build competitive landscape decks — market positioning, competitor deep-dives, comparative analysis, strategic synthesis. Use when asked for a competitive landscape, competitor analysis, peer comparison, market positioning assessment, or to evaluate competitive dynamics across an industry.
---

# Competitive Landscape Mapping

Build a complete competitive analysis deck. This is a two-phase process: gather requirements and get outline approval first, then build.

## Environment Check

This skill works in both the PowerPoint add-in and chat:
- **Add-in** — build slides directly into the open deck
- **Chat** — generate a `.pptx` file (or build into one the user uploaded)

## Phase 1 — Scope the Analysis

Before any research or slide-building, pin down what they actually want. Gather in one round:

- **Scope** — Single target company with competitors? Or multi-company side-by-side?
- **Competitor set** — Which companies are in scope?
- **Audience and depth** — Quick read for someone in the space, or a full primer?
- **Investment context** — Do they need bull/base/bear scenarios and signposts?

If they've uploaded an Excel/CSV with competitor data, confirm which columns map to which metrics before starting.

## Phase 2 — Outline, Approve, Then Build

**Do not create slides until the outline is approved.** Propose slide titles and one-line content notes, present them, get a yes. The outline is the cheap iteration point.

---

## Standards — Apply Throughout

### Prompt Fidelity

- **Slide titles and section names** — exact wording when user specifies
- **Chart vs. table** — not interchangeable
- **Complete data series** — if they list 7 competitors, include all 7
- **Exact values and ratios** — "surpasses DoorDash 4:1" means that ratio

### Data Source Priority

1. 10-Ks / annual reports (audited)
2. Earnings calls / investor presentations
3. Sell-side research (analyst estimates)
4. Industry reports (McKinsey, Gartner)
5. News (recent developments only; verify against primary sources)

### Data Comparability

- All competitor metrics from the same fiscal year; flag exceptions
- Same metric definitions across competitors
- Convert to USD for international comparisons
- Missing data shows as "-" or "N/A" with "[E]" flag for estimates
- Every number has a citation: "[Company] [Document] ([Date])"

### Design

- **Slide titles are insights, not labels.** "Scale leaders pulling away from niche players" — not "Competitive Analysis."
- **Signposts are quantified.** "Margin below 40%" — not "margins decline."
- **Typography:** Titles 28-32pt bold, body 14-16pt (never below 14pt), same element type = same size throughout
- **Color:** 2-3 colors max. Muted — navy, gray, one accent

---

## Analysis Workflow

### Step 0 — Industry-Defining Metrics

Identify 3-5 metrics this industry actually runs on:

| Industry | Key Metrics |
|---|---|
| SaaS | ARR, NRR, CAC payback, LTV/CAC, Rule of 40 |
| Payments | GPV, take rate, attach rate, transaction margin |
| Marketplaces | GMV, take rate, buyer/seller ratio, repeat rate |
| Retail | Same-store sales, inventory turns, sales per sq ft |
| Logistics | Volume, cost per unit, on-time delivery %, capacity utilization |

### Step 1 — Market Context

Size, growth, drivers, headwinds. With sources.

Correct: "Embedded payments is $80-100B in 2024, growing 20-25% CAGR (McKinsey 2024)"
Wrong: "The market is large and growing rapidly"

### Step 2 — Industry Economics

Map how value flows. Approach depends on industry structure:
- **Vertically structured** — value chain layers, typical margin at each
- **Platform/network** — ecosystem participants, value flows between them
- **Fragmented** — consolidation dynamics, margin differences by scale

### Step 3 — Target Company Profile

```
| Metric | Value |
|---|---|
| Revenue | $4.96B |
| Growth | +26% YoY |
| Gross Margin | 45% |
| Profitability | $373M Adj. EBITDA |
| Customers | 134K |
| Retention | 92% |
| Market Share | ~15% |
```

### Step 4 — Competitor Mapping

Group by whichever lens fits:
- By business model — platform / vertical / horizontal
- By segment — enterprise / SMB / consumer
- By posture — direct / adjacent / emerging

### Step 5 — Positioning Visualization

| Type | When |
|---|---|
| 2×2 matrix | Two dominant competitive factors |
| Radar/spider | Multi-factor comparison |
| Tier diagram | Natural clustering into strategic groups |
| Value chain map | Vertical industries |
| Ecosystem map | Platform markets |

### Step 6 — Competitor Deep-Dives

Two tables per competitor:

**Metrics:**
```
| Metric | Value |
|---|---|
| Revenue | $X.XB |
| Growth | +XX% YoY |
| Gross Margin | XX% |
| Market Cap | $X.XB |
```

**Qualitative:**
```
| Category | Assessment |
|---|---|
| Business | What they do (1 sentence) |
| Strengths | 2-3 bullets |
| Weaknesses | 2-3 bullets |
| Strategy | Current priorities |
```

### Step 7 — Comparative Analysis

```
| Dimension | Company A | Company B | Company C |
|---|---|---|---|
| Scale | ●●● $160B | ●●○ $45B | ●○○ $8B |
| Growth | ●●○ +26% | ●●● +35% | ●●○ +22% |
| Margins | ●●○ 7.5% | ●○○ 3.2% | ●●● 15% |
```

### Step 8 — Moat Assessment

Rate each competitor **Strong / Moderate / Weak** on:

| Moat | What to Assess |
|---|---|
| Network effects | User/supplier flywheel strength |
| Switching costs | Technical integration depth, contractual lock-in |
| Scale economies | Unit cost advantages at volume |
| Intangible assets | Brand, proprietary data, regulatory licenses, patents |

### Step 9 — Synthesis (For Investment Contexts)

```
| Scenario | Probability | Key driver |
|---|---|---|
| Bull | 30% | Market share gains, margin expansion |
| Base | 50% | Current trajectory continues |
| Bear | 20% | Competitive pressure, margin compression |
```

---

## Quality Checklist

- Slide titles match what user specified, verbatim
- Every competitor/year/data point they listed is present
- Every number has a citation
- All metrics from same fiscal period (or flagged)
- Titles state insights, not topics
- Charts are real chart objects (not text tables)
- No overlapping elements, text within containers

---

## Command Workflow: /competitive-analysis

Build a competitive landscape analysis for the specified company or industry.

If a company/industry is provided, use it. Otherwise ask what they want to analyze.

Then follow Phase 1 (scope) → get outline approval → Phase 2 (build).
