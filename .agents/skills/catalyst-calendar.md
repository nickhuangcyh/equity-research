---
name: catalyst-calendar
description: Build and maintain a calendar of upcoming catalysts across a coverage universe — earnings dates, conferences, product launches, regulatory decisions, and macro events. Helps prioritize attention and position ahead of events.
---

# Catalyst Calendar

Build and maintain a calendar of upcoming catalysts across the coverage universe.

## Workflow

### Step 1: Define Coverage Universe

- List of companies to track (tickers or names)
- Sector / industry focus
- Include macro events? (Fed meetings, economic data)
- Time horizon (next 2 weeks, month, quarter)

### Step 2: Gather Catalysts

For each company, identify upcoming events:

**Earnings & Financial Events**
- Quarterly earnings date and time (pre/post market)
- Annual shareholder meeting
- Investor day / analyst day
- Capital markets day
- Debt maturity / refinancing dates

**Corporate Events**
- Product launches or announcements
- FDA approvals / regulatory decisions
- Contract renewals or expirations
- M&A milestones (close dates, regulatory approvals)
- Management transitions
- Insider trading windows (lockup expirations)

**Industry Events**
- Major conferences (dates, which companies presenting)
- Trade shows and expos
- Industry data releases (monthly sales, traffic, etc.)

**Macro Events**
- Fed meetings (FOMC dates)
- Jobs report, CPI, GDP releases
- Geopolitical events with market impact

### Step 3: Calendar View

| Date | Event | Company/Sector | Type | Impact (H/M/L) | Our Positioning | Notes |
|---|---|---|---|---|---|---|

### Step 4: Weekly Preview

**This Week's Key Events:**
1. [Day]: [Company] Q[X] earnings — consensus [$X EPS], our estimate [$X], key focus: [metric]
2. [Day]: [Event] — why it matters for [stocks]
3. [Day]: [Macro release] — expectations and positioning

**Position Implications:**
- Events that could move specific positions
- Any pre-positioning recommended
- Risk management ahead of binary events

### Step 5: Output

- Excel workbook with calendar view and sortable columns
- Weekly preview note (markdown)

## Important Notes

- Earnings dates shift — verify against company IR pages closer to the date
- Pre-announce risk: track companies with a history of pre-announcing
- Conference attendance lists are valuable — which companies are presenting vs. conspicuously absent
- Color-code by impact level: Red = high, Yellow = moderate, Green = routine
- Archive past catalysts with the actual outcome — builds pattern recognition

---

## Command Workflow: /catalysts

Build or review the upcoming catalyst calendar.

If a timeframe is provided, use it. Otherwise default to the next 2 weeks.

Deliver: calendar view (Excel), weekly preview note with position implications.
