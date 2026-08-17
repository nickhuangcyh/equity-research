---
name: thesis-tracker
description: Maintain and update investment theses for portfolio positions and watchlist names. Track key data points, catalysts, and thesis milestones over time. Use when updating a thesis with new information, reviewing position rationale, or checking if a thesis is still intact.
---

# Thesis Tracker

Maintain and update investment theses, tracking key data points, catalysts, and milestones over time.

## Workflow

### Step 1: Define or Load Thesis

If creating a new thesis:
- **Company**: Name and ticker
- **Position**: Long or Short
- **Thesis statement**: 1-2 sentence core thesis
- **Key pillars**: 3-5 supporting arguments
- **Key risks**: 3-5 risks that would invalidate the thesis
- **Catalysts**: Upcoming events that could prove/disprove
- **Target price / valuation**: What's it worth if the thesis plays out
- **Stop-loss trigger**: What would make you exit

If updating an existing thesis, ask for the new data point or development.

### Step 2: Update Log

For each new data point or development:

- **Date**: When this happened
- **Data point**: What changed (earnings beat, management departure, competitor move, etc.)
- **Thesis impact**: Does this strengthen, weaken, or neutralize a specific pillar?
- **Action**: No change / Increase position / Trim / Exit
- **Updated conviction**: High / Medium / Low

### Step 3: Thesis Scorecard

Maintain a running scorecard:

| Pillar | Original Expectation | Current Status | Trend |
|---|---|---|---|
| Revenue growth >20% | On track | Q3 was 22% | Stable |
| Margin expansion | Behind | Margins flat YoY | Concerning |
| New product launch | Pending | Delayed to Q2 | Watch |

### Step 4: Catalyst Calendar

Track upcoming catalysts:

| Date | Event | Expected Impact | Notes |
|---|---|---|---|
| | | | |

### Step 5: Output

Thesis summary suitable for morning meeting discussion, portfolio review, or risk committee presentation.

Format: Concise markdown or Word doc with the scorecard, recent updates, and current conviction level.

## Important Notes

- A thesis should be falsifiable — if nothing could disprove it, it's not a thesis
- Track disconfirming evidence as rigorously as confirming evidence
- Review theses at least quarterly
- Store thesis data in a structured format so it can be referenced across sessions

---

## Command Workflow: /thesis

Create a new thesis or update an existing one with new data points.

If a ticker is provided, use it. Otherwise ask the user which position to review.
