#!/usr/bin/env python3
"""
evaluate_coco_math.py - Mathematical validation script for The Vita Coco Company (NASDAQ: COCO) DCF Model.
Verifies all cash flows, present values, equity bridges, and sensitivity tables across Bear, Base, and Bull cases.
"""

import numpy as np

# Market Data
stock_price = 63.50
shares_out = 57.86  # M
market_cap = stock_price * shares_out # $M
total_debt = 0.00   # $M
cash_st_inv = 279.00 # $M
net_debt = total_debt - cash_st_inv # -$279.00 M (Net Cash)
current_ev = market_cap + net_debt
tax_rate = 0.225

print(f"Stock Price: ${stock_price:.2f}")
print(f"Shares Out: {shares_out:.2f} M")
print(f"Market Cap: ${market_cap:,.2f} M (${market_cap/1e3:.2f} B)")
print(f"Total Debt: ${total_debt:,.2f} M")
print(f"Cash & ST Inv: ${cash_st_inv:,.2f} M")
print(f"Net Debt / (Net Cash): ${net_debt:,.2f} M")
print(f"Enterprise Value: ${current_ev:,.2f} M (${current_ev/1e3:.2f} B)")

# Historical
hist_rev = [493.61, 516.01, 609.80] # FY23A, FY24A, FY25A
hist_ebit = [56.49, 73.82, 82.50]

def run_dcf(case_name, rev_growth, gross_margin, ebit_margin, capex_pct, dna_pct, nwc_pct, term_g, wacc):
    revs = []
    prior_r = hist_rev[-1]
    for g in rev_growth:
        r = prior_r * (1 + g)
        revs.append(r)
        prior_r = r
    
    ebits = [r * m for r, m in zip(revs, ebit_margin)]
    nopats = [eb * (1 - tax_rate) for eb in ebits]
    dnas = [r * dna for r, dna in zip(revs, dna_pct)]
    capexs = [r * cx for r, cx in zip(revs, capex_pct)]
    
    nwcs = []
    prior_r = hist_rev[-1]
    for r, nwc in zip(revs, nwc_pct):
        nwcs.append((r - prior_r) * nwc)
        prior_r = r
        
    fcffs = [np_ + d - cx - nwc for np_, d, cx, nwc in zip(nopats, dnas, capexs, nwcs)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    dfs = [1 / ((1 + wacc) ** p) for p in periods]
    pv_fcffs = [f * d for f, d in zip(fcffs, dfs)]
    cum_pv_fcff = sum(pv_fcffs)
    
    term_fcff = fcffs[-1] * (1 + term_g)
    term_val = term_fcff / (wacc - term_g)
    pv_term_val = term_val / ((1 + wacc) ** periods[-1])
    
    implied_ev = cum_pv_fcff + pv_term_val
    implied_eq_val = implied_ev - net_debt # Since net debt is negative, this adds cash
    implied_share_price = implied_eq_val / shares_out
    upside = (implied_share_price / stock_price - 1) * 100
    tv_pct_ev = (pv_term_val / implied_ev) * 100
    
    print(f"\n==================== {case_name.upper()} ====================")
    print(f"Revenue FY26E-FY30E ($M): {[f'{r:,.1f}' for r in revs]}")
    print(f"EBIT FY26E-FY30E ($M):    {[f'{e:,.1f}' for e in ebits]}")
    print(f"NOPAT FY26E-FY30E ($M):   {[f'{np_:,.1f}' for np_ in nopats]}")
    print(f"D&A FY26E-FY30E ($M):     {[f'{d:,.1f}' for d in dnas]}")
    print(f"CapEx FY26E-FY30E ($M):   {[f'{c:,.1f}' for c in capexs]}")
    print(f"ΔNWC FY26E-FY30E ($M):    {[f'{n:,.1f}' for n in nwcs]}")
    print(f"FCFF FY26E-FY30E ($M):    {[f'{f:,.1f}' for f in fcffs]}")
    print(f"PV of FCFFs ($M):         {[f'{pv:,.1f}' for pv in pv_fcffs]}")
    print(f"Cum PV Explicit FCFF:     ${cum_pv_fcff:,.2f} M")
    print(f"Terminal Year FCFF:       ${term_fcff:,.2f} M")
    print(f"Terminal Value @ FY30E:   ${term_val:,.2f} M")
    print(f"PV of Terminal Value:     ${pv_term_val:,.2f} M ({tv_pct_ev:.1f}% of EV)")
    print(f"Implied Enterprise Value: ${implied_ev:,.2f} M (${implied_ev/1e3:.2f} B)")
    print(f"Implied Equity Value:     ${implied_eq_val:,.2f} M (${implied_eq_val/1e3:.2f} B)")
    print(f"Implied Share Price:      ${implied_share_price:.2f}")
    print(f"Current Stock Price:      ${stock_price:.2f}")
    print(f"Implied Upside/(Downside):{upside:+.1f}%")
    print(f"Implied EV / FY26E EBIT:  {implied_ev / ebits[0]:.1f}x")
    return implied_share_price, fcffs, pv_fcffs, cum_pv_fcff, pv_term_val, implied_ev, implied_eq_val

# Bear Case
p_bear, fcff_bear, pv_fcff_bear, cum_pv_bear, pv_tv_bear, ev_bear, eq_bear = run_dcf(
    "Bear Case",
    [0.260, 0.100, 0.070, 0.050, 0.040],
    [0.380, 0.375, 0.370, 0.365, 0.360],
    [0.140, 0.135, 0.130, 0.125, 0.120],
    [0.0050, 0.0050, 0.0045, 0.0045, 0.0040],
    [0.0020, 0.0020, 0.0020, 0.0020, 0.0020],
    [0.030, 0.030, 0.030, 0.030, 0.030],
    0.0225,
    0.0925
)

# Base Case
p_base, fcff_base, pv_fcff_base, cum_pv_base, pv_tv_base, ev_base, eq_base = run_dcf(
    "Base Case",
    [0.305, 0.160, 0.120, 0.090, 0.060],
    [0.400, 0.405, 0.410, 0.410, 0.410],
    [0.155, 0.160, 0.165, 0.165, 0.160],
    [0.0040, 0.0040, 0.0040, 0.0040, 0.0035],
    [0.0020, 0.0020, 0.0020, 0.0020, 0.0020],
    [0.025, 0.025, 0.025, 0.025, 0.025],
    0.0275,
    0.0875
)

# Bull Case
p_bull, fcff_bull, pv_fcff_bull, cum_pv_bull, pv_tv_bull, ev_bull, eq_bull = run_dcf(
    "Bull Case",
    [0.330, 0.200, 0.150, 0.110, 0.080],
    [0.410, 0.420, 0.425, 0.430, 0.430],
    [0.165, 0.175, 0.180, 0.185, 0.180],
    [0.0035, 0.0035, 0.0035, 0.0030, 0.0030],
    [0.0020, 0.0020, 0.0020, 0.0020, 0.0020],
    [0.020, 0.020, 0.020, 0.020, 0.020],
    0.0325,
    0.0825
)

# Sensitivity Tables Evaluation
print("\n" + "="*50)
print("SENSITIVITY TABLE 1: WACC vs Terminal Growth")
print("="*50)
term_g_axis = [0.0200, 0.0250, 0.0275, 0.0300, 0.0350]
wacc_axis = [0.0775, 0.0825, 0.0875, 0.0925, 0.0975]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]

print("WACC \\ g\t" + "\t".join(f"{g*100:.2f}%" for g in term_g_axis))
for w in wacc_axis:
    row_str = f"{w*100:.2f}%\t"
    for g in term_g_axis:
        pv_fcff_w = sum(f / ((1 + w)**p) for f, p in zip(fcff_base, periods))
        tv_w = (fcff_base[-1] * (1 + g)) / (w - g)
        pv_tv_w = tv_w / ((1 + w)**periods[-1])
        ev_w = pv_fcff_w + pv_tv_w
        eq_w = ev_w - net_debt
        price_w = eq_w / shares_out
        row_str += f"${price_w:.2f}\t"
    print(row_str)

print("\n" + "="*50)
print("SENSITIVITY TABLE 2: FY26E Rev Growth vs FY30E EBIT Margin")
print("="*50)
ebit_margin_axis = [0.140, 0.150, 0.160, 0.170, 0.180]
rev_growth_axis = [0.240, 0.270, 0.305, 0.340, 0.370]
base_rev_g = 0.305
base_ebit_m = 0.160

print("Rev \\ EBIT\t" + "\t".join(f"{m*100:.1f}%" for m in ebit_margin_axis))
for rg in rev_growth_axis:
    row_str = f"{rg*100:.1f}%\t"
    for em in ebit_margin_axis:
        scale_rev = (1 + rg) / (1 + base_rev_g)
        scale_ebit = em / base_ebit_m
        pv_fcff_scaled = sum(pv * scale_rev for pv in pv_fcff_base[:4]) + pv_fcff_base[4] * scale_rev * scale_ebit
        pv_tv_scaled = pv_tv_base * scale_rev * scale_ebit
        ev_s = pv_fcff_scaled + pv_tv_scaled
        eq_s = ev_s - net_debt
        price_s = eq_s / shares_out
        row_str += f"${price_s:.2f}\t"
    print(row_str)

print("\n" + "="*50)
print("SENSITIVITY TABLE 3: Beta vs Risk-Free Rate")
print("="*50)
rf_axis = [0.0413, 0.0438, 0.0463, 0.0488, 0.0513]
beta_axis = [0.65, 0.70, 0.75, 0.80, 0.85]

print("Beta \\ Rf\t" + "\t".join(f"{rf*100:.2f}%" for rf in rf_axis))
for beta in beta_axis:
    row_str = f"{beta:.2f}\t"
    for rf in rf_axis:
        w = rf + beta * 0.0550 # Since net cash, WACC = Ke
        pv_fcff_w = sum(f / ((1 + w)**p) for f, p in zip(fcff_base, periods))
        tv_w = (fcff_base[-1] * (1 + 0.0275)) / (w - 0.0275)
        pv_tv_w = tv_w / ((1 + w)**periods[-1])
        ev_w = pv_fcff_w + pv_tv_w
        eq_w = ev_w - net_debt
        price_w = eq_w / shares_out
        row_str += f"${price_w:.2f}\t"
    print(row_str)
