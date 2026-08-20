#!/usr/bin/env python3
"""
evaluate_vitl_math.py - Mathematical validation script for Vital Farms, Inc. (NASDAQ: VITL) DCF Model.
Verifies all cash flows, present values, equity bridges, and sensitivity tables across Bear, Base, and Bull cases.
"""

import numpy as np

# Market Data
stock_price = 11.10
shares_out = 42.94  # M
market_cap = stock_price * shares_out # $M
total_debt = 30.00   # $M
cash_st_inv = 160.00 # $M
net_debt = total_debt - cash_st_inv # -$130.00 M (Net Cash)
current_ev = market_cap + net_debt
tax_rate = 0.250

print(f"Stock Price: ${stock_price:.2f}")
print(f"Shares Out: {shares_out:.2f} M")
print(f"Market Cap: ${market_cap:,.2f} M (${market_cap/1e3:.2f} B)")
print(f"Total Debt: ${total_debt:,.2f} M")
print(f"Cash & ST Inv: ${cash_st_inv:,.2f} M")
print(f"Net Debt / (Net Cash): ${net_debt:,.2f} M")
print(f"Enterprise Value: ${current_ev:,.2f} M (${current_ev/1e3:.2f} B)")

# Historical
hist_rev = [471.86, 606.32, 759.40] # FY23A, FY24A, FY25A
hist_ebit = [33.32, 63.60, 84.20]

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
    [0.120, 0.100, 0.090, 0.080, 0.070],
    [0.360, 0.355, 0.350, 0.350, 0.350],
    [0.095, 0.098, 0.100, 0.100, 0.100],
    [0.0450, 0.0450, 0.0420, 0.0400, 0.0400],
    [0.0240, 0.0240, 0.0240, 0.0230, 0.0230],
    [0.018, 0.018, 0.018, 0.018, 0.018],
    0.0200,
    0.1050
)

# Base Case
p_base, fcff_base, pv_fcff_base, cum_pv_base, pv_tv_base, ev_base, eq_base = run_dcf(
    "Base Case",
    [0.180, 0.150, 0.130, 0.110, 0.090],
    [0.378, 0.380, 0.382, 0.385, 0.385],
    [0.112, 0.118, 0.122, 0.125, 0.125],
    [0.0450, 0.0420, 0.0400, 0.0400, 0.0400],
    [0.0240, 0.0240, 0.0240, 0.0240, 0.0240],
    [0.015, 0.015, 0.015, 0.015, 0.015],
    0.0250,
    0.1000
)

# Bull Case
p_bull, fcff_bull, pv_fcff_bull, cum_pv_bull, pv_tv_bull, ev_bull, eq_bull = run_dcf(
    "Bull Case",
    [0.220, 0.180, 0.160, 0.140, 0.120],
    [0.385, 0.390, 0.395, 0.400, 0.400],
    [0.120, 0.128, 0.135, 0.140, 0.145],
    [0.0420, 0.0400, 0.0380, 0.0380, 0.0380],
    [0.0240, 0.0240, 0.0240, 0.0240, 0.0240],
    [0.012, 0.012, 0.012, 0.012, 0.012],
    0.0300,
    0.0950
)

# Sensitivity Table 1: WACC vs Terminal Growth Rate
print("\n--- Sensitivity Table 1: WACC vs Terminal Growth Rate ($) ---")
term_g_axis = [0.0150, 0.0200, 0.0250, 0.0300, 0.0350]
wacc_axis = [0.0900, 0.0950, 0.1000, 0.1050, 0.1100]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]

header_t1 = "WACC \\ g  " + "".join([f"{g:>10.2%}" for g in term_g_axis])
print(header_t1)
for w in wacc_axis:
    row_str = f"{w:<10.2%}"
    for g in term_g_axis:
        pv_f = sum([f / ((1 + w) ** p) for f, p in zip(fcff_base, periods)])
        t_f = fcff_base[-1] * (1 + g)
        t_v = t_f / (w - g)
        pv_tv = t_v / ((1 + w) ** periods[-1])
        eq_v = pv_f + pv_tv - net_debt
        price = eq_v / shares_out
        row_str += f"{price:>10.2f}"
    print(row_str)

# Sensitivity Table 2: FY26E Rev Growth vs FY30E EBIT Margin
print("\n--- Sensitivity Table 2: FY26E Rev Growth vs FY30E EBIT Margin ($) ---")
ebit_margin_axis = [0.105, 0.115, 0.125, 0.135, 0.145]
rev_growth_axis = [0.120, 0.150, 0.180, 0.210, 0.240]

header_t2 = "Rev G \\ EBIT% " + "".join([f"{m:>10.1%}" for m in ebit_margin_axis])
print(header_t2)
base_rev_g = 0.180
base_ebit_m = 0.125

for rg in rev_growth_axis:
    row_str = f"{rg:<14.1%}"
    for em in ebit_margin_axis:
        rev_scale = (1 + rg) / (1 + base_rev_g)
        ebit_scale = em / base_ebit_m
        
        pv_f_scaled = sum([pv * rev_scale for pv in pv_fcff_base[:-1]]) + pv_fcff_base[-1] * rev_scale * ebit_scale
        pv_tv_scaled = pv_tv_base * rev_scale * ebit_scale
        eq_v = pv_f_scaled + pv_tv_scaled - net_debt
        price = eq_v / shares_out
        row_str += f"{price:>10.2f}"
    print(row_str)

# Sensitivity Table 3: Beta vs Risk-Free Rate
print("\n--- Sensitivity Table 3: Beta vs Risk-Free Rate ($) ---")
rf_axis = [0.0414, 0.0439, 0.0464, 0.0489, 0.0514]
beta_axis = [0.88, 0.98, 1.08, 1.18, 1.28]
erp = 0.0520

header_t3 = "Beta \\ Rf   " + "".join([f"{rf:>10.2%}" for rf in rf_axis])
print(header_t3)
for b in beta_axis:
    row_str = f"{b:<12.2f}"
    for rf in rf_axis:
        w = rf + b * erp
        pv_f = sum([f / ((1 + w) ** p) for f, p in zip(fcff_base, periods)])
        t_f = fcff_base[-1] * (1 + 0.0250)
        t_v = t_f / (w - 0.0250)
        pv_tv = t_v / ((1 + w) ** periods[-1])
        eq_v = pv_f + pv_tv - net_debt
        price = eq_v / shares_out
        row_str += f"{price:>10.2f}"
    print(row_str)
