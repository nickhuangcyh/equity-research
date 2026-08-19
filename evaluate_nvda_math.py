import numpy as np

# NVDA Inputs
stock_price = 226.55
shares_out = 24300.0  # M
market_cap = stock_price * shares_out # $M
total_debt = 8500.0   # $M
cash_st_inv = 62600.0 # $M
net_debt = total_debt - cash_st_inv # -$54,100 M (Net Cash)
current_ev = market_cap + net_debt
tax_rate = 0.145

print(f"Market Cap: ${market_cap:,.2f} M (${market_cap/1e6:.2f} T)")
print(f"Net Cash: ${-net_debt:,.2f} M (${-net_debt/1e3:.2f} B)")
print(f"Enterprise Value: ${current_ev:,.2f} M (${current_ev/1e6:.2f} T)")

# Historical
hist_rev = [60922.0, 130497.0, 215938.0]
hist_ebit = [32972.0, 81453.0, 130387.0]

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
    implied_eq_val = implied_ev - net_debt # since net_debt is negative, EV + Net Cash
    implied_share_price = implied_eq_val / shares_out
    upside = (implied_share_price / stock_price - 1) * 100
    tv_pct_ev = (pv_term_val / implied_ev) * 100
    
    print(f"\n==================== {case_name.upper()} ====================")
    print(f"Revenue FY27E-FY31E ($M): {[f'{r:,.1f}' for r in revs]}")
    print(f"EBIT FY27E-FY31E ($M):    {[f'{e:,.1f}' for e in ebits]}")
    print(f"FCFF FY27E-FY31E ($M):    {[f'{f:,.1f}' for f in fcffs]}")
    print(f"PV of FCFFs ($M):         {[f'{pv:,.1f}' for pv in pv_fcffs]}")
    print(f"Cum PV Explicit FCFF:     ${cum_pv_fcff:,.2f} M")
    print(f"Terminal Value @ FY31E:   ${term_val:,.2f} M")
    print(f"PV of Terminal Value:     ${pv_term_val:,.2f} M ({tv_pct_ev:.1f}% of EV)")
    print(f"Implied Enterprise Value: ${implied_ev:,.2f} M (${implied_ev/1e6:.2f} T)")
    print(f"Implied Equity Value:     ${implied_eq_val:,.2f} M (${implied_eq_val/1e6:.2f} T)")
    print(f"Implied Share Price:      ${implied_share_price:.2f}")
    print(f"Current Stock Price:      ${stock_price:.2f}")
    print(f"Implied Upside/(Downside):{upside:+.1f}%")
    print(f"Implied EV / FY27E EBIT:  {implied_ev / ebits[0]:.1f}x")
    return implied_share_price, fcffs, pv_term_val

# Bear Case
p_bear, _, _ = run_dcf(
    "Bear Case",
    [0.55, 0.25, 0.15, 0.08, 0.05],
    [0.68, 0.67, 0.66, 0.65, 0.64],
    [0.55, 0.54, 0.53, 0.52, 0.50],
    [0.035, 0.032, 0.030, 0.030, 0.028],
    [0.018, 0.018, 0.018, 0.018, 0.018],
    [0.015, 0.015, 0.015, 0.015, 0.015],
    0.025,
    0.135
)

# Base Case
p_base, base_fcffs, base_pv_tv = run_dcf(
    "Base Case",
    [0.812, 0.406, 0.250, 0.150, 0.100],
    [0.730, 0.735, 0.730, 0.725, 0.720],
    [0.610, 0.620, 0.615, 0.605, 0.595],
    [0.030, 0.028, 0.026, 0.025, 0.025],
    [0.016, 0.016, 0.016, 0.016, 0.016],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.030,
    0.115
)

# Bull Case
p_bull, _, _ = run_dcf(
    "Bull Case",
    [0.950, 0.500, 0.320, 0.200, 0.150],
    [0.750, 0.755, 0.750, 0.745, 0.740],
    [0.630, 0.640, 0.635, 0.625, 0.615],
    [0.028, 0.026, 0.025, 0.024, 0.023],
    [0.015, 0.015, 0.015, 0.015, 0.015],
    [0.008, 0.008, 0.008, 0.008, 0.008],
    0.035,
    0.100
)

# Sensitivity Table 1: WACC vs Terminal g
print("\n=== SENSITIVITY TABLE 1: WACC vs Terminal g ===")
wacc_axis = [0.1050, 0.1100, 0.1150, 0.1200, 0.1250]
term_g_axis = [0.0200, 0.0250, 0.0300, 0.0350, 0.0400]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]

lbl = "WACC \\ g"
header_str = f"{lbl:<10} | " + " | ".join(f"{g*100:.2f}%{' ':4}" for g in term_g_axis)
print(header_str)
print("-" * len(header_str))

for w in wacc_axis:
    dfs = [1 / ((1 + w) ** p) for p in periods]
    pv_fcffs = sum(f * d for f, d in zip(base_fcffs, dfs))
    row_strs = []
    for g in term_g_axis:
        tv = (base_fcffs[-1] * (1 + g)) / (w - g)
        pv_tv = tv / ((1 + w) ** periods[-1])
        ev = pv_fcffs + pv_tv
        eq = ev - net_debt
        price = eq / shares_out
        row_strs.append(f"${price:6.2f}")
    print(f"{w*100:.2f}%{' ':4} | " + " | ".join(row_strs))
