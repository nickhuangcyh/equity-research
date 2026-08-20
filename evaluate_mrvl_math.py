import numpy as np

# MRVL Inputs
stock_price = 237.27
shares_out = 893.30  # M
market_cap = stock_price * shares_out # $M
total_debt = 4961.30   # $M
cash_st_inv = 3843.60 # $M
net_debt = total_debt - cash_st_inv # $1,117.70 M
current_ev = market_cap + net_debt
tax_rate = 0.21

print(f"Stock Price: ${stock_price:.2f}")
print(f"Shares Out: {shares_out:.2f} M")
print(f"Market Cap: ${market_cap:,.2f} M (${market_cap/1e3:.2f} B)")
print(f"Total Debt: ${total_debt:,.2f} M")
print(f"Cash & ST Inv: ${cash_st_inv:,.2f} M")
print(f"Net Debt: ${net_debt:,.2f} M (${net_debt/1e3:.2f} B)")
print(f"Enterprise Value: ${current_ev:,.2f} M (${current_ev/1e3:.2f} B)")

# Historical
hist_rev = [5508.0, 5767.0, 8195.0]
hist_ebit = [-568.0, -720.0, 1323.0]

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
    implied_eq_val = implied_ev - net_debt
    implied_share_price = implied_eq_val / shares_out
    upside = (implied_share_price / stock_price - 1) * 100
    tv_pct_ev = (pv_term_val / implied_ev) * 100
    
    print(f"\n==================== {case_name.upper()} ====================")
    print(f"Revenue FY27E-FY31E ($M): {[f'{r:,.1f}' for r in revs]}")
    print(f"EBIT FY27E-FY31E ($M):    {[f'{e:,.1f}' for e in ebits]}")
    print(f"NOPAT FY27E-FY31E ($M):   {[f'{np_:,.1f}' for np_ in nopats]}")
    print(f"D&A FY27E-FY31E ($M):     {[f'{d:,.1f}' for d in dnas]}")
    print(f"CapEx FY27E-FY31E ($M):   {[f'{c:,.1f}' for c in capexs]}")
    print(f"ΔNWC FY27E-FY31E ($M):    {[f'{n:,.1f}' for n in nwcs]}")
    print(f"FCFF FY27E-FY31E ($M):    {[f'{f:,.1f}' for f in fcffs]}")
    print(f"PV of FCFFs ($M):         {[f'{pv:,.1f}' for pv in pv_fcffs]}")
    print(f"Cum PV Explicit FCFF:     ${cum_pv_fcff:,.2f} M")
    print(f"Terminal Year FCFF:       ${term_fcff:,.2f} M")
    print(f"Terminal Value @ FY31E:   ${term_val:,.2f} M")
    print(f"PV of Terminal Value:     ${pv_term_val:,.2f} M ({tv_pct_ev:.1f}% of EV)")
    print(f"Implied Enterprise Value: ${implied_ev:,.2f} M (${implied_ev/1e3:.2f} B)")
    print(f"Implied Equity Value:     ${implied_eq_val:,.2f} M (${implied_eq_val/1e3:.2f} B)")
    print(f"Implied Share Price:      ${implied_share_price:.2f}")
    print(f"Current Stock Price:      ${stock_price:.2f}")
    print(f"Implied Upside/(Downside):{upside:+.1f}%")
    print(f"Implied EV / FY27E EBIT:  {implied_ev / ebits[0]:.1f}x")
    return implied_share_price, fcffs, pv_fcffs, cum_pv_fcff, pv_term_val, implied_ev, implied_eq_val

# Bear Case
p_bear, fcff_bear, pv_fcff_bear, cum_pv_bear, pv_tv_bear, ev_bear, eq_bear = run_dcf(
    "Bear Case",
    [0.250, 0.200, 0.150, 0.100, 0.070],
    [0.480, 0.485, 0.490, 0.495, 0.500],
    [0.180, 0.200, 0.220, 0.230, 0.240],
    [0.048, 0.045, 0.042, 0.040, 0.038],
    [0.120, 0.100, 0.085, 0.075, 0.070],
    [0.025, 0.025, 0.025, 0.025, 0.025],
    0.0250,
    0.1300
)

# Base Case
p_base, fcff_base, pv_fcff_base, cum_pv_base, pv_tv_base, ev_base, eq_base = run_dcf(
    "Base Case",
    [0.403, 0.435, 0.300, 0.200, 0.150],
    [0.535, 0.550, 0.560, 0.565, 0.570],
    [0.240, 0.280, 0.300, 0.320, 0.330],
    [0.045, 0.042, 0.040, 0.038, 0.035],
    [0.120, 0.095, 0.080, 0.070, 0.065],
    [0.020, 0.020, 0.020, 0.020, 0.020],
    0.0300,
    0.1150
)

# Bull Case
p_bull, fcff_bull, pv_fcff_bull, cum_pv_bull, pv_tv_bull, ev_bull, eq_bull = run_dcf(
    "Bull Case",
    [0.480, 0.500, 0.350, 0.250, 0.180],
    [0.550, 0.570, 0.580, 0.585, 0.590],
    [0.260, 0.300, 0.330, 0.350, 0.360],
    [0.042, 0.040, 0.038, 0.035, 0.032],
    [0.120, 0.090, 0.075, 0.065, 0.060],
    [0.015, 0.015, 0.015, 0.015, 0.015],
    0.0350,
    0.1050
)

# Sensitivity Tables Evaluation
print("\n" + "="*50)
print("SENSITIVITY TABLE 1: WACC vs Terminal Growth")
print("="*50)
term_g_axis = [0.0200, 0.0250, 0.0300, 0.0350, 0.0400]
wacc_axis = [0.1050, 0.1100, 0.1150, 0.1200, 0.1250]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]

print("WACC \\ g\t" + "\t".join(f"{g*100:.1f}%" for g in term_g_axis))
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
print("SENSITIVITY TABLE 2: FY27E Rev Growth vs FY31E EBIT Margin")
print("="*50)
ebit_margin_axis = [0.290, 0.310, 0.330, 0.350, 0.370]
rev_growth_axis = [0.320, 0.360, 0.403, 0.440, 0.480]
base_rev_g = 0.403
base_ebit_m = 0.330

print("Rev \\ EBIT\t" + "\t".join(f"{m*100:.1f}%" for m in ebit_margin_axis))
for rg in rev_growth_axis:
    row_str = f"{rg*100:.1f}%\t"
    for em in ebit_margin_axis:
        # Scaled FCFs
        scale_rev = rg / base_rev_g
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
rf_axis = [0.0415, 0.0440, 0.0465, 0.0490, 0.0515]
beta_axis = [1.80, 2.00, 2.24, 2.45, 2.65]

print("Beta \\ Rf\t" + "\t".join(f"{rf*100:.2f}%" for rf in rf_axis))
for beta in beta_axis:
    row_str = f"{beta:.2f}\t"
    for rf in rf_axis:
        w = rf + beta * 0.0550 - 0.0547
        pv_fcff_w = sum(f / ((1 + w)**p) for f, p in zip(fcff_base, periods))
        tv_w = (fcff_base[-1] * (1 + 0.0300)) / (w - 0.0300)
        pv_tv_w = tv_w / ((1 + w)**periods[-1])
        ev_w = pv_fcff_w + pv_tv_w
        eq_w = ev_w - net_debt
        price_w = eq_w / shares_out
        row_str += f"${price_w:.2f}\t"
    print(row_str)
