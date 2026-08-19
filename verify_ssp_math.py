#!/usr/bin/env python3
"""
verify_ssp_math.py - Complete mathematical evaluation and verification of SSP DCF Model
"""

def run_ssp_case(name, g_rev, ebit_m, da_pct, capex_pct, nwc_pct, g_term, wacc):
    rev = [2150.60]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [max(0, e * 0.25) for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * da_pct[i-1] for i in range(1, 6)]
    capex = [rev[i] * capex_pct[i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * nwc_pct[i-1] for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    tv = (ufcf[-1] * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    total_debt = 2660.00
    cash = 54.67
    net_debt = total_debt - cash  # 2605.33
    equity_val = max(0, ev - net_debt)
    shares = 87.50
    implied_p = equity_val / shares
    curr_p = 3.35
    upside = (implied_p / curr_p - 1) * 100 if implied_p > 0 else -100.0
    
    print(f"=== {name} ===")
    print(f"Revenues ($M): {[round(r, 2) for r in rev[1:]]}")
    print(f"EBIT ($M): {[round(e, 2) for e in ebit]}")
    print(f"EBIT Margin (%): {[f'{e/r*100:.1f}%' for e, r in zip(ebit, rev[1:])]}")
    print(f"NOPAT ($M): {[round(n, 2) for n in nopat]}")
    print(f"D&A ($M): {[round(d, 2) for d in da]}")
    print(f"CapEx ($M): {[round(c, 2) for c in capex]}")
    print(f"Δ NWC ($M): {[round(w, 2) for w in nwc]}")
    print(f"UFCF ($M): {[round(u, 2) for u in ufcf]}")
    print(f"PV Explicit UFCFs: ${round(sum_pv_ufcf, 2)}M")
    print(f"Terminal Value: ${round(tv, 2)}M, PV TV: ${round(pv_tv, 2)}M ({(pv_tv/ev*100):.1f}% of EV)")
    print(f"Enterprise Value: ${round(ev, 2)}M")
    print(f"Net Debt: ${net_debt:.2f}M")
    print(f"Equity Value: ${round(equity_val, 2)}M")
    print(f"Implied Share Price: ${implied_p:.2f} (vs Current ${curr_p:.2f}) -> {upside:+.1f}%\n")
    return implied_p, ufcf, sum_pv_ufcf, pv_tv, ev, net_debt, equity_val

# WACC Calculation
rf = 0.0435
beta = 1.95
erp = 0.0550
ke = rf + beta * erp  # 15.075%
kd_pre = 0.0780
tax_rate = 0.2500
kd_after = kd_pre * (1 - tax_rate)  # 5.85%

curr_p = 3.35
shares = 87.50
mkt_cap = curr_p * shares  # 293.125
total_debt = 2660.00
cash = 54.67
net_debt = total_debt - cash  # 2605.33
total_cap = mkt_cap + net_debt  # 2898.455

we = mkt_cap / total_cap  # 10.113%
wd = net_debt / total_cap  # 89.887%
wacc_base = (we * ke) + (wd * kd_after)

print("=== WACC SUMMARY ===")
print(f"Cost of Equity: {ke:.3%}")
print(f"Cost of Debt (after-tax): {kd_after:.3%}")
print(f"Equity Weight: {we:.2%}, Debt Weight: {wd:.2%}")
print(f"Base WACC: {wacc_base:.3%}\n")

# Run Scenarios
print("--- SCENARIO EVALUATIONS ---")
p_bear, ufcf_bear, _, _, _, _, _ = run_ssp_case(
    "BEAR CASE (Selector = 1)",
    [0.040, -0.100, 0.060, -0.080, 0.030],
    [0.095, 0.065, 0.100, 0.060, 0.080],
    [0.065, 0.065, 0.065, 0.065, 0.065],
    [0.022, 0.022, 0.022, 0.022, 0.022],
    [0.005, 0.005, 0.005, 0.005, 0.005],
    0.000,
    0.0800
)

p_base, ufcf_base, pv_fcf_base, pv_tv_base, ev_base, nd_base, eq_base = run_ssp_case(
    "BASE CASE (Selector = 2)",
    [0.090, -0.060, 0.110, -0.050, 0.070],
    [0.125, 0.090, 0.145, 0.100, 0.120],
    [0.060, 0.060, 0.058, 0.060, 0.058],
    [0.020, 0.020, 0.020, 0.020, 0.020],
    [0.005, 0.005, 0.005, 0.005, 0.005],
    0.010,
    wacc_base
)

p_bull, ufcf_bull, _, _, _, _, _ = run_ssp_case(
    "BULL CASE (Selector = 3)",
    [0.140, -0.020, 0.150, -0.020, 0.100],
    [0.150, 0.115, 0.175, 0.130, 0.155],
    [0.058, 0.058, 0.055, 0.058, 0.055],
    [0.018, 0.018, 0.018, 0.018, 0.018],
    [0.005, 0.005, 0.005, 0.005, 0.005],
    0.020,
    0.0650
)

# Sensitivity Tables Evaluation
print("=== SENSITIVITY TABLE 1: WACC vs. Terminal g ===")
g_cols = [0.000, 0.005, 0.010, 0.015, 0.020]
wacc_rows = [wacc_base - 0.010, wacc_base - 0.005, wacc_base, wacc_base + 0.005, wacc_base + 0.010]

print("WACC \\ g\t" + "\t".join(f"{g:.1%}" for g in g_cols))
for w in wacc_rows:
    row_str = f"{w:.2%}\t"
    for g in g_cols:
        pv_f = sum([cf / ((1 + w)**p) for cf, p in zip(ufcf_base, [0.5, 1.5, 2.5, 3.5, 4.5])])
        tv_val = (ufcf_base[-1] * (1 + g)) / (w - g)
        pv_t = tv_val / ((1 + w)**4.5)
        ev_val = pv_f + pv_t
        eq_val = max(0, ev_val - net_debt)
        p_val = eq_val / shares
        row_str += f"${p_val:,.2f}\t"
    print(row_str)

print("\n=== SENSITIVITY TABLE 2: Revenue Growth Scale vs. Target EBIT Margin ===")
ebit_cols = [0.080, 0.100, 0.120, 0.140, 0.160]
growth_rows = [0.80, 0.90, 1.00, 1.10, 1.20]
base_em = 0.120

print("Growth \\ EM\t" + "\t".join(f"{em:.1%}" for em in ebit_cols))
for gr in growth_rows:
    row_str = f"{gr:.0%}\t"
    for em in ebit_cols:
        scale = gr * (em / base_em)
        pv_f = sum([(cf * scale) / ((1 + wacc_base)**p) for cf, p in zip(ufcf_base, [0.5, 1.5, 2.5, 3.5, 4.5])])
        tv_val = (ufcf_base[-1] * scale * (1 + 0.010)) / (wacc_base - 0.010)
        pv_t = tv_val / ((1 + wacc_base)**4.5)
        ev_val = pv_f + pv_t
        eq_val = max(0, ev_val - net_debt)
        p_val = eq_val / shares
        row_str += f"${p_val:,.2f}\t"
    print(row_str)

print("\n=== SENSITIVITY TABLE 3: Equity Beta vs. Risk-Free Rate ===")
rf_cols = [0.0335, 0.0385, 0.0435, 0.0485, 0.0535]
beta_rows = [1.55, 1.75, 1.95, 2.15, 2.35]

print("Beta \\ Rf\t" + "\t".join(f"{r:.2%}" for r in rf_cols))
for b in beta_rows:
    row_str = f"{b:.2f}\t"
    for r in rf_cols:
        ke_temp = r + b * erp
        w_temp = (we * ke_temp) + (wd * kd_after)
        pv_f = sum([cf / ((1 + w_temp)**p) for cf, p in zip(ufcf_base, [0.5, 1.5, 2.5, 3.5, 4.5])])
        tv_val = (ufcf_base[-1] * (1 + 0.010)) / (w_temp - 0.010)
        pv_t = tv_val / ((1 + w_temp)**4.5)
        ev_val = pv_f + pv_t
        eq_val = max(0, ev_val - net_debt)
        p_val = eq_val / shares
        row_str += f"${p_val:,.2f}\t"
    print(row_str)
