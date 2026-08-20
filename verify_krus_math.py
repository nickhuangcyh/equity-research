#!/usr/bin/env python3
"""
verify_krus_math.py - Exact mathematical verification for KRUS DCF Model
"""

def run_krus_case(name, g_rev, ebit_m, da_p, capex_p, nwc_p, g_term, wacc):
    rev = [282.80]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [max(0, e * 0.25) for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * da_p[i-1] for i in range(1, 6)]
    capex = [rev[i] * capex_p[i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * nwc_p[i-1] for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    terminal_fcf = ufcf[-1]
    tv = (terminal_fcf * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    net_debt = -50.99  # Net cash
    equity_val = ev - net_debt  # EV - (-50.99) = EV + 50.99
    shares = 12.15
    implied_p = equity_val / shares
    curr_p = 50.14
    upside = (implied_p / curr_p - 1) * 100
    
    print(f"=== {name} ===")
    print(f"Revenues ($M): {[round(r, 2) for r in rev[1:]]}")
    print(f"EBIT ($M): {[round(e, 2) for e in ebit]}")
    print(f"EBIT Margin (%): {[f'{e/r*100:.1f}%' for e, r in zip(ebit, rev[1:])]}")
    print(f"NOPAT ($M): {[round(n, 2) for n in nopat]}")
    print(f"D&A ($M): {[round(d, 2) for d in da]}")
    print(f"CapEx ($M): {[round(c, 2) for c in capex]}")
    print(f"Δ NWC ($M): {[round(w, 2) for w in nwc]}")
    print(f"UFCF ($M): {[round(u, 2) for u in ufcf]}")
    print(f"PV UFCFs: ${round(sum_pv_ufcf, 2)}M")
    print(f"Terminal Value: ${round(tv, 2)}M, PV TV: ${round(pv_tv, 2)}M ({(pv_tv/ev*100):.1f}% of EV)")
    print(f"Enterprise Value: ${round(ev, 2)}M")
    print(f"Net Debt: ${net_debt:.2f}M (Net Cash: ${abs(net_debt):.2f}M)")
    print(f"Equity Value: ${round(equity_val, 2)}M")
    print(f"Implied Share Price: ${implied_p:.2f} (vs Current ${curr_p:.2f}) -> {upside:+.1f}%\n")
    return implied_p, ufcf, rev[1:]

wacc_base = 0.1196
g_term_base = 0.0250

print("--- SCENARIO RUNS ---")
run_krus_case(
    "BEAR CASE (Selector = 1)",
    [0.120, 0.110, 0.090, 0.070, 0.050],
    [-0.015, 0.000, 0.015, 0.025, 0.035],
    [0.060, 0.060, 0.060, 0.060, 0.060],
    [0.110, 0.090, 0.075, 0.065, 0.055],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.020,
    0.1250
)

_, base_ufcf, base_rev = run_krus_case(
    "BASE CASE (Selector = 2)",
    [0.170, 0.180, 0.150, 0.120, 0.090],
    [0.005, 0.025, 0.045, 0.060, 0.070],
    [0.060, 0.060, 0.060, 0.060, 0.060],
    [0.120, 0.100, 0.085, 0.070, 0.060],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.025,
    0.1196
)

run_krus_case(
    "BULL CASE (Selector = 3)",
    [0.210, 0.220, 0.180, 0.150, 0.120],
    [0.020, 0.045, 0.065, 0.080, 0.095],
    [0.060, 0.060, 0.060, 0.060, 0.060],
    [0.130, 0.110, 0.090, 0.075, 0.065],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.030,
    0.1150
)

print("\n--- SENSITIVITY TABLE 1: WACC vs. Terminal Growth (g) ---")
t1_wacc = [0.1000, 0.1100, 0.1196, 0.1300, 0.1400]
t1_g = [0.015, 0.020, 0.025, 0.030, 0.035]
print("WACC \\ g  " + "  ".join([f"{g:7.1%}" for g in t1_g]))
for w in t1_wacc:
    row_strs = []
    for g in t1_g:
        pv_fcf = sum(cf / ((1 + w)**p) for cf, p in zip(base_ufcf, [0.5, 1.5, 2.5, 3.5, 4.5]))
        tv = (base_ufcf[-1] * (1 + g)) / (w - g)
        pv_tv = tv / ((1 + w)**4.5)
        ev = pv_fcf + pv_tv
        eq = ev - (-50.99)
        p = eq / 12.15
        row_strs.append(f"${p:6.2f}")
    print(f"{w:8.2%}  " + "  ".join(row_strs))

print("\n--- SENSITIVITY TABLE 2: Terminal EBIT Margin vs. Revenue Growth Delta ---")
t2_rev_delta = [-0.100, -0.050, 0.000, 0.050, 0.100]
t2_ebit_margins = [0.050, 0.060, 0.070, 0.080, 0.090]
print("EBIT Mgn \\ Rev Δ  " + "  ".join([f"{d:+7.1%}" for d in t2_rev_delta]))
for m in t2_ebit_margins:
    row_strs = []
    for d in t2_rev_delta:
        # Scaled FCFs
        pv_explicit = sum((cf * (1 + d)) / ((1 + wacc_base)**p) for cf, p in zip(base_ufcf[:4], [0.5, 1.5, 2.5, 3.5]))
        # Year 5 scaled
        scaled_rev_y5 = base_rev[-1] * (1 + d)
        scaled_ebit_y5 = scaled_rev_y5 * m
        scaled_nopat_y5 = scaled_ebit_y5 * (1 - 0.25)
        scaled_da_y5 = scaled_rev_y5 * 0.060
        scaled_capex_y5 = scaled_rev_y5 * 0.060
        scaled_ufcf_y5 = scaled_nopat_y5 + scaled_da_y5 - scaled_capex_y5
        pv_y5 = scaled_ufcf_y5 / ((1 + wacc_base)**4.5)
        
        tv = (scaled_ufcf_y5 * (1 + g_term_base)) / (wacc_base - g_term_base)
        pv_tv = tv / ((1 + wacc_base)**4.5)
        ev = pv_explicit + pv_y5 + pv_tv
        eq = ev - (-50.99)
        p = eq / 12.15
        row_strs.append(f"${p:6.2f}")
    print(f"{m:8.1%}       " + "  ".join(row_strs))

print("\n--- SENSITIVITY TABLE 3: Beta vs. Risk-Free Rate (Rf) ---")
t3_rf = [0.0364, 0.0414, 0.0464, 0.0514, 0.0564]
t3_beta = [1.13, 1.23, 1.33, 1.43, 1.53]
print("Beta \\ Rf  " + "  ".join([f"{rf:7.2%}" for rf in t3_rf]))
for b in t3_beta:
    row_strs = []
    for rf in t3_rf:
        w = rf + b * 0.0550
        pv_fcf = sum(cf / ((1 + w)**p) for cf, p in zip(base_ufcf, [0.5, 1.5, 2.5, 3.5, 4.5]))
        tv = (base_ufcf[-1] * (1 + g_term_base)) / (w - g_term_base)
        pv_tv = tv / ((1 + w)**4.5)
        ev = pv_fcf + pv_tv
        eq = ev - (-50.99)
        p = eq / 12.15
        row_strs.append(f"${p:6.2f}")
    print(f"{b:8.2f}  " + "  ".join(row_strs))
