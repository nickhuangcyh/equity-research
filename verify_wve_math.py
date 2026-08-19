def run_wve_case(name, g_rev, gross_m, rd, sga, term_fcf, g_term, wacc):
    rev = [42.70]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    gross_profit = [rev[i] * gross_m[i-1] for i in range(1, 6)]
    ebit = [gross_profit[i-1] - rd[i-1] - sga[i-1] for i in range(1, 6)]
    tax = [max(0, e * 0.21) for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * 0.030 for i in range(1, 6)]
    capex = [rev[i] * 0.050 for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * 0.030 for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    tv = (term_fcf * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    net_debt = -447.01  # Net cash
    equity_val = ev - net_debt  # EV - (-447.01) = EV + 447.01
    shares = 200.17
    implied_p = equity_val / shares
    curr_p = 5.22
    upside = (implied_p / curr_p - 1) * 100
    
    print(f"=== {name} ===")
    print(f"Revenues ($M): {[round(r, 2) for r in rev[1:]]}")
    print(f"Gross Profit ($M): {[round(gp, 2) for gp in gross_profit]}")
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
    return implied_p, ufcf

# WACC Calculation
wacc_base = 0.1350
g_term_base = 0.030

print("--- SCENARIO RUNS ---")
run_wve_case(
    "BEAR CASE (Selector = 1)",
    [0.054, 0.333, 0.500, 0.556, 0.571],
    [0.800, 0.810, 0.820, 0.830, 0.840],
    [195.0, 205.0, 215.0, 225.0, 235.0],
    [45.0, 50.0, 60.0, 72.0, 85.0],
    40.0,
    0.020,
    0.1350
)

_, base_ufcf = run_wve_case(
    "BASE CASE (Selector = 2)",
    [0.218, 0.635, 0.765, 0.867, 0.714],
    [0.820, 0.830, 0.840, 0.850, 0.860],
    [195.0, 210.0, 225.0, 240.0, 255.0],
    [45.0, 52.0, 65.0, 80.0, 95.0],
    110.0,
    0.030,
    0.1350
)

run_wve_case(
    "BULL CASE (Selector = 3)",
    [0.522, 0.846, 1.083, 0.920, 0.667],
    [0.840, 0.850, 0.860, 0.870, 0.880],
    [195.0, 215.0, 235.0, 255.0, 275.0],
    [45.0, 55.0, 75.0, 95.0, 115.0],
    240.0,
    0.035,
    0.1350
)

print("\n--- SENSITIVITY TABLE 1: WACC vs. Terminal Growth (g) ---")
t1_wacc = [0.1150, 0.1250, 0.1350, 0.1450, 0.1550]
t1_g = [0.020, 0.025, 0.030, 0.035, 0.040]
print("WACC \\ g  " + "  ".join([f"{g:7.1%}" for g in t1_g]))
for w in t1_wacc:
    row_strs = []
    for g in t1_g:
        pv_fcf = sum(cf / ((1 + w)**p) for cf, p in zip(base_ufcf, [0.5, 1.5, 2.5, 3.5, 4.5]))
        tv = (110.0 * (1 + g)) / (w - g)
        pv_tv = tv / ((1 + w)**4.5)
        ev = pv_fcf + pv_tv
        eq = ev - (-447.01)
        p = eq / 200.17
        row_strs.append(f"${p:6.2f}")
    print(f"{w:8.2%}  " + "  ".join(row_strs))

print("\n--- SENSITIVITY TABLE 2: Terminal FCF vs. Revenue Growth Delta ---")
t2_fcf = [70.0, 90.0, 110.0, 130.0, 150.0]
t2_rev_delta = [-0.20, -0.10, 0.00, 0.10, 0.20]
print("Term FCF \\ Rev Δ " + "  ".join([f"{rd:+7.1%}" for rd in t2_rev_delta]))
for fcf in t2_fcf:
    row_strs = []
    for rd in t2_rev_delta:
        pv_fcf = sum((cf * (1 + rd)) / ((1 + wacc_base)**p) for cf, p in zip(base_ufcf, [0.5, 1.5, 2.5, 3.5, 4.5]))
        tv = (fcf * (1 + g_term_base)) / (wacc_base - g_term_base)
        pv_tv = tv / ((1 + wacc_base)**4.5)
        ev = pv_fcf + pv_tv
        eq = ev - (-447.01)
        p = eq / 200.17
        row_strs.append(f"${p:6.2f}")
    print(f"${fcf:6.1f}M  " + "  ".join(row_strs))

print("\n--- SENSITIVITY TABLE 3: Beta vs. Risk-Free Rate ---")
t3_beta = [1.29, 1.49, 1.69, 1.89, 2.09]
t3_rf = [0.0369, 0.0419, 0.0469, 0.0519, 0.0569]
print("Beta \\ Rf  " + "  ".join([f"{rf:7.2%}" for rf in t3_rf]))
for b in t3_beta:
    row_strs = []
    for rf in t3_rf:
        w = rf + b * 0.055
        pv_fcf = sum(cf / ((1 + w)**p) for cf, p in zip(base_ufcf, [0.5, 1.5, 2.5, 3.5, 4.5]))
        tv = (110.0 * (1 + g_term_base)) / (w - g_term_base)
        pv_tv = tv / ((1 + w)**4.5)
        ev = pv_fcf + pv_tv
        eq = ev - (-447.01)
        p = eq / 200.17
        row_strs.append(f"${p:6.2f}")
    print(f"{b:8.2f}  " + "  ".join(row_strs))
