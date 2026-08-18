def run_eols_case(name, g_rev, ebit_m, g_term, wacc):
    rev = [266.27]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [max(0, e * 0.21) for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * [0.015, 0.014, 0.013, 0.012, 0.011][i-1] for i in range(1, 6)]
    capex = [rev[i] * [0.011, 0.010, 0.010, 0.009, 0.009][i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * 0.015 for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    tv = (ufcf[-1] * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    net_debt = 38.00
    equity_val = ev - net_debt
    shares = 65.20
    implied_p = equity_val / shares
    curr_p = 8.15
    upside = (implied_p / curr_p - 1) * 100
    
    print(f"=== {name} ===")
    print(f"Revenues ($M): {[round(r, 2) for r in rev[1:]]}")
    print(f"EBIT ($M): {[round(e, 2) for e in ebit]}")
    print(f"NOPAT ($M): {[round(n, 2) for n in nopat]}")
    print(f"D&A ($M): {[round(d, 2) for d in da]}")
    print(f"CapEx ($M): {[round(c, 2) for c in capex]}")
    print(f"Δ NWC ($M): {[round(w, 2) for w in nwc]}")
    print(f"UFCF ($M): {[round(u, 2) for u in ufcf]}")
    print(f"PV UFCFs: ${round(sum_pv_ufcf, 2)}M")
    print(f"Terminal Value: ${round(tv, 2)}M, PV TV: ${round(pv_tv, 2)}M ({(pv_tv/ev*100):.1f}% of EV)")
    print(f"Enterprise Value: ${round(ev, 2)}M")
    print(f"Net Debt: ${net_debt:.2f}M")
    print(f"Equity Value: ${round(equity_val, 2)}M")
    print(f"Implied Share Price: ${implied_p:.2f} (vs Current ${curr_p:.2f}) -> {upside:+.1f}%\n")

# WACC Calculation:
# Ke = 4.70% + 1.36 * 5.50% = 12.18%
# Kd_after = 9.50% * (1 - 0.21) = 7.505%
# MCap = 8.15 * 65.20 = 531.38M
# Net Debt = 125.0 - 87.0 = 38.00M
# EV = 531.38 + 38.00 = 569.38M
# We = 531.38 / 569.38 = 93.326%
# Wd = 38.00 / 569.38 = 6.674%
# WACC = 12.18% * 0.93326 + 7.505% * 0.06674 = 11.8677%

ke = 0.0470 + 1.36 * 0.0550
kd_after = 0.0950 * (1 - 0.21)
mcap = 8.15 * 65.20
net_debt = 125.0 - 87.0
ev_val = mcap + net_debt
we = mcap / ev_val
wd = net_debt / ev_val
wacc_val = (ke * we) + (kd_after * wd)

print(f"Calculated Ke: {ke*100:.2f}%")
print(f"Calculated Kd (after-tax): {kd_after*100:.2f}%")
print(f"Weight Equity: {we*100:.2f}%, Weight Debt: {wd*100:.2f}%")
print(f"Calculated WACC: {wacc_val*100:.2f}%\n")

run_eols_case("BASE CASE", [0.150, 0.135, 0.115, 0.095, 0.075], [0.016, 0.080, 0.134, 0.180, 0.210], 0.025, wacc_val)
run_eols_case("BEAR CASE", [0.100, 0.085, 0.070, 0.060, 0.050], [-0.015, 0.027, 0.065, 0.097, 0.127], 0.020, wacc_val)
run_eols_case("BULL CASE", [0.180, 0.160, 0.140, 0.115, 0.090], [0.040, 0.117, 0.175, 0.232, 0.263], 0.030, wacc_val)
