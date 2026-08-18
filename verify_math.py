# Verify valuation math for Freshpet (FRPT) DCF Model

def run_case(name, g_rev, ebit_m, g_term, wacc):
    rev = [1102.0]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [e * 0.21 for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * [0.08, 0.075, 0.070, 0.065, 0.060][i-1] for i in range(1, 6)]
    capex = [rev[i] * [0.12, 0.10, 0.08, 0.06, 0.05][i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * 0.02 for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    tv = (ufcf[-1] * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    net_debt = 119.3
    equity_val = ev - net_debt
    shares = 56.04
    implied_p = equity_val / shares
    curr_p = 72.76
    upside = (implied_p / curr_p - 1) * 100
    
    print(f"=== {name} ===")
    print(f"Revenues: {[round(r, 1) for r in rev[1:]]}")
    print(f"EBIT: {[round(e, 1) for e in ebit]}")
    print(f"NOPAT: {[round(n, 1) for n in nopat]}")
    print(f"D&A: {[round(d, 1) for d in da]}")
    print(f"CapEx: {[round(c, 1) for c in capex]}")
    print(f"UFCF: {[round(u, 1) for u in ufcf]}")
    print(f"PV UFCFs: ${round(sum_pv_ufcf, 1)}M")
    print(f"Terminal Value: ${round(tv, 1)}M, PV TV: ${round(pv_tv, 1)}M ({(pv_tv/ev*100):.1f}% of EV)")
    print(f"Enterprise Value: ${round(ev, 1)}M")
    print(f"Net Debt: ${net_debt}M")
    print(f"Equity Value: ${round(equity_val, 1)}M")
    print(f"Implied Share Price: ${implied_p:.2f} (vs Current ${curr_p:.2f}) -> {upside:+.1f}%\n")

# WACC Calculation:
# Ke = 4.25% + 1.35 * 5.5% = 11.675%
# Kd_after = 4.5% * (1 - 0.21) = 3.555%
# MCap = 72.76 * 56.04 = 4077.47M
# Net Debt = 397.3 - 278.0 = 119.30M
# EV = 4077.47 + 119.30 = 4196.77M
# We = 4077.47 / 4196.77 = 97.16%
# Wd = 119.30 / 4196.77 = 2.84%
# WACC = 11.675% * 0.9716 + 3.555% * 0.0284 = 11.444% (0.11444)

wacc_val = (0.11675 * (4077.47/4196.77)) + (0.03555 * (119.30/4196.77))
print(f"Calculated WACC: {wacc_val*100:.2f}%\n")

run_case("BASE CASE", [0.12, 0.11, 0.095, 0.08, 0.065], [0.085, 0.10, 0.115, 0.125, 0.135], 0.025, wacc_val)
run_case("BEAR CASE", [0.09, 0.075, 0.06, 0.05, 0.04], [0.075, 0.085, 0.09, 0.095, 0.10], 0.020, wacc_val)
run_case("BULL CASE", [0.15, 0.14, 0.125, 0.11, 0.09], [0.095, 0.115, 0.135, 0.15, 0.16], 0.030, wacc_val)
