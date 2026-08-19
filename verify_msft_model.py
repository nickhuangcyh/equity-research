import openpyxl

def run_msft_case(name, g_rev, ebit_m, da_m, capex_m, nwc_m, g_term, wacc):
    rev_hist = [198270.0, 211915.0, 245120.0, 281724.0, 331839.0]
    rev = [331839.0]
    for g in g_rev:
        rev.append(rev[-1] * (1 + g))
    
    # Projections FY27E to FY31E
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [e * 0.20 for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * da_m[i-1] for i in range(1, 6)]
    capex = [rev[i] * capex_m[i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * nwc_m[i-1] for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcf = [cf / ((1 + wacc)**p) for cf, p in zip(ufcf, periods)]
    sum_pv_ufcf = sum(pv_ufcf)
    
    tv_nom = (ufcf[-1] * (1 + g_term)) / (wacc - g_term)
    pv_tv = tv_nom / ((1 + wacc)**4.5)
    
    ev = sum_pv_ufcf + pv_tv
    total_debt = 47600.00
    cash_st = 76843.00
    net_debt = total_debt - cash_st  # -$29,243.00
    equity_val = ev - net_debt  # ev + 29243
    shares = 7435.00
    implied_p = equity_val / shares
    curr_p = 481.59
    upside = (implied_p / curr_p - 1) * 100
    
    tv_pct_ev = (pv_tv / ev) * 100
    fy31_ebitda = ebit[-1] + da[-1]
    implied_exit_multiple = ev / fy31_ebitda
    
    print(f"==========================================")
    print(f"=== {name} ===")
    print(f"==========================================")
    print(f"Revenues FY27E-FY31E ($M): {[round(r, 1) for r in rev[1:]]}")
    print(f"Revenue YoY Growth: {[f'{g*100:.1f}%' for g in g_rev]}")
    print(f"EBIT ($M): {[round(e, 1) for e in ebit]}")
    print(f"EBIT Margin: {[f'{m*100:.1f}%' for m in ebit_m]}")
    print(f"NOPAT ($M): {[round(n, 1) for n in nopat]}")
    print(f"D&A ($M): {[round(d, 1) for d in da]}")
    print(f"CapEx ($M): {[round(c, 1) for c in capex]}")
    print(f"Δ NWC ($M): {[round(w, 1) for w in nwc]}")
    print(f"UFCF ($M): {[round(u, 1) for u in ufcf]}")
    print(f"FCF Conversion Margin: {[f'{(u/r)*100:.1f}%' for u, r in zip(ufcf, rev[1:])]}")
    print(f"\n--- Valuation Bridge ---")
    print(f"Explicit 5-Year PV of FCFs: ${sum_pv_ufcf:,.1f} M ({sum_pv_ufcf/ev*100:.1f}% of EV)")
    print(f"Terminal Value (Nominal): ${tv_nom:,.1f} M")
    print(f"PV of Terminal Value: ${pv_tv:,.1f} M ({tv_pct_ev:.1f}% of EV)")
    print(f"Enterprise Value (EV): ${ev:,.1f} M")
    print(f"(-) Total Debt: $(47,600.0) M")
    print(f"(+) Cash & Short-Term Investments: $76,843.0 M")
    print(f"Net Debt (Net Cash Adjustment): ${-net_debt:,.1f} M")
    print(f"Implied Equity Value: ${equity_val:,.1f} M")
    print(f"Diluted Shares Outstanding: {shares:,.2f} M")
    print(f"Implied Intrinsic Value Per Share: ${implied_p:.2f}")
    print(f"Current Market Stock Price: ${curr_p:.2f}")
    print(f"Implied Upside / (Downside): {upside:+.1f}%")
    print(f"PV of TV % of EV: {tv_pct_ev:.1f}%")
    print(f"Implied Exit Multiple (EV / FY31E EBITDA): {implied_exit_multiple:.1f}x\n")
    
    return {
        "name": name,
        "rev": rev[1:],
        "ebit": ebit,
        "ufcf": ufcf,
        "sum_pv_ufcf": sum_pv_ufcf,
        "pv_tv": pv_tv,
        "ev": ev,
        "equity_val": equity_val,
        "implied_p": implied_p,
        "upside": upside,
        "tv_pct_ev": tv_pct_ev,
        "implied_exit_multiple": implied_exit_multiple,
        "wacc": wacc,
        "g_term": g_term,
    }

# 1. Calculate WACC
rf = 0.0475
beta = 1.10
erp = 0.0550
ke = rf + beta * erp # 10.80%

kd_pre = 0.0495
tax_rate = 0.2000
kd_after = kd_pre * (1 - tax_rate) # 3.96%

curr_p = 481.59
shares = 7435.00
mcap = curr_p * shares # 3,580,621.65
debt = 47600.00
total_cap = mcap + debt # 3,628,221.65

we = mcap / total_cap # 98.6881%
wd = debt / total_cap # 1.3119%
wacc_base = (ke * we) + (kd_after * wd) # 10.7103%

print("=== WACC PARAMETERS ===")
print(f"Risk-Free Rate (Rf): {rf*100:.2f}%")
print(f"Equity Beta: {beta:.2f}")
print(f"Equity Risk Premium (ERP): {erp*100:.2f}%")
print(f"Cost of Equity (Ke): {ke*100:.2f}%")
print(f"Pre-Tax Cost of Debt (Kd): {kd_pre*100:.2f}%")
print(f"After-Tax Cost of Debt: {kd_after*100:.2f}%")
print(f"Market Cap: ${mcap:,.1f} M ({we*100:.2f}%)")
print(f"Total Debt: ${debt:,.1f} M ({wd*100:.2f}%)")
print(f"Base WACC: {wacc_base*100:.2f}%\n")

# 2. Run Scenarios
# Base Case
base_res = run_msft_case(
    "BASE CASE (AI Cloud Scaling & Copilot Monetization)",
    [0.145, 0.130, 0.115, 0.100, 0.085],
    [0.470, 0.478, 0.486, 0.493, 0.500],
    [0.130, 0.140, 0.135, 0.130, 0.120],
    [0.280, 0.230, 0.190, 0.160, 0.135],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.0300,
    wacc_base
)

# Bear Case
bear_res = run_msft_case(
    "BEAR CASE (AI Digestion & Cloud Slowdown)",
    [0.110, 0.095, 0.085, 0.075, 0.065],
    [0.455, 0.450, 0.445, 0.440, 0.435],
    [0.135, 0.145, 0.140, 0.135, 0.125],
    [0.300, 0.260, 0.220, 0.180, 0.150],
    [0.015, 0.015, 0.015, 0.015, 0.015],
    0.0250,
    0.1150
)

# Bull Case
bull_res = run_msft_case(
    "BULL CASE (AI Supercycle & Azure Platform Dominance)",
    [0.175, 0.160, 0.140, 0.125, 0.105],
    [0.480, 0.492, 0.505, 0.515, 0.525],
    [0.125, 0.135, 0.130, 0.125, 0.115],
    [0.260, 0.210, 0.175, 0.145, 0.120],
    [0.010, 0.010, 0.010, 0.010, 0.010],
    0.0350,
    0.1000
)

# 3. Sensitivity Grid Calculations (Base Case UFCFs)
ufcfs = base_res["ufcf"]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]
net_cash_adj = 76843.0 - 47600.0 # +29243.0

print("=== SENSITIVITY TABLE 1: WACC vs Terminal Growth Rate (g) ===")
g_range = [0.020, 0.025, 0.030, 0.035, 0.040]
w_range = [0.0971, 0.1021, 0.1071, 0.1121, 0.1171]

header_str = "WACC \\ g\t" + "\t".join([f"{g*100:.1f}%" for g in g_range])
print(header_str)
for w in w_range:
    row_vals = []
    for g in g_range:
        pv_f = sum([cf / ((1 + w)**p) for cf, p in zip(ufcfs, periods)])
        pv_t = ((ufcfs[-1] * (1 + g)) / (w - g)) / ((1 + w)**4.5)
        p_val = (pv_f + pv_t + net_cash_adj) / shares
        row_vals.append(f"${p_val:.2f}")
    print(f"{w*100:.2f}%\t" + "\t".join(row_vals))

print("\n=== SENSITIVITY TABLE 2: Revenue Growth Multiplier vs Target FY31E EBIT Margin ===")
gr_mults = [0.80, 0.90, 1.00, 1.10, 1.20]
ebit_margins = [0.460, 0.480, 0.500, 0.520, 0.540]
header_str2 = "Growth \\ EBIT\t" + "\t".join([f"{em*100:.1f}%" for em in ebit_margins])
print(header_str2)
base_ebit_31 = 0.500
w_base = wacc_base
g_base = 0.030
for gr in gr_mults:
    row_vals = []
    for em in ebit_margins:
        adj_factor = gr * (em / base_ebit_31)
        adj_ufcfs = [cf * adj_factor for cf in ufcfs]
        pv_f = sum([cf / ((1 + w_base)**p) for cf, p in zip(adj_ufcfs, periods)])
        pv_t = ((adj_ufcfs[-1] * (1 + g_base)) / (w_base - g_base)) / ((1 + w_base)**4.5)
        p_val = (pv_f + pv_t + net_cash_adj) / shares
        row_vals.append(f"${p_val:.2f}")
    print(f"{gr*100:.0f}%\t\t" + "\t".join(row_vals))

print("\n=== SENSITIVITY TABLE 3: Equity Beta vs Risk-Free Rate (Rf) ===")
betas = [0.90, 1.00, 1.10, 1.20, 1.30]
rfs = [0.0375, 0.0425, 0.0475, 0.0525, 0.0575]
header_str3 = "Beta \\ Rf\t" + "\t".join([f"{r*100:.2f}%" for r in rfs])
print(header_str3)
for b in betas:
    row_vals = []
    for r in rfs:
        ke_temp = r + b * erp
        w_temp = (ke_temp * we) + (kd_after * wd)
        pv_f = sum([cf / ((1 + w_temp)**p) for cf, p in zip(ufcfs, periods)])
        pv_t = ((ufcfs[-1] * (1 + g_base)) / (w_temp - g_base)) / ((1 + w_temp)**4.5)
        p_val = (pv_f + pv_t + net_cash_adj) / shares
        row_vals.append(f"${p_val:.2f}")
    print(f"{b:.2f}\t\t" + "\t".join(row_vals))
