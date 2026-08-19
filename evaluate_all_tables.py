import numpy as np

stock_price = 226.55
shares_out = 24300.0
total_debt = 8500.0
cash_st_inv = 62600.0
net_debt = total_debt - cash_st_inv
tax_rate = 0.145
hist_rev_last = 215938.0

# Base case parameters
base_wacc = 0.115
base_term_g = 0.030
periods = [0.5, 1.5, 2.5, 3.5, 4.5]

def get_base_fcffs():
    rev_g = [0.812, 0.406, 0.250, 0.150, 0.100]
    ebit_m = [0.610, 0.620, 0.615, 0.605, 0.595]
    capex_p = [0.030, 0.028, 0.026, 0.025, 0.025]
    dna_p = [0.016, 0.016, 0.016, 0.016, 0.016]
    nwc_p = [0.010, 0.010, 0.010, 0.010, 0.010]
    
    revs = []
    pr = hist_rev_last
    for g in rev_g:
        r = pr * (1 + g)
        revs.append(r)
        pr = r
    
    ebits = [r * m for r, m in zip(revs, ebit_m)]
    nopats = [eb * (1 - tax_rate) for eb in ebits]
    dnas = [r * dna for r, dna in zip(revs, dna_p)]
    capexs = [r * cx for r, cx in zip(revs, capex_p)]
    
    nwcs = []
    pr = hist_rev_last
    for r, nwc in zip(revs, nwc_p):
        nwcs.append((r - pr) * nwc)
        pr = r
        
    fcffs = [np_ + d - cx - nwc for np_, d, cx, nwc in zip(nopats, dnas, capexs, nwcs)]
    return revs, ebits, fcffs

revs, ebits, fcffs = get_base_fcffs()
dfs = [1 / ((1 + base_wacc) ** p) for p in periods]
pv_fcffs = [f * d for f, d in zip(fcffs, dfs)]
cum_pv_fcff = sum(pv_fcffs)
term_val = (fcffs[-1] * (1 + base_term_g)) / (base_wacc - base_term_g)
pv_term_val = term_val / ((1 + base_wacc) ** periods[-1])
base_ev = cum_pv_fcff + pv_term_val
base_price = (base_ev - net_debt) / shares_out

print(f"Base Implied Share Price: ${base_price:.2f}")

# Table 2: FY27E Rev Growth vs FY31E EBIT Margin
print("\n=== SENSITIVITY TABLE 2: FY27E Rev Growth vs FY31E EBIT Margin ===")
rev_growth_axis = [0.650, 0.730, 0.812, 0.890, 0.970]
ebit_margin_axis = [0.550, 0.575, 0.595, 0.615, 0.635]

lbl = "Rev G \\ EBIT M"
header_str = f"{lbl:<15} | " + " | ".join(f"{m*100:.1f}%{' ':4}" for m in ebit_margin_axis)
print(header_str)
print("-" * len(header_str))

for rg in rev_growth_axis:
    row_strs = []
    scale_rev = rg / 0.812
    for em in ebit_margin_axis:
        scale_em = em / 0.595
        # Scaled explicit PVs and terminal PV
        scaled_pv_fcffs = sum(pv_fcffs[:-1]) * scale_rev + (pv_fcffs[-1] * scale_rev * scale_em)
        scaled_pv_tv = pv_term_val * scale_rev * scale_em
        ev = scaled_pv_fcffs + scaled_pv_tv
        eq = ev - net_debt
        price = eq / shares_out
        row_strs.append(f"${price:6.2f}")
    print(f"{rg*100:.1f}%{' ':9} | " + " | ".join(row_strs))

# Table 3: Beta vs Risk-Free Rate
print("\n=== SENSITIVITY TABLE 3: Beta vs 10Y UST Yield ===")
beta_axis = [1.80, 2.00, 2.22, 2.40, 2.60]
rf_axis = [0.0420, 0.0445, 0.0470, 0.0495, 0.0520]

lbl = "Beta \\ Rf"
header_str = f"{lbl:<15} | " + " | ".join(f"{rf*100:.2f}%{' ':4}" for rf in rf_axis)
print(header_str)
print("-" * len(header_str))

for b in beta_axis:
    row_strs = []
    for rf in rf_axis:
        ke = rf + b * 0.055
        # Adjusted discount rate accounting for capital structure (equity weight ~101%, net cash debt weight ~ -1%)
        # Kd_after_tax = 0.048 * (1 - 0.145) = 0.0410
        # WACC = Ke * 1.010 - 0.0410 * 0.010 = Ke - ~0.0004
        w = ke * 1.0099 - 0.04104 * 0.0099
        dfs_dyn = [1 / ((1 + w) ** p) for p in periods]
        pv_f = sum(f * d for f, d in zip(fcffs, dfs_dyn))
        tv_dyn = (fcffs[-1] * (1 + base_term_g)) / (w - base_term_g)
        pv_t = tv_dyn / ((1 + w) ** periods[-1])
        ev = pv_f + pv_t
        eq = ev - net_debt
        price = eq / shares_out
        row_strs.append(f"${price:6.2f}")
    print(f"{b:<15.2f} | " + " | ".join(row_strs))
