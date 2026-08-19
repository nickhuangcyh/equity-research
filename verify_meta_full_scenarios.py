#!/usr/bin/env python3

def calc_dcf(growth_rates, ebit_margins, da_rates, capex_rates, nwc_rates, wacc, term_g, base_rev=200966.0, tax_rate=0.16, shares=2570.0, net_debt=-6600.0):
    revs = []
    ebits = []
    nopats = []
    das = []
    capexs = []
    nwcs = []
    fcfs = []
    pvs = []
    
    prev_rev = base_rev
    for i in range(5):
        g = growth_rates[i]
        em = ebit_margins[i]
        da_r = da_rates[i]
        cap_r = capex_rates[i]
        nwc_r = nwc_rates[i]
        
        rev = prev_rev * (1 + g)
        ebit = rev * em
        tax = ebit * tax_rate
        nopat = ebit - tax
        da = rev * da_r
        capex = rev * cap_r
        dnwc = (rev - prev_rev) * nwc_r
        fcf = nopat + da - capex - dnwc
        
        t = 0.5 + i
        df = 1.0 / ((1 + wacc) ** t)
        pv = fcf * df
        
        revs.append(rev)
        ebits.append(ebit)
        nopats.append(nopat)
        das.append(da)
        capexs.append(capex)
        nwcs.append(dnwc)
        fcfs.append(fcf)
        pvs.append(pv)
        
        prev_rev = rev
        
    cum_pv_fcf = sum(pvs)
    final_fcf = fcfs[-1]
    term_fcf = final_fcf * (1 + term_g)
    tv = term_fcf / (wacc - term_g)
    pv_tv = tv / ((1 + wacc) ** 4.5)
    
    ev = cum_pv_fcf + pv_tv
    eq_val = ev - net_debt
    implied_price = eq_val / shares
    
    return {
        "revs": revs,
        "fcfs": fcfs,
        "cum_pv_fcf": cum_pv_fcf,
        "pv_tv": pv_tv,
        "ev": ev,
        "eq_val": eq_val,
        "implied_price": implied_price,
        "pv_tv_pct": pv_tv / ev
    }

print("=== SCENARIO COMPARISON ===")
# Base Case
base = calc_dcf(
    growth_rates=[0.160, 0.140, 0.120, 0.100, 0.080],
    ebit_margins=[0.415, 0.420, 0.425, 0.430, 0.435],
    da_rates=[0.110, 0.120, 0.125, 0.125, 0.120],
    capex_rates=[0.320, 0.280, 0.250, 0.230, 0.220],
    nwc_rates=[0.010, 0.010, 0.010, 0.010, 0.010],
    wacc=0.1100,
    term_g=0.030
)
print(f"Base Case: EV=${base['ev']:,.1f}M | EqVal=${base['eq_val']:,.1f}M | Price=${base['implied_price']:.2f} | PV_TV%={base['pv_tv_pct']*100:.1f}%")

# Bear Case
bear = calc_dcf(
    growth_rates=[0.120, 0.100, 0.080, 0.070, 0.060],
    ebit_margins=[0.380, 0.370, 0.360, 0.350, 0.350],
    da_rates=[0.115, 0.120, 0.125, 0.125, 0.120],
    capex_rates=[0.330, 0.300, 0.270, 0.250, 0.240],
    nwc_rates=[0.010, 0.010, 0.010, 0.010, 0.010],
    wacc=0.1200,
    term_g=0.025
)
print(f"Bear Case: EV=${bear['ev']:,.1f}M | EqVal=${bear['eq_val']:,.1f}M | Price=${bear['implied_price']:.2f} | PV_TV%={bear['pv_tv_pct']*100:.1f}%")

# Bull Case
bull = calc_dcf(
    growth_rates=[0.200, 0.170, 0.150, 0.130, 0.110],
    ebit_margins=[0.430, 0.440, 0.450, 0.455, 0.460],
    da_rates=[0.105, 0.115, 0.120, 0.120, 0.115],
    capex_rates=[0.300, 0.260, 0.230, 0.210, 0.200],
    nwc_rates=[0.010, 0.010, 0.010, 0.010, 0.010],
    wacc=0.1000,
    term_g=0.035
)
print(f"Bull Case: EV=${bull['ev']:,.1f}M | EqVal=${bull['eq_val']:,.1f}M | Price=${bull['implied_price']:.2f} | PV_TV%={bull['pv_tv_pct']*100:.1f}%")
