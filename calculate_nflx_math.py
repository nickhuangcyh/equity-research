#!/usr/bin/env python3
"""
calculate_nflx_math.py - Compute all exact mathematical outputs of the NFLX DCF Model across all cases and sensitivity tables.
"""

def run_dcf(case=2):
    # Market Data
    price = 77.77
    shares = 4260.00
    market_cap = price * shares
    cash = 9068.00
    debt = 14463.00
    net_debt = debt - cash
    ev_market = market_cap + net_debt
    tax_rate = 0.160
    
    # WACC Inputs
    rf = 0.0470
    beta = 1.25
    erp = 0.0550
    ke = rf + beta * erp # 0.11575
    kd_pre = 0.0525
    kd_post = kd_pre * (1 - tax_rate) # 0.0441
    
    we = market_cap / ev_market
    wd = net_debt / ev_market
    wacc_base = (ke * we) + (kd_post * wd) # ~0.1146
    
    # Historical FY25A
    rev_25 = 45183.04
    ebit_25 = 13326.60
    
    # Scenarios
    if case == 1: # Bear
        name = "Bear Case"
        rev_growths = [0.100, 0.085, 0.075, 0.065, 0.055]
        ebit_margins = [0.300, 0.302, 0.305, 0.308, 0.310]
        da_pct = 0.009
        capex_pct = 0.016
        nwc_pct = 0.020
        term_g = 0.020
        wacc = 0.1200
    elif case == 2: # Base
        name = "Base Case"
        rev_growths = [0.133, 0.120, 0.110, 0.095, 0.080]
        ebit_margins = [0.315, 0.325, 0.335, 0.340, 0.345]
        da_pct = 0.009
        capex_pct = 0.015
        nwc_pct = 0.020
        term_g = 0.030
        wacc = wacc_base
    else: # Bull
        name = "Bull Case"
        rev_growths = [0.150, 0.140, 0.130, 0.115, 0.100]
        ebit_margins = [0.325, 0.338, 0.350, 0.358, 0.365]
        da_pct = 0.009
        capex_pct = 0.014
        nwc_pct = 0.020
        term_g = 0.035
        wacc = 0.1050
        
    # Project Revenue & EBIT
    revs = []
    ebits = []
    nopats = []
    das = []
    capexs = []
    nwcs = []
    ufcfs = []
    
    prev_rev = rev_25
    for i in range(5):
        r_growth = rev_growths[i]
        curr_rev = prev_rev * (1 + r_growth)
        curr_ebit = curr_rev * ebit_margins[i]
        curr_tax = curr_ebit * tax_rate
        curr_nopat = curr_ebit - curr_tax
        curr_da = curr_rev * da_pct
        curr_capex = curr_rev * capex_pct
        curr_nwc = (curr_rev - prev_rev) * nwc_pct
        curr_ufcf = curr_nopat + curr_da - curr_capex - curr_nwc
        
        revs.append(curr_rev)
        ebits.append(curr_ebit)
        nopats.append(curr_nopat)
        das.append(curr_da)
        capexs.append(curr_capex)
        nwcs.append(curr_nwc)
        ufcfs.append(curr_ufcf)
        prev_rev = curr_rev
        
    # Discounting
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_ufcfs = [ufcfs[i] / ((1 + wacc) ** periods[i]) for i in range(5)]
    cum_pv_ufcf = sum(pv_ufcfs)
    
    # Terminal Value
    terminal_ufcf = ufcfs[-1] * (1 + term_g)
    nominal_tv = terminal_ufcf / (wacc - term_g)
    pv_tv = nominal_tv / ((1 + wacc) ** 4.5)
    
    ev_implied = cum_pv_ufcf + pv_tv
    equity_val = ev_implied - net_debt
    implied_price = equity_val / shares
    upside = (implied_price / price) - 1.0
    
    print(f"=== {name} ===")
    print(f"WACC: {wacc*100:.2f}%, Terminal g: {term_g*100:.2f}%")
    print(f"Revenues ($M): {[round(x, 1) for x in revs]}")
    print(f"EBIT ($M): {[round(x, 1) for x in ebits]}")
    print(f"NOPAT ($M): {[round(x, 1) for x in nopats]}")
    print(f"UFCF ($M): {[round(x, 1) for x in ufcfs]}")
    print(f"PV of UFCF ($M): {[round(x, 1) for x in pv_ufcfs]}")
    print(f"Cum PV of UFCF ($M): {cum_pv_ufcf:,.2f}")
    print(f"Nominal TV ($M): {nominal_tv:,.2f}")
    print(f"PV of TV ($M): {pv_tv:,.2f} ({pv_tv/ev_implied*100:.1f}% of EV)")
    print(f"Enterprise Value ($M): {ev_implied:,.2f}")
    print(f"Net Debt ($M): {net_debt:,.2f}")
    print(f"Equity Value ($M): {equity_val:,.2f}")
    print(f"Implied Price per Share: ${implied_price:.2f}")
    print(f"Current Price: ${price:.2f}")
    print(f"Implied Upside: {upside*100:+.1f}%\n")
    
    return {
        "name": name,
        "wacc": wacc,
        "term_g": term_g,
        "revs": revs,
        "ebits": ebits,
        "nopats": nopats,
        "das": das,
        "capexs": capexs,
        "nwcs": nwcs,
        "ufcfs": ufcfs,
        "pv_ufcfs": pv_ufcfs,
        "cum_pv_ufcf": cum_pv_ufcf,
        "nominal_tv": nominal_tv,
        "pv_tv": pv_tv,
        "ev": ev_implied,
        "net_debt": net_debt,
        "equity_val": equity_val,
        "price": implied_price,
        "upside": upside,
    }

def run_sensitivity_tables():
    base_res = run_dcf(2)
    price = 77.77
    shares = 4260.00
    net_debt = 5395.00
    ufcfs = base_res["ufcfs"]
    
    print("=== SENSITIVITY TABLE 1: WACC vs Terminal g ===")
    g_cols = [0.020, 0.025, 0.030, 0.035, 0.040]
    wacc_rows = [0.1046, 0.1096, 0.1146, 0.1196, 0.1246]
    
    header = "WACC \\ g\t" + "\t".join([f"{g*100:.1f}%" for g in g_cols])
    print(header)
    for w in wacc_rows:
        row_str = f"{w*100:.2f}%\t"
        for g in g_cols:
            pv_f = sum([ufcfs[i] / ((1 + w) ** (i + 0.5)) for i in range(5)])
            tv = (ufcfs[-1] * (1 + g)) / (w - g)
            pv_tv = tv / ((1 + w) ** 4.5)
            ev = pv_f + pv_tv
            eq = ev - net_debt
            p = eq / shares
            row_str += f"${p:.2f}\t"
        print(row_str)
        
    print("\n=== SENSITIVITY TABLE 2: Growth Multiplier vs FY30E EBIT Margin ===")
    ebit_cols = [0.315, 0.330, 0.345, 0.360, 0.375]
    growth_rows = [0.80, 0.90, 1.00, 1.10, 1.20]
    w = base_res["wacc"]
    g = base_res["term_g"]
    base_ebit_target = 0.345
    
    header = "Growth \\ Margin\t" + "\t".join([f"{em*100:.1f}%" for em in ebit_cols])
    print(header)
    for gr in growth_rows:
        row_str = f"{int(gr*100)}%\t"
        for em in ebit_cols:
            scaled_ufcfs = [ufcfs[i] * gr * (em / base_ebit_target) for i in range(5)]
            pv_f = sum([scaled_ufcfs[i] / ((1 + w) ** (i + 0.5)) for i in range(5)])
            tv = (scaled_ufcfs[-1] * (1 + g)) / (w - g)
            pv_tv = tv / ((1 + w) ** 4.5)
            ev = pv_f + pv_tv
            eq = ev - net_debt
            p = eq / shares
            row_str += f"${p:.2f}\t"
        print(row_str)
        
    print("\n=== SENSITIVITY TABLE 3: Beta vs Risk-Free Rate ===")
    rf_cols = [0.0370, 0.0420, 0.0470, 0.0520, 0.0570]
    beta_rows = [1.05, 1.15, 1.25, 1.35, 1.45]
    erp = 0.0550
    kd_post = 0.0441
    we = (77.77 * 4260.0) / (77.77 * 4260.0 + 5395.0)
    wd = 5395.0 / (77.77 * 4260.0 + 5395.0)
    
    header = "Beta \\ Rf\t" + "\t".join([f"{rf*100:.2f}%" for rf in rf_cols])
    print(header)
    for b in beta_rows:
        row_str = f"{b:.2f}\t"
        for rf in rf_cols:
            ke_curr = rf + b * erp
            w_curr = (ke_curr * we) + (kd_post * wd)
            pv_f = sum([ufcfs[i] / ((1 + w_curr) ** (i + 0.5)) for i in range(5)])
            tv = (ufcfs[-1] * (1 + g)) / (w_curr - g)
            pv_tv = tv / ((1 + w_curr) ** 4.5)
            ev = pv_f + pv_tv
            eq = ev - net_debt
            p = eq / shares
            row_str += f"${p:.2f}\t"
        print(row_str)

if __name__ == "__main__":
    print("--- RUNNING 3 SCENARIOS ---")
    run_dcf(1)
    run_dcf(2)
    run_dcf(3)
    print("--- RUNNING SENSITIVITY MATRICES ---")
    run_sensitivity_tables()
