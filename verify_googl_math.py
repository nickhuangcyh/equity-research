#!/usr/bin/env python3
"""
verify_googl_math.py - Inspect and print all DCF output metrics across Bear, Base, and Bull cases
"""
import openpyxl

def verify_googl_math():
    # Load model with evaluated data (or compute in python matching formulas)
    # Since openpyxl doesn't evaluate formulas natively without LibreOffice, let's write an exact evaluator
    # to print the table of values for our report.
    
    # 1. Inputs:
    price = 341.37
    shares = 12200.00
    market_cap = price * shares
    cash = 242500.00
    debt = 98200.00
    net_debt = debt - cash # -144300.00
    ev = market_cap + net_debt
    
    # WACC
    rf = 0.0469
    beta = 1.24
    erp = 0.0550
    ke = rf + beta * erp # 11.51%
    kd = 0.0485
    t = 0.1750
    kd_after_tax = kd * (1 - t) # 4.00125%
    
    we = market_cap / ev
    wd = net_debt / ev
    wacc = we * ke + wd * kd_after_tax
    
    print(f"Market Cap: ${market_cap:,.2f} M")
    print(f"Net Debt: ${net_debt:,.2f} M")
    print(f"Enterprise Value: ${ev:,.2f} M")
    print(f"Ke: {ke*100:.2f}%, Kd(after-tax): {kd_after_tax*100:.2f}%")
    print(f"We: {we*100:.2f}%, Wd: {wd*100:.2f}%")
    print(f"WACC: {wacc*100:.2f}%")
    
    # Scenarios
    scenarios = {
        "Bear (Case 1)": {
            "rev_growth": [0.110, 0.090, 0.075, 0.060, 0.050],
            "gm": [0.560, 0.555, 0.550, 0.545, 0.540],
            "sm": [0.080, 0.080, 0.078, 0.078, 0.078],
            "rd": [0.145, 0.145, 0.140, 0.140, 0.140],
            "ga": [0.045, 0.045, 0.042, 0.042, 0.042],
            "capex_p": [0.390, 0.300, 0.220, 0.180, 0.140],
            "term_g": 0.025,
            "wacc": wacc
        },
        "Base (Case 2)": {
            "rev_growth": [0.160, 0.140, 0.120, 0.105, 0.090],
            "gm": [0.580, 0.582, 0.585, 0.588, 0.590],
            "sm": [0.075, 0.072, 0.070, 0.068, 0.065],
            "rd": [0.138, 0.135, 0.132, 0.128, 0.125],
            "ga": [0.042, 0.040, 0.038, 0.036, 0.035],
            "capex_p": [0.380, 0.260, 0.180, 0.140, 0.110],
            "term_g": 0.030,
            "wacc": wacc
        },
        "Bull (Case 3)": {
            "rev_growth": [0.200, 0.175, 0.150, 0.130, 0.110],
            "gm": [0.590, 0.595, 0.600, 0.605, 0.610],
            "sm": [0.070, 0.065, 0.062, 0.060, 0.058],
            "rd": [0.130, 0.125, 0.120, 0.115, 0.110],
            "ga": [0.038, 0.035, 0.032, 0.030, 0.030],
            "capex_p": [0.370, 0.240, 0.160, 0.120, 0.090],
            "term_g": 0.035,
            "wacc": wacc
        }
    }
    
    rev_2025 = 402836.00
    
    for sc_name, sc in scenarios.items():
        print(f"\n=================== {sc_name} ===================")
        cur_rev = rev_2025
        revs = []
        ebits = []
        nopats = []
        das = []
        capexs = []
        nwcs = []
        ufcfs = []
        pvs = []
        
        da_rates = [0.055, 0.062, 0.068, 0.070, 0.070]
        
        for yr_idx in range(5):
            prev_rev = cur_rev
            cur_rev = prev_rev * (1 + sc["rev_growth"][yr_idx])
            revs.append(cur_rev)
            
            gp = cur_rev * sc["gm"][yr_idx]
            sm = cur_rev * sc["sm"][yr_idx]
            rd = cur_rev * sc["rd"][yr_idx]
            ga = cur_rev * sc["ga"][yr_idx]
            opex = sm + rd + ga
            ebit = gp - opex
            ebits.append(ebit)
            
            tax = ebit * t
            nopat = ebit - tax
            nopats.append(nopat)
            
            da = cur_rev * da_rates[yr_idx]
            das.append(da)
            
            capex = cur_rev * sc["capex_p"][yr_idx]
            capexs.append(capex)
            
            nwc = (cur_rev - prev_rev) * 0.02
            nwcs.append(nwc)
            
            ufcf = nopat + da - capex - nwc
            ufcfs.append(ufcf)
            
            period = yr_idx + 0.5
            df = 1 / ((1 + sc["wacc"]) ** period)
            pv = ufcf * df
            pvs.append(pv)
            
        cum_pv_fcf = sum(pvs)
        term_fcf = ufcfs[-1] * (1 + sc["term_g"])
        tv_nominal = term_fcf / (sc["wacc"] - sc["term_g"])
        tv_df = 1 / ((1 + sc["wacc"]) ** 4.5)
        pv_tv = tv_nominal * tv_df
        
        implied_ev = cum_pv_fcf + pv_tv
        implied_eq_val = implied_ev + (-net_debt) # add net cash
        implied_share_price = implied_eq_val / shares
        upside = (implied_share_price / price - 1) * 100
        
        print(f"Revenues (FY26E-FY30E): {[round(r, 1) for r in revs]}")
        print(f"EBIT (FY26E-FY30E): {[round(e, 1) for e in ebits]}")
        print(f"EBIT Margin %: {[round(e/r*100, 1) for e, r in zip(ebits, revs)]}")
        print(f"UFCF (FY26E-FY30E): {[round(f, 1) for f in ufcfs]}")
        print(f"Cum PV of 5Y FCFs: ${cum_pv_fcf:,.2f} M")
        print(f"Terminal FCF: ${term_fcf:,.2f} M | TV Nominal: ${tv_nominal:,.2f} M | PV of TV: ${pv_tv:,.2f} M")
        print(f"Implied Enterprise Value: ${implied_ev:,.2f} M")
        print(f"Net Debt Adj (Net Cash): ${-net_debt:,.2f} M")
        print(f"Implied Equity Value: ${implied_eq_val:,.2f} M")
        print(f"Implied Intrinsic Price Per Share: ${implied_share_price:.2f}")
        print(f"Current Share Price: ${price:.2f}")
        print(f"Implied Return / Upside: {upside:+.1f}%")
        print(f"PV of TV % of EV: {pv_tv/implied_ev*100:.1f}%")
        print(f"Implied Exit Multiple (EV/FY30 EBITDA): {implied_ev/(ebits[-1]+das[-1]):.2f}x")

if __name__ == "__main__":
    verify_googl_math()
