#!/usr/bin/env python3
"""
verify_elf_model.py
Evaluates Excel formulas and verifies sensitivity table center cell sanity.
"""

def verify():
    # Base inputs:
    rev_26 = 1636.50
    growth_base = [0.160, 0.135, 0.110, 0.090, 0.070]
    ebit_m_base = [0.100, 0.120, 0.135, 0.145, 0.150]
    da_pct = [0.040, 0.036, 0.032, 0.030, 0.028]
    capex_pct = [0.015, 0.015, 0.015, 0.015, 0.015]
    nwc_pct = [0.020, 0.020, 0.020, 0.020, 0.020]
    tax_rate = 0.25
    rf = 0.0473
    beta = 1.56
    erp = 0.0550
    kd = 0.0650
    price = 91.54
    shares = 59.06
    cash = 289.70
    debt = 841.70
    net_debt = debt - cash
    mcap = price * shares
    tot_cap = mcap + debt
    we = mcap / tot_cap
    wd = debt / tot_cap
    ke = rf + beta * erp
    kd_after = kd * (1 - tax_rate)
    wacc = (ke * we) + (kd_after * wd)
    g_base = 0.030

    print(f"Ke: {ke:.4%}, WACC: {wacc:.4%}, Net Debt: ${net_debt:.2f}M")

    # Projections
    revs = []
    ebits = []
    nopats = []
    das = []
    capexs = []
    nwcs = []
    ufcfs = []

    curr_rev = rev_26
    for i in range(5):
        nxt_rev = curr_rev * (1 + growth_base[i])
        revs.append(nxt_rev)
        ebit = nxt_rev * ebit_m_base[i]
        ebits.append(ebit)
        nopat = ebit * (1 - tax_rate)
        nopats.append(nopat)
        da = nxt_rev * da_pct[i]
        das.append(da)
        capex = nxt_rev * capex_pct[i]
        capexs.append(capex)
        dnwc = (nxt_rev - curr_rev) * nwc_pct[i]
        nwcs.append(dnwc)
        ufcf = nopat + da - capex - dnwc
        ufcfs.append(ufcf)
        curr_rev = nxt_rev

    for yr, r, eb, f in zip(range(2027, 2032), revs, ebits, ufcfs):
        print(f"FY{yr}E: Rev=${r:.1f}M, EBIT=${eb:.1f}M, UFCF=${f:.1f}M")

    # PV
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    pv_fcfs = [f / ((1 + wacc) ** p) for f, p in zip(ufcfs, periods)]
    sum_pv_fcf = sum(pv_fcfs)
    print(f"Sum PV(FCF): ${sum_pv_fcf:.2f}M")

    tv = (ufcfs[-1] * (1 + g_base)) / (wacc - g_base)
    pv_tv = tv / ((1 + wacc) ** 5)
    ev = sum_pv_fcf + pv_tv
    eq_val = ev - net_debt
    implied_price = eq_val / shares
    print(f"TV: ${tv:.2f}M, PV(TV): ${pv_tv:.2f}M, EV: ${ev:.2f}M, EqVal: ${eq_val:.2f}M")
    print(f"Implied Price per share: ${implied_price:.2f}")

if __name__ == "__main__":
    verify()
