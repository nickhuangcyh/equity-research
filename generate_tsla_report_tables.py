#!/usr/bin/env python3
"""
generate_tsla_report_tables.py - Compute all sensitivity table cells for TSLA report
"""
import openpyxl

def generate_tables():
    wb = openpyxl.load_workbook("TSLA_DCF_Model_Gemini-3.7-Flash_20260819.xlsx", data_only=False)
    # We can evaluate using the formula logic directly
    stock_p = 336.50
    shares = 3760.00
    cash = 44059.00
    debt = 8376.00
    net_debt_adj = cash - debt  # +35683.00
    
    # Base Case UFCFs
    rev = [94827.0]
    rev_g = [0.110, 0.150, 0.160, 0.140, 0.120]
    for g in rev_g:
        rev.append(rev[-1] * (1 + g))
        
    ebit_m = [0.060, 0.080, 0.105, 0.122, 0.138]
    ebit = [rev[i] * ebit_m[i-1] for i in range(1, 6)]
    tax = [max(0, e * 0.21) for e in ebit]
    nopat = [e - t for e, t in zip(ebit, tax)]
    da = [rev[i] * [0.058, 0.055, 0.052, 0.050, 0.048][i-1] for i in range(1, 6)]
    capex = [rev[i] * [0.085, 0.080, 0.075, 0.070, 0.065][i-1] for i in range(1, 6)]
    nwc = [(rev[i] - rev[i-1]) * 0.025 for i in range(1, 6)]
    ufcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, nwc)]
    
    wacc_base = 0.149594
    erp = 0.0550
    we = 1.02894
    wd = -0.02894
    kd_after = 0.0550 * (1 - 0.21)
    
    periods = [0.5, 1.5, 2.5, 3.5, 4.5]
    
    def calc_dcf_price(ufcfs, wacc, g_term):
        df = [1 / ((1 + wacc)**p) for p in periods]
        pv_ufcf = sum(u * d for u, d in zip(ufcfs, df))
        tv = (ufcfs[-1] * (1 + g_term)) / (wacc - g_term)
        pv_tv = tv / ((1 + wacc)**4.5)
        ev = pv_ufcf + pv_tv
        eq_val = ev + net_debt_adj
        return eq_val / shares

    print("### TABLE 1: WACC vs. Terminal Growth Rate (g)")
    g_cols = [0.020, 0.025, 0.030, 0.035, 0.040]
    wacc_rows = [0.1296, 0.1396, 0.1496, 0.1596, 0.1696]
    header = "| WACC \\ 終端成長率 (g) | 2.0% | 2.5% | **3.0% (Base)** | 3.5% | 4.0% |"
    sep = "| :---: | :---: | :---: | :---: | :---: | :---: |"
    print(header)
    print(sep)
    for w in wacc_rows:
        row_str = f"| **{w*100:.2f}%**" + (" (Base)" if abs(w-0.1496)<0.0001 else "") + " |"
        for g in g_cols:
            p = calc_dcf_price(ufcf, w, g)
            if abs(w-0.1496)<0.0001 and abs(g-0.030)<0.0001:
                row_str += f" **${p:.2f}** |"
            else:
                row_str += f" ${p:.2f} |"
        print(row_str)

    print("\n### TABLE 2: Revenue Growth Multiplier vs. Target FY30E EBIT Margin")
    ebit_cols = [0.098, 0.118, 0.138, 0.158, 0.178]
    growth_rows = [0.80, 0.90, 1.00, 1.10, 1.20]
    header2 = "| 成長係數 \\ 目標 EBIT Margin | 9.8% | 11.8% | **13.8% (Base)** | 15.8% | 17.8% |"
    print(header2)
    print(sep)
    for gr in growth_rows:
        row_str = f"| **{gr:.2f}x**" + (" (Base)" if gr==1.0 else "") + " |"
        for em in ebit_cols:
            scale = gr * (em / 0.138)
            scaled_ufcf = [u * scale for u in ufcf]
            p = calc_dcf_price(scaled_ufcf, wacc_base, 0.030)
            if gr==1.0 and abs(em-0.138)<0.001:
                row_str += f" **${p:.2f}** |"
            else:
                row_str += f" ${p:.2f} |"
        print(row_str)

    print("\n### TABLE 3: Beta vs. Risk-Free Rate (Rf)")
    rf_cols = [0.0365, 0.0415, 0.0465, 0.0515, 0.0565]
    beta_rows = [1.42, 1.62, 1.82, 2.02, 2.22]
    header3 = "| 貝塔值 \\ 無風險利率 (Rf) | 3.65% | 4.15% | **4.65% (Base)** | 5.15% | 5.65% |"
    print(header3)
    print(sep)
    for b in beta_rows:
        row_str = f"| **{b:.2f}**" + (" (Base)" if b==1.82 else "") + " |"
        for rf in rf_cols:
            wacc_dyn = ((rf + b * erp) * we) + (kd_after * wd)
            p = calc_dcf_price(ufcf, wacc_dyn, 0.030)
            if b==1.82 and abs(rf-0.0465)<0.0001:
                row_str += f" **${p:.2f}** |"
            else:
                row_str += f" ${p:.2f} |"
        print(row_str)

if __name__ == "__main__":
    generate_tables()
