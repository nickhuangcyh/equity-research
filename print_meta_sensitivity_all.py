#!/usr/bin/env python3

def calc_t2_grid():
    fcfs = [31988.96, 50911.66, 68735.50, 83585.60, 93585.16]
    w = 0.11
    g = 0.03
    base_ebit_m = 0.435
    mult_list = [0.80, 0.90, 1.00, 1.10, 1.20]
    ebit_m_list = [0.395, 0.415, 0.435, 0.455, 0.475]
    net_debt = -6600.0
    shares = 2570.0
    
    print("\nTABLE 2: Growth Multiplier vs EBIT Margin")
    header = "Mult \\ EBIT\t" + "\t".join([f"{em*100:.1f}%" for em in ebit_m_list])
    print(header)
    for m in mult_list:
        row_str = f"{int(m*100)}%\t"
        for em in ebit_m_list:
            adj = m * (em / base_ebit_m)
            pv_fcf = sum([fcf * adj / ((1 + w) ** (0.5 + i)) for i, fcf in enumerate(fcfs)])
            tv = (fcfs[-1] * adj * (1 + g)) / (w - g)
            pv_tv = tv / ((1 + w) ** 4.5)
            ev = pv_fcf + pv_tv
            price = (ev - net_debt) / shares
            row_str += f"${price:.2f}\t"
        print(row_str)

def calc_t3_grid():
    fcfs = [31988.96, 50911.66, 68735.50, 83585.60, 93585.16]
    rf_list = [0.0370, 0.0420, 0.0470, 0.0520, 0.0570]
    beta_list = [1.04, 1.14, 1.24, 1.34, 1.44]
    erp = 0.0550
    kd_after_tax = 0.04368
    we = 0.9435
    wd = 0.0565
    g = 0.03
    net_debt = -6600.0
    shares = 2570.0
    
    print("\nTABLE 3: Beta vs Risk-Free Rate (Rf)")
    header = "Beta \\ Rf\t" + "\t".join([f"{rf*100:.2f}%" for rf in rf_list])
    print(header)
    for b in beta_list:
        row_str = f"{b:.2f}\t"
        for rf in rf_list:
            ke = rf + b * erp
            w = ke * we + kd_after_tax * wd
            pv_fcf = sum([fcf / ((1 + w) ** (0.5 + i)) for i, fcf in enumerate(fcfs)])
            tv = (fcfs[-1] * (1 + g)) / (w - g)
            pv_tv = tv / ((1 + w) ** 4.5)
            ev = pv_fcf + pv_tv
            price = (ev - net_debt) / shares
            row_str += f"${price:.2f}\t"
        print(row_str)

calc_t2_grid()
calc_t3_grid()
