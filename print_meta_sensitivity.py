#!/usr/bin/env python3

def calc_t1_grid():
    fcfs = [31988.96, 50911.66, 68735.50, 83585.60, 93585.16]
    wacc_list = [0.09, 0.10, 0.11, 0.12, 0.13]
    g_list = [0.020, 0.025, 0.030, 0.035, 0.040]
    net_debt = -6600.0
    shares = 2570.0
    
    print("TABLE 1: WACC vs Terminal Growth Rate (g)")
    header = "WACC \\ g\t" + "\t".join([f"{g*100:.1f}%" for g in g_list])
    print(header)
    for w in wacc_list:
        row_str = f"{w*100:.1f}%\t"
        for g in g_list:
            pv_fcf = sum([fcf / ((1 + w) ** (0.5 + i)) for i, fcf in enumerate(fcfs)])
            tv = (fcfs[-1] * (1 + g)) / (w - g)
            pv_tv = tv / ((1 + w) ** 4.5)
            ev = pv_fcf + pv_tv
            price = (ev - net_debt) / shares
            row_str += f"${price:.2f}\t"
        print(row_str)

calc_t1_grid()
