beta_axis = [1.80, 2.00, 2.22, 2.40, 2.60]
rf_axis = [0.0420, 0.0445, 0.0470, 0.0495, 0.0520]
fcffs = [196840.6, 283438.5, 353344.1, 400925.1, 433923.8]
periods = [0.5, 1.5, 2.5, 3.5, 4.5]
net_debt = -54100.0
shares_out = 24300.0
base_term_g = 0.030

print("=== TABLE 3: Adjusted CAPM Center Check ===")
for b in beta_axis:
    row_strs = []
    for rf in rf_axis:
        w = rf + b * 0.055 - 0.0541
        dfs = [1 / ((1 + w) ** p) for p in periods]
        pv_f = sum(f * d for f, d in zip(fcffs, dfs))
        tv = (fcffs[-1] * (1 + base_term_g)) / (w - base_term_g)
        pv_t = tv / ((1 + w) ** periods[-1])
        ev = pv_f + pv_t
        eq = ev - net_debt
        price = eq / shares_out
        row_strs.append(f"${price:6.2f}")
    print(f"{b:<6.2f} | " + " | ".join(row_strs))
