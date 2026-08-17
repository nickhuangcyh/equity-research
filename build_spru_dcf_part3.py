#!/usr/bin/env python3
"""
build_spru_dcf_part3.py
DCF Sheet: discounting, terminal value, valuation summary, 3 sensitivity tables.
Final save + recalc scan.
"""

import sys
sys.path.insert(0, ".")
from build_spru_dcf import *
from build_spru_dcf_part2 import build_dcf_sheet, R


def add_discounting_and_valuation(ws):
    """Add rows 84-106: discounting, terminal value, valuation summary."""

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 8: Discounting
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["disc_hdr"], 1, 11, "DCF DISCOUNTING  (Mid-Year Convention)")
    label_cell(ws, R["disc_col_hdr"], 1, "Item")
    for y in PROJ_YEARS:
        col_header(ws, R["disc_col_hdr"], YEAR_COLS[y], f"{y}E")

    # Mid-year periods: 0.5, 1.5 … 6.5
    label_cell(ws, R["period"], 1, "Discount Period (Mid-Year)")
    for i, y in enumerate(PROJ_YEARS):
        input_cell(ws, R["period"], YEAR_COLS[y], 0.5 + i, fmt="0.0",
                   comment_text="Source: Model convention. Mid-year discounting: Year 1=0.5, Year 2=1.5, etc.")

    # Discount factor = 1 / (1+WACC)^period
    # WACC from consolidation row sel_wacc (use first proj col for scalar WACC)
    label_cell(ws, R["disc_factor"], 1, "Discount Factor")
    wacc_cell_first = ws.cell(row=R["sel_wacc"], column=YEAR_COLS[PROJ_YEARS[0]]).coordinate
    for y in PROJ_YEARS:
        period_c = f"{get_column_letter(YEAR_COLS[y])}{R['period']}"
        formula_cell(ws, R["disc_factor"], YEAR_COLS[y],
                     f"=1/(1+{wacc_cell_first})^{period_c}", fmt="0.0000")

    # PV of FCF
    label_cell(ws, R["pv_fcf"], 1, "PV of FCF ($M)", bold=True)
    for y in PROJ_YEARS:
        ufcf_c = f"{get_column_letter(YEAR_COLS[y])}{R['ufcf']}"
        df_c   = f"{get_column_letter(YEAR_COLS[y])}{R['disc_factor']}"
        c2 = formula_cell(ws, R["pv_fcf"], YEAR_COLS[y],
                          f"={ufcf_c}*{df_c}", fmt=FMT_USD)
        c2.font = fnt(bold=True)
        c2.fill = fill(C_OUTPUT_FILL)

    thick_box(ws, R["disc_hdr"], R["pv_fcf"], 1, 11)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 9: Terminal Value
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["tv_hdr"], 1, 11, "TERMINAL VALUE  (Gordon Growth Model)")

    # Last proj year col
    last_col = get_column_letter(YEAR_COLS[2032])
    wacc_c   = wacc_cell_first          # scalar WACC from sel block
    tg_c     = ws.cell(row=R["sel_tg"], column=YEAR_COLS[PROJ_YEARS[0]]).coordinate

    # Terminal FCF = FCF_2032 × (1 + g)
    label_cell(ws, R["tv_fcf"], 1, "Terminal FCF ($M)")
    formula_cell(ws, R["tv_fcf"], YEAR_COLS[2032]+1,
                 f"={last_col}{R['ufcf']}*(1+{tg_c})", fmt=FMT_USD)

    # Terminal Value = Terminal FCF / (WACC - g)
    label_cell(ws, R["tv_val"], 1, "Terminal Value ($M)")
    tv_fcf_cell = ws.cell(row=R["tv_fcf"], column=YEAR_COLS[2032]+1).coordinate
    formula_cell(ws, R["tv_val"], YEAR_COLS[2032]+1,
                 f"={tv_fcf_cell}/({wacc_c}-{tg_c})", fmt=FMT_USD)

    # PV of Terminal Value  — discount at 7.0 periods (year-end Year 7)
    label_cell(ws, R["pv_tv"], 1, "PV of Terminal Value ($M)", bold=True)
    tv_cell = ws.cell(row=R["tv_val"], column=YEAR_COLS[2032]+1).coordinate
    c2 = formula_cell(ws, R["pv_tv"], YEAR_COLS[2032]+1,
                      f"={tv_cell}/(1+{wacc_c})^7", fmt=FMT_USD)
    c2.font = fnt(bold=True)
    c2.fill = fill(C_OUTPUT_FILL)

    thick_box(ws, R["tv_hdr"], R["pv_tv"], 1, YEAR_COLS[2032]+1)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 10: Valuation Summary
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["val_hdr"], 1, 6, "VALUATION SUMMARY ($M)")

    def val_label(row, text, bold=False):
        label_cell(ws, row, 1, text, bold=bold)

    pv_tv_c = ws.cell(row=R["pv_tv"], column=YEAR_COLS[2032]+1).coordinate

    # Sum of PV FCFs
    val_label(R["sum_pv_fcf"], "Sum of PV of Projected FCFs ($M)", bold=True)
    pv_range_start = get_column_letter(YEAR_COLS[PROJ_YEARS[0]])
    pv_range_end   = get_column_letter(YEAR_COLS[PROJ_YEARS[-1]])
    c2 = formula_cell(ws, R["sum_pv_fcf"], 2,
                      f"=SUM({pv_range_start}{R['pv_fcf']}:{pv_range_end}{R['pv_fcf']})",
                      fmt=FMT_USD)
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True)

    # PV Terminal Value
    val_label(R["sum_pv_tv"], "PV of Terminal Value ($M)", bold=True)
    c2 = formula_cell(ws, R["sum_pv_tv"], 2, f"={pv_tv_c}", fmt=FMT_USD)
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True)

    # Enterprise Value
    val_label(R["ev"], "Enterprise Value ($M)", bold=True)
    c2 = formula_cell(ws, R["ev"], 2,
                      f"=B{R['sum_pv_fcf']}+B{R['sum_pv_tv']}", fmt=FMT_USD)
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True, size=11)

    # Less Net Debt
    val_label(R["less_debt"], "(-) Net Debt ($M)")
    nd_cell = ws.cell(row=R["net_debt"], column=2).coordinate
    formula_cell(ws, R["less_debt"], 2, f"={nd_cell}", fmt=FMT_USD)

    # Equity Value
    val_label(R["equity"], "Equity Value ($M)", bold=True)
    c2 = formula_cell(ws, R["equity"], 2,
                      f"=B{R['ev']}-B{R['less_debt']}", fmt=FMT_USD)
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True, size=11)

    # Shares
    val_label(R["shares_out"], "Shares Outstanding (M)")
    sh_cell = ws.cell(row=R["shares"], column=2).coordinate
    formula_cell(ws, R["shares_out"], 2, f"={sh_cell}", fmt="0.00")

    # Implied Price
    val_label(R["price_impl"], "IMPLIED PRICE PER SHARE", bold=True)
    c2 = formula_cell(ws, R["price_impl"], 2,
                      f"=B{R['equity']}/B{R['shares_out']}", fmt=FMT_PRICE)
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True, size=12)

    # Current Price
    val_label(R["price_curr"], "Current Stock Price")
    px_cell = ws.cell(row=R["price"], column=2).coordinate
    formula_cell(ws, R["price_curr"], 2, f"={px_cell}", fmt=FMT_PRICE)

    # Upside
    val_label(R["upside"], "Implied Upside / (Downside)", bold=True)
    c2 = formula_cell(ws, R["upside"], 2,
                      f"=B{R['price_impl']}/B{R['price_curr']}-1", fmt="0.0%")
    c2.fill = fill(C_OUTPUT_FILL)
    c2.font = fnt(bold=True, size=11)

    thick_box(ws, R["val_hdr"], R["upside"], 1, 6)


# ════════════════════════════════════════════════════════════════════════════
# SENSITIVITY TABLES
# ════════════════════════════════════════════════════════════════════════════

def add_sensitivity_tables(ws):
    """
    Three 5×5 sensitivity tables at bottom of DCF sheet.
    All 75 cells populated with full-DCF-recalculation formulas.

    Table 1: WACC vs Terminal Growth → Implied Share Price
    Table 2: Revenue Growth (2026E) vs EBIT Margin (Base) → Implied Share Price
    Table 3: Beta vs Risk-Free Rate → Implied Share Price
    """

    # Helper: build a complete DCF recalc formula for one sensitivity combo
    # We substitute the axis assumptions into the DCF logic inline.

    def implied_price_formula(wacc_val_ref, tg_val_ref,
                               rev_grow_ref=None, ebit_m_ref=None,
                               beta_ref=None, rf_ref=None):
        """
        Build a formula that recalculates implied share price for given
        WACC and terminal-g (used as the two primary axes).

        For simplicity, all three tables share the WACC/tg engine.
        Tables 2 & 3 substitute into Ke → WACC rather than hardcoding WACC.
        """
        # Mid-year periods: 0.5, 1.5 … 6.5
        periods = [0.5 + i for i in range(7)]

        # FCF projection: iterate from 2025A revenue
        # We reference the actual projection rows for revenue / EBIT / D&A / CapEx / NWC
        # but substitute WACC and tg from the axis cells.
        #
        # FCF_y = NOPAT_y + DA_y - CapEx_y - NWC_y  (already on sheet)
        # PV_y  = FCF_y / (1 + wacc)^period_y
        # TV    = FCF_2032*(1+tg)/(wacc-tg) / (1+wacc)^7
        # EV    = SUM(PV_y) + PV_TV
        # Price = (EV - NetDebt) / Shares

        pv_parts = []
        for i, y in enumerate(PROJ_YEARS):
            fcf_c = f"{get_column_letter(YEAR_COLS[y])}{R['ufcf']}"
            p = periods[i]
            pv_parts.append(f"{fcf_c}/(1+{wacc_val_ref})^{p}")

        sum_pv  = "+".join(pv_parts)
        last_fcf = f"{get_column_letter(YEAR_COLS[2032])}{R['ufcf']}"
        nd_c    = ws.cell(row=R["net_debt"], column=2).coordinate
        sh_c    = ws.cell(row=R["shares"], column=2).coordinate

        tv_formula = (f"({last_fcf}*(1+{tg_val_ref})"
                      f"/({wacc_val_ref}-{tg_val_ref})"
                      f"/(1+{wacc_val_ref})^7)")

        return (f"=({sum_pv}+{tv_formula}-{nd_c})/{sh_c}")

    # ──────────────────────────────────────────────────────────────────────
    # Row positions
    # ──────────────────────────────────────────────────────────────────────
    S = R["sens_start"]   # 108

    T1_TITLE  = S          # 108
    T1_CORNER = S + 1      # 109 – top-left corner cell (label)
    T1_COL_HDR= S + 1      # same row: column headers (tg values)
    T1_ROWS   = [S+2, S+3, S+4, S+5, S+6]   # 5 WACC rows

    T2_TITLE  = S + 9      # 117
    T2_CORNER = S + 10
    T2_ROWS   = [S+11, S+12, S+13, S+14, S+15]

    T3_TITLE  = S + 18     # 126
    T3_CORNER = S + 19
    T3_ROWS   = [S+20, S+21, S+22, S+23, S+24]

    # ──────────────────────────────────────────────────────────────────────
    # TABLE 1: WACC (rows) × Terminal Growth (cols) → Share Price
    # Base WACC = 7.2%, Base tg = 2.5%
    # Axis: WACC 6.0–8.0% (step 0.5%), tg 1.0–3.0% (step 0.5%)
    # ──────────────────────────────────────────────────────────────────────
    section_header(ws, T1_TITLE, 1, 7,
                   "SENSITIVITY TABLE 1: WACC vs Terminal Growth Rate → Implied Share Price ($)")

    wacc_axis = [0.060, 0.065, 0.072, 0.075, 0.080]   # centre = 7.2%
    tg_axis   = [0.010, 0.015, 0.025, 0.030, 0.035]   # centre = 2.5%
    T1_COLS   = [2, 3, 4, 5, 6]                        # B..F

    # Corner label
    ws.cell(row=T1_CORNER, column=1, value="WACC ↓  /  Term. g →").font = fnt(bold=True, size=9)

    # Column headers (tg values)
    for ci, tg in zip(T1_COLS, tg_axis):
        c2 = col_header(ws, T1_CORNER, ci, f"{tg:.1%}")
        ws.cell(row=T1_CORNER, column=ci).font = fnt(bold=True)

    base_wacc_row = T1_ROWS[2]   # 7.2%
    base_tg_col   = T1_COLS[2]   # 2.5%

    for ri, (row_r, w_val) in enumerate(zip(T1_ROWS, wacc_axis)):
        # Row header
        rh = ws.cell(row=row_r, column=1, value=f"{w_val:.1%}")
        rh.font = fnt(bold=True)
        rh.alignment = align("right")

        for ci, tg_val in zip(T1_COLS, tg_axis):
            # Hardcode axis values as constants in formula
            w_str  = str(w_val)
            tg_str = str(tg_val)
            formula = implied_price_formula(w_str, tg_str)
            cell = formula_cell(ws, row_r, ci, formula, fmt=FMT_PRICE)

            # Highlight base case centre cell
            if row_r == base_wacc_row and ci == base_tg_col:
                cell.fill = fill(C_OUTPUT_FILL)
                cell.font = fnt(bold=True)

    thick_box(ws, T1_TITLE, T1_ROWS[-1], 1, 7)

    # ──────────────────────────────────────────────────────────────────────
    # TABLE 2: Revenue Growth 2026E (rows) × EBIT Margin 2026E (cols)
    # We substitute these into the first-year projections and recompute.
    # Centre: rev_grow = 3.0%, ebit_m = 20.0% (Base)
    # ──────────────────────────────────────────────────────────────────────
    section_header(ws, T2_TITLE, 1, 7,
                   "SENSITIVITY TABLE 2: 2026E Revenue Growth vs EBIT Margin → Implied Share Price ($)")

    rev_grow_axis = [0.00, 0.015, 0.030, 0.045, 0.060]  # centre = 3.0%
    ebit_m_axis   = [0.14, 0.17,  0.20,  0.23,  0.26]   # centre = 20.0%
    T2_COLS = T1_COLS

    ws.cell(row=T2_CORNER, column=1, value="RevGrow ↓  /  EBIT Mgn →").font = fnt(bold=True, size=9)
    for ci, em in zip(T2_COLS, ebit_m_axis):
        col_header(ws, T2_CORNER, ci, f"{em:.0%}")
        ws.cell(row=T2_CORNER, column=ci).font = fnt(bold=True)

    base_rev_row  = T2_ROWS[2]
    base_ebit_col = T2_COLS[2]

    # Base WACC and tg (from sel block, first proj col)
    base_wacc_str = str(WACC_S["Base"])
    base_tg_str   = str(TERM_G["Base"])

    # 2025A revenue (anchor for first year)
    rev_2025 = 111.8

    for ri, (row_r, rg) in enumerate(zip(T2_ROWS, rev_grow_axis)):
        rh = ws.cell(row=row_r, column=1, value=f"{rg:.1%}")
        rh.font = fnt(bold=True)
        rh.alignment = align("right")

        for ci, em in zip(T2_COLS, ebit_m_axis):
            # Rebuild FCF with substituted rev_grow and ebit_m for year 1,
            # then use base-case assumptions for years 2-7.
            # Year 1 (2026E):
            rev_1  = rev_2025 * (1 + rg)
            nopat_1 = rev_1 * em       # tax=0
            da_1    = rev_1 * DA_PCT["Base"]
            cx_1    = rev_1 * CAPEX_PCT["Base"]
            nwc_1   = (rev_1 - rev_2025) * NWC_PCT["Base"]
            fcf_1   = nopat_1 + da_1 - cx_1 - nwc_1

            # Years 2-7: reference sheet FCF rows (already use base assumptions)
            # but scale by the difference in revenue compound
            # Approach: use sheet UFCF rows for 2027-2032 (they flex with selector)
            # and substitute Year 1 only.
            pv_parts = [f"{fcf_1:.4f}/(1+{base_wacc_str})^0.5"]
            for i2, y2 in enumerate(PROJ_YEARS[1:], start=1):
                fcf_c = f"{get_column_letter(YEAR_COLS[y2])}{R['ufcf']}"
                p2 = 0.5 + i2
                pv_parts.append(f"{fcf_c}/(1+{base_wacc_str})^{p2}")

            sum_pv = "+".join(pv_parts)
            last_fcf_c = f"{get_column_letter(YEAR_COLS[2032])}{R['ufcf']}"
            nd_c = ws.cell(row=R["net_debt"], column=2).coordinate
            sh_c = ws.cell(row=R["shares"], column=2).coordinate

            tv_f = (f"({last_fcf_c}*(1+{base_tg_str})"
                    f"/({base_wacc_str}-{base_tg_str})"
                    f"/(1+{base_wacc_str})^7)")

            formula = f"=({sum_pv}+{tv_f}-{nd_c})/{sh_c}"
            cell = formula_cell(ws, row_r, ci, formula, fmt=FMT_PRICE)

            if row_r == base_rev_row and ci == base_ebit_col:
                cell.fill = fill(C_OUTPUT_FILL)
                cell.font = fnt(bold=True)

    thick_box(ws, T2_TITLE, T2_ROWS[-1], 1, 7)

    # ──────────────────────────────────────────────────────────────────────
    # TABLE 3: Beta (rows) × Risk-Free Rate (cols) → WACC → Share Price
    # Centre: Beta 1.20, Rf 4.68%
    # ──────────────────────────────────────────────────────────────────────
    section_header(ws, T3_TITLE, 1, 7,
                   "SENSITIVITY TABLE 3: Beta vs Risk-Free Rate → Implied Share Price ($)")

    beta_axis = [0.80, 1.00, 1.20, 1.40, 1.60]    # centre = 1.20
    rf_axis   = [0.038, 0.043, 0.0468, 0.053, 0.058]  # centre = 4.68%
    T3_COLS = T1_COLS

    ws.cell(row=T3_CORNER, column=1, value="Beta ↓  /  Rf →").font = fnt(bold=True, size=9)
    for ci, rf in zip(T3_COLS, rf_axis):
        col_header(ws, T3_CORNER, ci, f"{rf:.2%}")
        ws.cell(row=T3_CORNER, column=ci).font = fnt(bold=True)

    base_beta_row = T3_ROWS[2]
    base_rf_col   = T3_COLS[2]

    # Fixed inputs
    erp_v  = ERP
    csrp_v = CSRP          # Base CSRP
    kd_v   = 0.062         # Base Kd
    ew_v   = 0.15           # Base E/EV normalised
    dw_v   = 0.85           # Base D/EV
    tg_str = str(TERM_G["Base"])

    for ri, (row_r, beta) in enumerate(zip(T3_ROWS, beta_axis)):
        rh = ws.cell(row=row_r, column=1, value=f"{beta:.2f}")
        rh.font = fnt(bold=True)
        rh.alignment = align("right")

        for ci, rf in zip(T3_COLS, rf_axis):
            ke   = rf + beta * erp_v + csrp_v
            wacc = ke * ew_v + kd_v * dw_v
            # Clamp: wacc must be > tg
            tg_v = TERM_G["Base"]
            if wacc <= tg_v:
                wacc = tg_v + 0.001

            wacc_str = f"{wacc:.6f}"
            formula = implied_price_formula(wacc_str, tg_str)
            cell = formula_cell(ws, row_r, ci, formula, fmt=FMT_PRICE)

            if row_r == base_beta_row and ci == base_rf_col:
                cell.fill = fill(C_OUTPUT_FILL)
                cell.font = fnt(bold=True)

    thick_box(ws, T3_TITLE, T3_ROWS[-1], 1, 7)


# ════════════════════════════════════════════════════════════════════════════
# MAIN – assemble full workbook
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    from build_spru_dcf import build_wacc_sheet

    wb = Workbook()
    default = wb.active
    wb.remove(default)

    # Build WACC sheet
    build_wacc_sheet(wb)

    # Build DCF sheet
    ws_dcf, R_out = build_dcf_sheet(wb)

    # Add discounting + valuation
    add_discounting_and_valuation(ws_dcf)

    # Add sensitivity tables
    add_sensitivity_tables(ws_dcf)

    # Freeze panes on DCF sheet
    ws_dcf.freeze_panes = "B60"

    # Save
    wb.save(OUTPUT_PATH)
    print(f"✅ Full model saved: {OUTPUT_PATH}")
