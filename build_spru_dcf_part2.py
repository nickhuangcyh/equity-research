#!/usr/bin/env python3
"""
build_spru_dcf_part2.py
DCF Sheet: header, case selector, market data, assumption blocks, 
revenue projection, FCF build.
Appended into the workbook created by Part 1.
"""

import sys
sys.path.insert(0, ".")
from build_spru_dcf import *   # all helpers and constants

# ════════════════════════════════════════════════════════════════════════════
# ROW MAP  – define ALL rows before writing any formulas
# ════════════════════════════════════════════════════════════════════════════
R = {}

# Header block
R["title"]       = 1
R["subtitle"]    = 2
R["blank1"]      = 3
R["selector_lbl"]= 4
R["selector"]    = 5   # B5 = case selector (1/2/3)
R["case_name"]   = 6
R["blank2"]      = 7

# Market data
R["mkt_hdr"]     = 8
R["mkt_col_hdr"] = 9
R["price"]       = 10
R["shares"]      = 11
R["mktcap"]      = 12
R["net_debt"]    = 13
R["blank3"]      = 14

# Assumption blocks header
R["assum_hdr"]   = 15
R["assum_col_hdr"]= 16

# Bear block
R["bear_hdr"]    = 17
R["bear_yr_hdr"] = 18
R["bear_grow"]   = 19
R["bear_ebit"]   = 20
R["bear_da"]     = 21
R["bear_capex"]  = 22
R["bear_nwc"]    = 23
R["bear_tax"]    = 24
R["bear_tg"]     = 25
R["bear_wacc"]   = 26

# Base block
R["base_hdr"]    = 27
R["base_yr_hdr"] = 28
R["base_grow"]   = 29
R["base_ebit"]   = 30
R["base_da"]     = 31
R["base_capex"]  = 32
R["base_nwc"]    = 33
R["base_tax"]    = 34
R["base_tg"]     = 35
R["base_wacc"]   = 36

# Bull block
R["bull_hdr"]    = 37
R["bull_yr_hdr"] = 38
R["bull_grow"]   = 39
R["bull_ebit"]   = 40
R["bull_da"]     = 41
R["bull_capex"]  = 42
R["bull_nwc"]    = 43
R["bull_tax"]    = 44
R["bull_tg"]     = 45
R["bull_wacc"]   = 46

# Consolidation (Selected) block
R["sel_hdr"]     = 47
R["sel_yr_hdr"]  = 48
R["sel_grow"]    = 49
R["sel_ebit"]    = 50
R["sel_da"]      = 51
R["sel_capex"]   = 52
R["sel_nwc"]     = 53
R["sel_tax"]     = 54
R["sel_tg"]      = 55
R["sel_wacc"]    = 56

R["blank4"]      = 57

# Revenue & EBIT projection
R["is_hdr"]      = 58
R["is_col_hdr"]  = 59
R["rev"]         = 60
R["rev_grow"]    = 61
R["blank5"]      = 62
R["ebit"]        = 63
R["ebit_m"]      = 64
R["blank6"]      = 65
R["tax"]         = 66
R["tax_r"]       = 67
R["blank7"]      = 68
R["nopat"]       = 69
R["blank8"]      = 70

# FCF build
R["fcf_hdr"]     = 71
R["fcf_col_hdr"] = 72
R["nopat2"]      = 73
R["da"]          = 74
R["da_pct"]      = 75
R["capex"]       = 76
R["capex_pct"]   = 77
R["nwc"]         = 78
R["nwc_pct"]     = 79
R["blank9"]      = 80
R["ufcf"]        = 81
R["fcf_m"]       = 82
R["blank10"]     = 83

# Discounting
R["disc_hdr"]    = 84
R["disc_col_hdr"]= 85
R["period"]      = 86
R["disc_factor"] = 87
R["pv_fcf"]      = 88
R["blank11"]     = 89

# Terminal value
R["tv_hdr"]      = 90
R["tv_fcf"]      = 91
R["tv_val"]      = 92
R["pv_tv"]       = 93
R["blank12"]     = 94

# Valuation summary
R["val_hdr"]     = 95
R["sum_pv_fcf"]  = 96
R["sum_pv_tv"]   = 97
R["ev"]          = 98
R["less_debt"]   = 99
R["equity"]      = 100
R["blank13"]     = 101
R["shares_out"]  = 102
R["price_impl"]  = 103
R["price_curr"]  = 104
R["upside"]      = 105
R["blank14"]     = 106

# Sensitivity tables start at row 108
R["sens_start"]  = 108


def build_dcf_sheet(wb):
    ws = wb.create_sheet("DCF", 0)   # Insert as first sheet

    # ── Column widths ──────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 32
    for c in range(2, 12):           # B..K  (years)
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.column_dimensions["L"].width = 12   # Selected
    ws.column_dimensions["M"].width = 4    # spacer
    ws.column_dimensions["N"].width = 12   # Bear assumptions
    ws.column_dimensions["O"].width = 12   # Base assumptions
    ws.column_dimensions["P"].width = 12   # Bull assumptions

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 1: Header
    # ══════════════════════════════════════════════════════════════════════
    ws.merge_cells("A1:K1")
    c = ws["A1"]
    c.value = "Spruce Power Holding Corporation (NYSE: SPRU) – DCF Valuation Model"
    c.font = fnt(bold=True, size=14, color=C_WHITE)
    c.fill = fill(C_HDR_DARK)
    c.alignment = align("center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:K2")
    c = ws["A2"]
    c.value = "7-Year Projection  |  2026E – 2032E  |  Bear / Base / Bull  |  As of 2026-08-17  |  $ in millions"
    c.font = fnt(italic=True, size=9)
    c.alignment = align("center")

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 2: Case Selector
    # ══════════════════════════════════════════════════════════════════════
    ws.cell(row=R["selector_lbl"], column=1,
            value="► Case Selector (1=Bear  2=Base  3=Bull)").font = fnt(bold=True)
    sel = input_cell(ws, R["selector"], 2, 2, fmt="0",
                     comment_text="Source: User input. Enter 1 for Bear, 2 for Base, 3 for Bull.")
    sel.font = fnt(bold=True, size=12, color=C_BLUE_FONT)
    sel.fill = fill("FFF2CC")   # yellow highlight for selector
    sel.alignment = align("center")

    ws.cell(row=R["case_name"], column=1, value="Selected Scenario:").font = fnt(bold=True)
    c = ws.cell(row=R["case_name"], column=2,
                value='=IF(B5=1,"Bear Case",IF(B5=2,"Base Case","Bull Case"))')
    c.font = fnt(bold=True, color=C_BLACK_FONT)
    c.fill = fill(C_OUTPUT_FILL)
    c.alignment = align("center")
    ws.merge_cells(start_row=R["case_name"], start_column=2,
                   end_row=R["case_name"], end_column=4)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 3: Market Data
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["mkt_hdr"], 1, 5, "MARKET DATA & KEY INPUTS")
    for col_idx, txt in enumerate(["Item", "Value", "Unit", "Notes"], start=1):
        col_header(ws, R["mkt_col_hdr"], col_idx, txt, center=(col_idx > 1))

    def mkt_row(row, lbl, val, unit, note, fmt=None, comment=None):
        label_cell(ws, row, 1, lbl)
        input_cell(ws, row, 2, val, fmt=fmt or FMT_USD, comment_text=comment)
        ws.cell(row=row, column=3, value=unit).alignment = align("center")
        ws.cell(row=row, column=4, value=note).font = fnt(italic=True, size=9)

    mkt_row(R["price"],    "Current Stock Price",        STOCK_PRICE, "$/share",
            "NYSE close 2026-08-14", fmt=FMT_PRICE,
            comment="Source: NYSE close, 2026-08-14. StockAnalysis.com")
    mkt_row(R["shares"],   "Shares Outstanding (Basic)", SHARES_BASIC, "M shares",
            "Q2 2026 10-Q (Jun 30, 2026)",
            comment="Source: Q2 2026 10-Q, 2026-08-12. Shares as of June 30, 2026: 19,249,671")
    formula_cell(ws, R["mktcap"], 2, f"=B{R['price']}*B{R['shares']}", fmt=FMT_USD)
    label_cell(ws, R["mktcap"], 1, "Market Cap ($M)")
    ws.cell(row=R["mktcap"], column=3, value="$M")
    mkt_row(R["net_debt"], "Net Debt ($M)",               NET_DEBT, "$M",
            "Debt $679.5M − Cash $81.5M  (Q2 2026)",
            comment="Source: Q2 2026 10-Q, 2026-08-12. Total debt $679.5M less total cash $81.5M.")

    thick_box(ws, R["mkt_hdr"], R["net_debt"], 1, 5)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 4: Scenario Assumption Blocks
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["assum_hdr"], 1, 11,
                   "SCENARIO ASSUMPTIONS  (Bear | Base | Bull — Projection Years)")

    ASSUM_ROWS = [
        ("bear", R["bear_hdr"],  R["bear_yr_hdr"],  R["bear_grow"],  "Bear"),
        ("base", R["base_hdr"],  R["base_yr_hdr"],  R["base_grow"],  "Base"),
        ("bull", R["bull_hdr"],  R["bull_yr_hdr"],  R["bull_grow"],  "Bull"),
    ]

    SCEN_ROWS = {
        "Bear": {"hdr": R["bear_hdr"], "yr": R["bear_yr_hdr"],
                 "grow": R["bear_grow"], "ebit": R["bear_ebit"],
                 "da": R["bear_da"], "capex": R["bear_capex"],
                 "nwc": R["bear_nwc"], "tax": R["bear_tax"],
                 "tg": R["bear_tg"], "wacc": R["bear_wacc"]},
        "Base": {"hdr": R["base_hdr"], "yr": R["base_yr_hdr"],
                 "grow": R["base_grow"], "ebit": R["base_ebit"],
                 "da": R["base_da"], "capex": R["base_capex"],
                 "nwc": R["base_nwc"], "tax": R["base_tax"],
                 "tg": R["base_tg"], "wacc": R["base_wacc"]},
        "Bull": {"hdr": R["bull_hdr"], "yr": R["bull_yr_hdr"],
                 "grow": R["bull_grow"], "ebit": R["bull_ebit"],
                 "da": R["bull_da"], "capex": R["bull_capex"],
                 "nwc": R["bull_nwc"], "tax": R["bull_tax"],
                 "tg": R["bull_tg"], "wacc": R["bull_wacc"]},
    }

    SCEN_LABELS = {
        "grow":  "Revenue Growth (%)",
        "ebit":  "EBIT Margin (%)",
        "da":    "D&A % of Revenue",
        "capex": "CapEx % of Revenue",
        "nwc":   "NWC Change % of ΔRev",
        "tax":   "Tax Rate (%)",
        "tg":    "Terminal Growth Rate",
        "wacc":  "WACC",
    }

    SCEN_VALS = {
        "Bear": {
            "grow":  REV_GROWTH["Bear"],
            "ebit":  EBIT_MARGIN["Bear"],
            "da":    [DA_PCT["Bear"]] * 7,
            "capex": [CAPEX_PCT["Bear"]] * 7,
            "nwc":   [NWC_PCT["Bear"]] * 7,
            "tax":   [TAX_RATE["Bear"]] * 7,
            "tg":    [TERM_G["Bear"]] + [None]*6,
            "wacc":  [WACC_S["Bear"]] + [None]*6,
        },
        "Base": {
            "grow":  REV_GROWTH["Base"],
            "ebit":  EBIT_MARGIN["Base"],
            "da":    [DA_PCT["Base"]] * 7,
            "capex": [CAPEX_PCT["Base"]] * 7,
            "nwc":   [NWC_PCT["Base"]] * 7,
            "tax":   [TAX_RATE["Base"]] * 7,
            "tg":    [TERM_G["Base"]] + [None]*6,
            "wacc":  [WACC_S["Base"]] + [None]*6,
        },
        "Bull": {
            "grow":  REV_GROWTH["Bull"],
            "ebit":  EBIT_MARGIN["Bull"],
            "da":    [DA_PCT["Bull"]] * 7,
            "capex": [CAPEX_PCT["Bull"]] * 7,
            "nwc":   [NWC_PCT["Bull"]] * 7,
            "tax":   [TAX_RATE["Bull"]] * 7,
            "tg":    [TERM_G["Bull"]] + [None]*6,
            "wacc":  [WACC_S["Bull"]] + [None]*6,
        },
    }

    ASSUM_COMMENTS = {
        "grow":  "Source: Model assumptions 2026-08-17. Revenue CAGR anchored to portfolio size, Spruce Pro ramp, and M&A activity.",
        "ebit":  "Source: Model assumptions 2026-08-17. Based on Q2 2026 EBIT margin 32% (single-quarter) and structural SG&A reduction.",
        "da":    "Source: Model assumptions 2026-08-17. D&A driven by solar asset depreciation (useful life 30yr).",
        "capex": "Source: Model assumptions 2026-08-17. CapEx minimal – no new-build, maintenance only.",
        "nwc":   "Source: Model assumptions 2026-08-17. NWC change on incremental revenue; contracted revenue = low WC drag.",
        "tax":   "Source: Q2 2026 10-Q, 2026-08-12. Large NOL balance shields all taxable income in forecast period.",
        "tg":    "Source: Model assumptions 2026-08-17. Terminal growth reflects long-term contracted + renewal cash flows.",
        "wacc":  "Source: CAPM calculation 2026-08-17. Rf=4.68%, Beta=1.20, ERP=5.50%, CSRP per scenario.",
    }

    for scen in ["Bear", "Base", "Bull"]:
        sr = SCEN_ROWS[scen]
        # Section header
        section_header(ws, sr["hdr"], 1, 11, f"{scen.upper()} CASE ASSUMPTIONS")
        # Year header row
        label_cell(ws, sr["yr"], 1, "Assumption")
        for i, y in enumerate(PROJ_YEARS):
            col_header(ws, sr["yr"], 2+len(HIST_YEARS)+i, f"{y}E")

        # Data rows
        for key in ["grow","ebit","da","capex","nwc","tax","tg","wacc"]:
            row_r = sr[key]
            label_cell(ws, row_r, 1, "  " + SCEN_LABELS[key])
            vals = SCEN_VALS[scen][key]
            for i, v in enumerate(vals):
                c_idx = 2 + len(HIST_YEARS) + i   # starts at E column
                if v is not None:
                    input_cell(ws, row_r, c_idx, v, fmt=FMT_PCT,
                               comment_text=ASSUM_COMMENTS[key])

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 5: Consolidation (Selected) block
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["sel_hdr"], 1, 11,
                   "SELECTED CASE — Consolidation Column (drives all projection formulas)")

    label_cell(ws, R["sel_yr_hdr"], 1, "Assumption")
    for i, y in enumerate(PROJ_YEARS):
        col_header(ws, R["sel_yr_hdr"], 2+len(HIST_YEARS)+i, f"{y}E")

    SEL_KEYS = ["grow","ebit","da","capex","nwc","tax","tg","wacc"]

    for key in SEL_KEYS:
        row_r = R[f"sel_{key}"]
        label_cell(ws, row_r, 1, "  " + SCEN_LABELS[key])
        for i in range(7):
            # Column index for this projection year
            c_idx = 2 + len(HIST_YEARS) + i   # E..K
            bear_addr = ws.cell(row=SCEN_ROWS["Bear"][key], column=c_idx).coordinate
            base_addr = ws.cell(row=SCEN_ROWS["Base"][key], column=c_idx).coordinate
            bull_addr = ws.cell(row=SCEN_ROWS["Bull"][key], column=c_idx).coordinate

            # Only terminal g and wacc have values in first column only
            bear_val = SCEN_VALS["Bear"][key][i]
            base_val = SCEN_VALS["Base"][key][i]
            bull_val = SCEN_VALS["Bull"][key][i]

            if bear_val is None:
                # Use first year value for scalar assumptions
                bear_addr = ws.cell(row=SCEN_ROWS["Bear"][key], column=2+len(HIST_YEARS)).coordinate
                base_addr = ws.cell(row=SCEN_ROWS["Base"][key], column=2+len(HIST_YEARS)).coordinate
                bull_addr = ws.cell(row=SCEN_ROWS["Bull"][key], column=2+len(HIST_YEARS)).coordinate

            formula = (f"=IF($B$5=1,{bear_addr},"
                       f"IF($B$5=2,{base_addr},{bull_addr}))")
            c2 = formula_cell(ws, row_r, c_idx, formula, fmt=FMT_PCT)
            c2.fill = fill(C_HDR_LIGHT)

    thick_box(ws, R["sel_hdr"], R["sel_wacc"], 1, 11)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 6: Revenue & EBIT Projection
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["is_hdr"], 1, 11,
                   "INCOME STATEMENT SUMMARY ($M)")

    # Column headers
    label_cell(ws, R["is_col_hdr"], 1, "Item")
    for y in ALL_YEARS:
        col_header(ws, R["is_col_hdr"], YEAR_COLS[y],
                   f"{y}A" if y <= 2025 else f"{y}E")

    # Revenue row
    label_cell(ws, R["rev"], 1, "Revenue ($M)", bold=True)
    # Historical actuals
    for y, v in [(2023,79.9),(2024,82.1),(2025,111.8)]:
        c2 = input_cell(ws, R["rev"], YEAR_COLS[y], v, fmt=FMT_USD,
                        comment_text=f"Source: SPRU {y} 10-K / earnings release.")
    # Projections: first year off 2025A, rest chain
    c_2026 = get_column_letter(YEAR_COLS[2026])
    d_2025 = ws.cell(row=R["rev"], column=YEAR_COLS[2025]).coordinate
    # sel_grow row for each proj year
    for i, y in enumerate(PROJ_YEARS):
        prev_col = get_column_letter(YEAR_COLS[PROJ_YEARS[i-1]] if i > 0 else YEAR_COLS[2025])
        grow_cell = ws.cell(row=R["sel_grow"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["rev"], YEAR_COLS[y],
                     f"={prev_col}{R['rev']}*(1+{grow_cell})",
                     fmt=FMT_USD)

    # Revenue growth %
    label_cell(ws, R["rev_grow"], 1, "  % Growth YoY")
    for y in HIST_YEARS[1:]:  # 2024, 2025
        prev = YEAR_COLS[y-1]
        cur  = YEAR_COLS[y]
        formula_cell(ws, R["rev_grow"], cur,
                     f"={get_column_letter(cur)}{R['rev']}/{get_column_letter(prev)}{R['rev']}-1",
                     fmt=FMT_PCT)
    for y in PROJ_YEARS:
        grow_cell = ws.cell(row=R["sel_grow"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["rev_grow"], YEAR_COLS[y],
                     f"={grow_cell}", fmt=FMT_PCT)

    # EBIT
    label_cell(ws, R["ebit"], 1, "EBIT ($M)", bold=True)
    for y, v in [(2023,-13.4),(2024,-16.8),(2025,17.9)]:
        input_cell(ws, R["ebit"], YEAR_COLS[y], v, fmt=FMT_USD,
                   comment_text=f"Source: SPRU {y} 10-K / earnings release.")
    for y in PROJ_YEARS:
        rev_c = f"{get_column_letter(YEAR_COLS[y])}{R['rev']}"
        ebit_m_c = ws.cell(row=R["sel_ebit"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["ebit"], YEAR_COLS[y],
                     f"={rev_c}*{ebit_m_c}", fmt=FMT_USD)

    # EBIT margin
    label_cell(ws, R["ebit_m"], 1, "  EBIT Margin %")
    for y in HIST_YEARS:
        ev = ws.cell(row=R["ebit"], column=YEAR_COLS[y]).value
        rv = ws.cell(row=R["rev"],  column=YEAR_COLS[y]).value
        if rv and rv != 0:
            input_cell(ws, R["ebit_m"], YEAR_COLS[y],
                       (ev or 0) / rv, fmt=FMT_PCT,
                       comment_text=f"Source: Derived from {y} actuals.")
    for y in PROJ_YEARS:
        ebit_m_c = ws.cell(row=R["sel_ebit"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["ebit_m"], YEAR_COLS[y], f"={ebit_m_c}", fmt=FMT_PCT)

    # Tax (0%)
    label_cell(ws, R["tax"], 1, "  Taxes ($M)")
    for y in PROJ_YEARS:
        ebit_c = f"{get_column_letter(YEAR_COLS[y])}{R['ebit']}"
        tax_r_c = ws.cell(row=R["sel_tax"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["tax"], YEAR_COLS[y],
                     f"={ebit_c}*{tax_r_c}", fmt=FMT_USD)

    label_cell(ws, R["tax_r"], 1, "  Tax Rate %")
    for y in PROJ_YEARS:
        tax_r_c = ws.cell(row=R["sel_tax"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["tax_r"], YEAR_COLS[y], f"={tax_r_c}", fmt=FMT_PCT)

    # NOPAT
    label_cell(ws, R["nopat"], 1, "NOPAT ($M)", bold=True)
    for y in HIST_YEARS:
        ebit_v = ws.cell(row=R["ebit"], column=YEAR_COLS[y]).value
        input_cell(ws, R["nopat"], YEAR_COLS[y], ebit_v or 0, fmt=FMT_USD,
                   comment_text="Source: Derived (NOPAT = EBIT × (1−T), T=0).")
    for y in PROJ_YEARS:
        ebit_c = f"{get_column_letter(YEAR_COLS[y])}{R['ebit']}"
        tax_c  = f"{get_column_letter(YEAR_COLS[y])}{R['tax']}"
        formula_cell(ws, R["nopat"], YEAR_COLS[y],
                     f"={ebit_c}-{tax_c}", fmt=FMT_USD)

    thick_box(ws, R["is_hdr"], R["nopat"], 1, 11)

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 7: FCF Build
    # ══════════════════════════════════════════════════════════════════════
    section_header(ws, R["fcf_hdr"], 1, 11, "FREE CASH FLOW BUILD ($M)")
    label_cell(ws, R["fcf_col_hdr"], 1, "Item")
    for y in ALL_YEARS:
        col_header(ws, R["fcf_col_hdr"], YEAR_COLS[y],
                   f"{y}A" if y <= 2025 else f"{y}E")

    # NOPAT line (mirror)
    label_cell(ws, R["nopat2"], 1, "NOPAT ($M)", bold=True)
    for y in ALL_YEARS:
        nopat_src = f"{get_column_letter(YEAR_COLS[y])}{R['nopat']}"
        formula_cell(ws, R["nopat2"], YEAR_COLS[y], f"={nopat_src}", fmt=FMT_USD)

    # D&A
    label_cell(ws, R["da"], 1, "(+) D&A ($M)")
    for y, v in [(2023,21.6),(2024,20.3),(2025,26.1)]:
        input_cell(ws, R["da"], YEAR_COLS[y], v, fmt=FMT_USD,
                   comment_text=f"Source: SPRU {y} 10-K. D&A including solar system depreciation.")
    for y in PROJ_YEARS:
        rev_c  = f"{get_column_letter(YEAR_COLS[y])}{R['rev']}"
        da_pct_c = ws.cell(row=R["sel_da"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["da"], YEAR_COLS[y], f"={rev_c}*{da_pct_c}", fmt=FMT_USD)

    label_cell(ws, R["da_pct"], 1, "    % of Revenue")
    for y in HIST_YEARS:
        da_v  = ws.cell(row=R["da"],  column=YEAR_COLS[y]).value
        rev_v = ws.cell(row=R["rev"], column=YEAR_COLS[y]).value
        if rev_v:
            formula_cell(ws, R["da_pct"], YEAR_COLS[y],
                         f"={get_column_letter(YEAR_COLS[y])}{R['da']}/{get_column_letter(YEAR_COLS[y])}{R['rev']}",
                         fmt=FMT_PCT)
    for y in PROJ_YEARS:
        da_pct_c = ws.cell(row=R["sel_da"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["da_pct"], YEAR_COLS[y], f"={da_pct_c}", fmt=FMT_PCT)

    # CapEx
    label_cell(ws, R["capex"], 1, "(-) CapEx ($M)")
    for y, v in [(2023,0.3),(2024,0.4),(2025,0.2)]:
        input_cell(ws, R["capex"], YEAR_COLS[y], v, fmt=FMT_USD,
                   comment_text=f"Source: SPRU {y} cash flow statement. Purchases of property/equipment.")
    for y in PROJ_YEARS:
        rev_c    = f"{get_column_letter(YEAR_COLS[y])}{R['rev']}"
        capex_c  = ws.cell(row=R["sel_capex"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["capex"], YEAR_COLS[y], f"={rev_c}*{capex_c}", fmt=FMT_USD)

    label_cell(ws, R["capex_pct"], 1, "    % of Revenue")
    for y in PROJ_YEARS:
        capex_c = ws.cell(row=R["sel_capex"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["capex_pct"], YEAR_COLS[y], f"={capex_c}", fmt=FMT_PCT)

    # NWC change
    label_cell(ws, R["nwc"], 1, "(-) Δ NWC ($M)")
    for y in HIST_YEARS:
        input_cell(ws, R["nwc"], YEAR_COLS[y], 0.0, fmt=FMT_USD,
                   comment_text="Source: Approximated. Contracted revenue model has minimal NWC drag.")
    for i, y in enumerate(PROJ_YEARS):
        prev_y = PROJ_YEARS[i-1] if i > 0 else 2025
        rev_c  = f"{get_column_letter(YEAR_COLS[y])}{R['rev']}"
        prev_c = f"{get_column_letter(YEAR_COLS[prev_y])}{R['rev']}"
        nwc_c  = ws.cell(row=R["sel_nwc"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["nwc"], YEAR_COLS[y],
                     f"=({rev_c}-{prev_c})*{nwc_c}", fmt=FMT_USD)

    label_cell(ws, R["nwc_pct"], 1, "    % of ΔRevenue")
    for y in PROJ_YEARS:
        nwc_c = ws.cell(row=R["sel_nwc"], column=YEAR_COLS[y]).coordinate
        formula_cell(ws, R["nwc_pct"], YEAR_COLS[y], f"={nwc_c}", fmt=FMT_PCT)

    # Unlevered FCF
    label_cell(ws, R["ufcf"], 1, "Unlevered Free Cash Flow ($M)", bold=True)
    for y in HIST_YEARS:
        nopat_c = f"{get_column_letter(YEAR_COLS[y])}{R['nopat2']}"
        da_c    = f"{get_column_letter(YEAR_COLS[y])}{R['da']}"
        capex_c = f"{get_column_letter(YEAR_COLS[y])}{R['capex']}"
        nwc_c   = f"{get_column_letter(YEAR_COLS[y])}{R['nwc']}"
        c2 = formula_cell(ws, R["ufcf"], YEAR_COLS[y],
                          f"={nopat_c}+{da_c}-{capex_c}-{nwc_c}", fmt=FMT_USD)
        c2.font = fnt(bold=True)
    for y in PROJ_YEARS:
        nopat_c = f"{get_column_letter(YEAR_COLS[y])}{R['nopat2']}"
        da_c    = f"{get_column_letter(YEAR_COLS[y])}{R['da']}"
        capex_c = f"{get_column_letter(YEAR_COLS[y])}{R['capex']}"
        nwc_c   = f"{get_column_letter(YEAR_COLS[y])}{R['nwc']}"
        c2 = formula_cell(ws, R["ufcf"], YEAR_COLS[y],
                          f"={nopat_c}+{da_c}-{capex_c}-{nwc_c}", fmt=FMT_USD)
        c2.font = fnt(bold=True)
        c2.fill = fill(C_OUTPUT_FILL)

    label_cell(ws, R["fcf_m"], 1, "  FCF Margin %")
    for y in ALL_YEARS:
        ufcf_c = f"{get_column_letter(YEAR_COLS[y])}{R['ufcf']}"
        rev_c  = f"{get_column_letter(YEAR_COLS[y])}{R['rev']}"
        formula_cell(ws, R["fcf_m"], YEAR_COLS[y],
                     f"={ufcf_c}/{rev_c}", fmt=FMT_PCT)

    thick_box(ws, R["fcf_hdr"], R["fcf_m"], 1, 11)

    return ws, R


if __name__ == "__main__":
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    build_wacc_sheet(wb)
    ws_dcf, R_out = build_dcf_sheet(wb)

    wb.save(OUTPUT_PATH)
    print(f"Part 2 saved: {OUTPUT_PATH}")
    print("DCF sheet: assumptions + revenue + FCF built.")
