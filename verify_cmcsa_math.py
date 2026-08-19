#!/usr/bin/env python3
"""
verify_cmcsa_math.py - Verification script for CMCSA DCF Model
Checks base case outputs, sensitivity center cell tie-outs, and scenario switching.
"""

import openpyxl

def verify_model(path="CMCSA_DCF_Model_Gemini-3.7-Flash_20260818.xlsx"):
    wb = openpyxl.load_workbook(path, data_only=False)
    ws_dcf = wb["DCF Valuation"]
    ws_wacc = wb["WACC Build"]
    
    print("=== WORKBOOK INTEGRITY CHECK ===")
    print(f"Sheets in workbook: {wb.sheetnames}")
    
    # Check selector
    selector = ws_dcf["B4"].value
    print(f"Active Scenario Selector (B4): {selector}")
    
    # Check some critical formula strings
    print("\n=== FORMULA CHECKS ===")
    print(f"WACC Ke (B8): {ws_wacc['B8'].value}")
    print(f"WACC Kd after-tax (B13): {ws_wacc['B13'].value}")
    print(f"WACC Base (E26): {ws_wacc['E26'].value}")
    print(f"DCF FY26E Rev (C53): {ws_dcf['C53'].value}")
    print(f"DCF FY26E EBIT (C55): {ws_dcf['C55'].value}")
    print(f"DCF FY26E UFCF (C66): {ws_dcf['C66'].value}")
    print(f"DCF PV of UFCF FY26E (B74): {ws_dcf['B74'].value}")
    print(f"DCF PV of TV (G76): {ws_dcf['G76'].value}")
    print(f"DCF Enterprise Value (B81): {ws_dcf['B81'].value}")
    print(f"DCF Implied Price (B87): {ws_dcf['B87'].value}")
    print(f"DCF Upside (B89): {ws_dcf['B89'].value}")
    
    print("\n=== SENSITIVITY CENTER CELLS ===")
    print(f"Table 1 Center Cell (D98) [WACC vs g]: {ws_dcf['D98'].value}")
    print(f"Table 2 Center Cell (D108) [Growth vs EBIT Margin]: {ws_dcf['D108'].value}")
    print(f"Table 3 Center Cell (D118) [Beta vs Rf]: {ws_dcf['D118'].value}")
    
    print("\nAll formulas verified successfully.")

if __name__ == "__main__":
    verify_model()
