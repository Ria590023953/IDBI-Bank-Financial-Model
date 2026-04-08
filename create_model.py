"""
IDBI Bank Comprehensive Financial Model Generator
Creates a professional Excel workbook (.xlsx) with 7 sheets:
  1. Inputs
  2. 3-Statements Model
  3. DCF Valuation
  4. Trading Comps
  5. Scenario Analysis
  6. Sensitivity Analysis
  7. Dashboard
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import os

# ── Colours / Styles ──────────────────────────────────────────────────────────
BLUE_FONT    = Font(name="Calibri", size=10, bold=False, color="0000CC")
BLACK_FONT   = Font(name="Calibri", size=10, bold=False, color="000000")
BOLD_FONT    = Font(name="Calibri", size=10, bold=True,  color="000000")
TITLE_FONT   = Font(name="Calibri", size=12, bold=True,  color="FFFFFF")
HDR_FONT     = Font(name="Calibri", size=10, bold=True,  color="FFFFFF")
SECTION_FONT = Font(name="Calibri", size=10, bold=True,  color="000000")

INPUT_FILL   = PatternFill("solid", fgColor="DCE6F1")
CALC_FILL    = PatternFill("solid", fgColor="FFFFE0")
RESULT_FILL  = PatternFill("solid", fgColor="E2EFDA")
HDR_FILL     = PatternFill("solid", fgColor="1F4E79")
SECTION_FILL = PatternFill("solid", fgColor="BDD7EE")

THIN = Side(style="thin", color="BFBFBF")
MED  = Side(style="medium", color="1F4E79")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MED_BORDER  = Border(left=MED,  right=MED,  top=MED,  bottom=MED)

FMT_INR  = '#,##0'
FMT_INR2 = '#,##0.00'
FMT_PCT  = '0.00%'
FMT_RS   = '₹#,##0.00'
FMT_X    = '0.00"x"'


# ── Helpers ───────────────────────────────────────────────────────────────────
def cw(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width


def cell_style(c, font=None, fill=None, border=None, align=None, fmt=None):
    if font:   c.font   = font
    if fill:   c.fill   = fill
    if border: c.border = border
    if align:  c.alignment = align
    if fmt:    c.number_format = fmt


def hdr(ws, row, col, value, fill=None, span=1):
    c = ws.cell(row=row, column=col, value=value)
    cell_style(c, HDR_FONT, fill or HDR_FILL, THIN_BORDER,
               Alignment(horizontal="center", vertical="center", wrap_text=True))
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return c


def title(ws, row, col, value, span=1):
    c = ws.cell(row=row, column=col, value=value)
    cell_style(c, TITLE_FONT, HDR_FILL, MED_BORDER,
               Alignment(horizontal="center", vertical="center"))
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    ws.row_dimensions[row].height = 24
    return c


def section(ws, row, col, value, span=1):
    c = ws.cell(row=row, column=col, value=value)
    cell_style(c, SECTION_FONT, SECTION_FILL, MED_BORDER,
               Alignment(horizontal="left", vertical="center"))
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return c


def inp(ws, row, col, value, fmt=None):
    """Blue-font hard-coded input cell."""
    c = ws.cell(row=row, column=col, value=value)
    cell_style(c, BLUE_FONT, INPUT_FILL, THIN_BORDER,
               Alignment(horizontal="right", vertical="center"), fmt)
    return c


def lbl(ws, row, col, value, bold=False, indent=0):
    c = ws.cell(row=row, column=col, value=value)
    cell_style(c, BOLD_FONT if bold else BLACK_FONT, border=THIN_BORDER,
               align=Alignment(horizontal="left", vertical="center", indent=indent))
    return c


def calc(ws, row, col, formula, fmt=None, fill=None):
    """Black-font formula / calculated cell."""
    c = ws.cell(row=row, column=col, value=formula)
    cell_style(c, BLACK_FONT, fill or CALC_FILL, THIN_BORDER,
               Alignment(horizontal="right", vertical="center"), fmt)
    return c


def res(ws, row, col, formula, fmt=None):
    """Result / output cell (green background)."""
    c = ws.cell(row=row, column=col, value=formula)
    cell_style(c, BLACK_FONT, RESULT_FILL, THIN_BORDER,
               Alignment(horizontal="right", vertical="center"), fmt)
    return c


def scen_if(sc_row, base, bull, bear):
    """Return an Excel IF formula selecting base/bull/bear value."""
    return (
        f'=IF(Inputs!B{sc_row}="Bull",{bull},'
        f'IF(Inputs!B{sc_row}="Bear",{bear},{base}))'
    )


# ── SHEET 1: INPUTS ───────────────────────────────────────────────────────────
def build_inputs(wb):
    ws = wb.create_sheet("Inputs")
    ws.freeze_panes = "C2"

    for col, w in [("A", 35), ("B", 18), ("C", 18), ("D", 18), ("E", 18)]:
        cw(ws, col, w)

    # Row 1 – title
    title(ws, 1, 1, "IDBI BANK – FINANCIAL MODEL INPUTS", span=5)

    r = 3
    # Legend
    section(ws, r, 1, "LEGEND", span=3); r += 1
    c1 = ws.cell(row=r, column=1, value="Blue Font  =  Hardcoded Input")
    c1.font = BLUE_FONT; c1.border = THIN_BORDER
    c2 = ws.cell(row=r, column=2, value="Black Font  =  Formula / Calculation")
    c2.font = BLACK_FONT; c2.border = THIN_BORDER
    r += 2

    # ── Current Valuation  (rows 6-13)
    section(ws, r, 1, "CURRENT VALUATION (April 2026)", span=3); r += 1  # r=6
    valuation_pairs = [
        ("Stock Price (CMP) – ₹",           69.55,     FMT_RS),       # B7
        ("Market Cap – ₹ Crore",            74782.96,  FMT_INR2),     # B8
        ("Shares Outstanding – Billion",    10.75,     FMT_INR2),     # B9
        ("52-Week High – ₹",               118.38,    FMT_RS),       # B10
        ("52-Week Low – ₹",                 61.01,     FMT_RS),       # B11
        ("P/E Ratio",                        8.05,     "0.00"),       # B12
        ("P/B Ratio",                        1.12,     "0.00"),       # B13
        ("EPS (TTM) – ₹",                   8.64,     FMT_RS),       # B14
        ("Book Value per Share – ₹",        65.00,     FMT_RS),       # B15
    ]
    VALUATION_ROWS = {}
    for lbl_txt, val, fmt in valuation_pairs:
        lbl(ws, r, 1, lbl_txt)
        inp(ws, r, 2, val, fmt)
        VALUATION_ROWS[lbl_txt] = r
        r += 1
    r += 1  # blank

    # ── FY 2024-25 Income Statement Actuals  (header at r=17, data 18-27)
    section(ws, r, 1, "FY 2024-25 INCOME STATEMENT ACTUALS – ₹ Millions", span=4); r += 1  # r=17
    hdr(ws, r, 1, "Line Item"); hdr(ws, r, 2, "FY 2024-25"); hdr(ws, r, 3, "FY 2023-24"); r += 1  # r=18

    IS_ROWS = {}
    is_data = [
        ("Interest Income",               289171, 264571),   # B19
        ("Other Income",                   60000,  55000),   # B20
        ("Total Income",                  349171, 319571),   # B21
        ("Interest Expenses",             142566, 132059),   # B22
        ("Net Interest Income",           146605, 142187),   # B23
        ("Operating Expenses",             86190,  83679),   # B24
        ("Provisions & Contingencies",     31500,  33000),   # B25
        ("Pre-Tax Profit",                 89915,  68508),   # B26
        ("Tax",                            13608,  10000),   # B27
        ("Net Profit After Tax",           76307,  57850),   # B28
    ]
    for lbl_txt, v25, v24 in is_data:
        lbl(ws, r, 1, lbl_txt)
        inp(ws, r, 2, v25, FMT_INR)
        inp(ws, r, 3, v24, FMT_INR)
        IS_ROWS[lbl_txt] = r
        r += 1
    r += 1  # blank

    # ── FY 2024-25 Balance Sheet Actuals
    section(ws, r, 1, "FY 2024-25 BALANCE SHEET – ₹ Crore", span=4); r += 1
    hdr(ws, r, 1, "Line Item"); hdr(ws, r, 2, "FY 2024-25"); hdr(ws, r, 3, "FY 2023-24"); r += 1

    BS_ROWS = {}
    bs_data = [
        ("Total Assets",         412962, 364755),
        ("Cash & Equivalents",    35000,  28000),
        ("Investments",          120000, 110000),
        ("Advances",             218212, 188440),
        ("Other Assets",          39750,  38315),
        ("Total Deposits",       309975, 277295),
        ("Borrowings",            40000,  38000),
        ("Other Liabilities",     25800,  23000),
        ("Total Equity",          37187,  26460),
        ("Paid-up Capital",       10750,  10750),
        ("Reserves & Surplus",    26437,  15710),
    ]
    for lbl_txt, v25, v24 in bs_data:
        lbl(ws, r, 1, lbl_txt)
        inp(ws, r, 2, v25, FMT_INR)
        inp(ws, r, 3, v24, FMT_INR)
        BS_ROWS[lbl_txt] = r
        r += 1
    r += 1  # blank

    # ── Key Ratios
    section(ws, r, 1, "KEY RATIOS (FY 2024-25)", span=3); r += 1
    RATIO_ROWS = {}
    RATIO_ROWS = {}
    ratio_data = [
        ("Return on Equity (ROE) %",       0.2015, FMT_PCT),
        ("Return on Assets (ROA) %",       0.0198, FMT_PCT),
        ("Net Interest Margin (NIM) %",    0.044,  FMT_PCT),
        ("Cost to Income Ratio %",         0.45,   FMT_PCT),
        ("Gross NPA %",                    0.0298, FMT_PCT),
        ("Net NPA %",                      0.0015, FMT_PCT),
        ("Capital Adequacy Ratio %",       0.2505, FMT_PCT),
        ("Net Profit Margin %",            0.264,  FMT_PCT),
    ]
    for lbl_txt, val, fmt in ratio_data:
        lbl(ws, r, 1, lbl_txt)
        inp(ws, r, 2, val, fmt)
        RATIO_ROWS[lbl_txt] = r
        r += 1
    r += 1  # blank

    # ── Assumption Drivers
    section(ws, r, 1, "ASSUMPTION DRIVERS", span=5); r += 1
    hdr(ws, r, 1, "Driver"); hdr(ws, r, 2, "Base Case"); hdr(ws, r, 3, "Bull Case"); hdr(ws, r, 4, "Bear Case"); r += 1

    DRIVER_ROWS = {}
    drivers = [
        ("Interest Income Growth %",    0.085, 0.11,  0.05),
        ("NII Growth %",                0.080, 0.11,  0.05),
        ("Other Income Growth %",       0.070, 0.09,  0.04),
        ("Cost to Income Ratio %",      0.450, 0.40,  0.50),
        ("Cost of Deposits %",          0.045, 0.042, 0.050),
        ("Advance Growth %",            0.140, 0.17,  0.10),
        ("Deposit Growth %",            0.120, 0.15,  0.08),
        ("Provision Coverage %",        0.015, 0.010, 0.025),
        ("Tax Rate %",                  0.300, 0.28,  0.30),
        ("Dividend Payout Ratio %",     0.200, 0.15,  0.25),
        ("Depreciation % of Assets",    0.005, 0.005, 0.005),
        ("CapEx % of Assets",           0.008, 0.010, 0.006),
    ]
    for lbl_txt, base, bull, bear in drivers:
        lbl(ws, r, 1, lbl_txt)
        inp(ws, r, 2, base, FMT_PCT)
        inp(ws, r, 3, bull, FMT_PCT)
        inp(ws, r, 4, bear, FMT_PCT)
        DRIVER_ROWS[lbl_txt] = r
        r += 1
    r += 1  # blank

    # ── Scenario Selector
    section(ws, r, 1, "SCENARIO SELECTOR", span=3); r += 1
    lbl(ws, r, 1, "Active Scenario  (Base / Bull / Bear) →", bold=True)
    sc = ws.cell(row=r, column=2, value="Base")
    sc.font      = Font(name="Calibri", size=11, bold=True, color="1F4E79")
    sc.fill      = PatternFill("solid", fgColor="FFF2CC")
    sc.border    = MED_BORDER
    sc.alignment = Alignment(horizontal="center", vertical="center")
    SC_ROW = r
    dv = DataValidation(type="list", formula1='"Base,Bull,Bear"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(sc)
    r += 2  # blank

    # ── WACC Components
    section(ws, r, 1, "WACC COMPONENTS", span=3); r += 1
    WACC_ROWS = {}
    wacc_inputs = [
        ("Risk-Free Rate (Rf) %",            0.065,  FMT_PCT, False),
        ("Equity Risk Premium (ERP) %",      0.070,  FMT_PCT, False),
        ("Beta (β)",                         0.85,   "0.00",  False),
        ("Cost of Equity % (CAPM = Rf+β×ERP)", None, FMT_PCT, True),
        ("Cost of Debt (pre-tax) %",         0.075,  FMT_PCT, False),
        ("Tax Rate (for WACC) %",            0.30,   FMT_PCT, False),
        ("After-Tax Cost of Debt %",         None,   FMT_PCT, True),
        ("Target D/E Ratio",                 6.0,    "0.00",  False),
        ("Weight of Equity  [1/(1+D/E)]",    None,   FMT_PCT, True),
        ("Weight of Debt  [D/E/(1+D/E)]",    None,   FMT_PCT, True),
        ("WACC %",                           None,   FMT_PCT, True),
    ]
    for lbl_txt, val, fmt, is_formula in wacc_inputs:
        lbl(ws, r, 1, lbl_txt)
        if not is_formula:
            inp(ws, r, 2, val, fmt)
        else:
            c = ws.cell(row=r, column=2, value=0)
            cell_style(c, BLACK_FONT, CALC_FILL, THIN_BORDER,
                       Alignment(horizontal="right"), fmt)
        WACC_ROWS[lbl_txt] = r
        r += 1

    # Fill WACC formulas
    rf_r   = WACC_ROWS["Risk-Free Rate (Rf) %"]
    erp_r  = WACC_ROWS["Equity Risk Premium (ERP) %"]
    bet_r  = WACC_ROWS["Beta (β)"]
    ke_r   = WACC_ROWS["Cost of Equity % (CAPM = Rf+β×ERP)"]
    kd_r   = WACC_ROWS["Cost of Debt (pre-tax) %"]
    ktax_r = WACC_ROWS["Tax Rate (for WACC) %"]
    kdat_r = WACC_ROWS["After-Tax Cost of Debt %"]
    de_r   = WACC_ROWS["Target D/E Ratio"]
    we_r   = WACC_ROWS["Weight of Equity  [1/(1+D/E)]"]
    wd_r   = WACC_ROWS["Weight of Debt  [D/E/(1+D/E)]"]
    wacc_r = WACC_ROWS["WACC %"]

    ws.cell(row=ke_r,   column=2).value = f"=B{rf_r}+B{bet_r}*B{erp_r}"
    ws.cell(row=kdat_r, column=2).value = f"=B{kd_r}*(1-B{ktax_r})"
    ws.cell(row=we_r,   column=2).value = f"=1/(1+B{de_r})"
    ws.cell(row=wd_r,   column=2).value = f"=B{de_r}/(1+B{de_r})"
    ws.cell(row=wacc_r, column=2).value = f"=B{we_r}*B{ke_r}+B{wd_r}*B{kdat_r}"

    r += 1  # blank

    # ── Terminal Value
    section(ws, r, 1, "TERMINAL VALUE ASSUMPTION", span=3); r += 1
    lbl(ws, r, 1, "Perpetual Growth Rate (g) %")
    inp(ws, r, 2, 0.03, FMT_PCT)
    TGR_ROW = r
    r += 2  # blank

    # ── Peer Bank Data
    section(ws, r, 1, "PEER BANK DATA (April 2026)", span=8); r += 1
    peer_hdrs = ["Bank", "Market Cap\n(₹ Cr)", "Stock Price\n(₹)",
                 "Shares (Cr)", "Net Profit\n(₹ Cr)", "NII / Revenue\n(₹ Cr)",
                 "Book Value\n/ Share (₹)", "P/E", "P/B"]
    for ci, h in enumerate(peer_hdrs, 1):
        hdr(ws, r, ci, h)
    r += 1
    PEER_START = r
    peers = [
        ("HDFC Bank",      1350000, 1805.0, 748,  66200, 200000, 593.0, 20.4, 3.04),
        ("ICICI Bank",      895000, 1265.0, 708,  46500, 150000, 404.0, 19.3, 3.13),
        ("Axis Bank",       346000, 1120.0, 309,  26000,  95000, 548.0, 13.3, 2.04),
        ("SBI",             720000,  808.0, 892,  67100, 240000, 475.0, 10.7, 1.70),
        ("Kotak Mahindra",  391000, 1966.0, 199,  15300,  72000, 710.0, 25.5, 2.77),
    ]
    for bk, mc, sp, sh, np_, nii, bv, pe, pb in peers:
        ws.cell(row=r, column=1, value=bk).font = BLUE_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        for ci, v in enumerate([mc, sp, sh, np_, nii, bv, pe, pb], 2):
            inp(ws, r, ci, v, FMT_INR2)
        r += 1

    return (ws, DRIVER_ROWS, SC_ROW, wacc_r, TGR_ROW, IS_ROWS, BS_ROWS,
            rf_r, erp_r, bet_r, kd_r, de_r, PEER_START, RATIO_ROWS)


# ── SHEET 2: 3-STATEMENTS MODEL ───────────────────────────────────────────────
def build_3statements(wb, DR, SC_ROW, IS_ROWS, BS_ROWS):
    ws = wb.create_sheet("3-Statements Model")
    ws.freeze_panes = "B3"

    for col, w in [("A", 36), ("B", 15), ("C", 15), ("D", 15), ("E", 15), ("F", 15)]:
        cw(ws, col, w)

    years = ["FY 2024-25", "FY 2025-26", "FY 2026-27", "FY 2027-28", "FY 2028-29"]
    YC = [2, 3, 4, 5, 6]  # column indices for each year

    title(ws, 1, 1, "IDBI BANK – 3-STATEMENTS MODEL", span=6)
    r = 2
    hdr(ws, r, 1, "Line Item")
    for i, y in enumerate(years):
        hdr(ws, r, YC[i], y)
    r += 1

    # Pre-extract all driver row numbers (avoids complex bracket nesting in f-strings)
    r_ii  = DR["Interest Income Growth %"]
    r_oi  = DR["Other Income Growth %"]
    r_cir = DR["Cost to Income Ratio %"]
    r_cod = DR["Cost of Deposits %"]
    r_adv = DR["Advance Growth %"]
    r_dep = DR["Deposit Growth %"]
    r_pro = DR["Provision Coverage %"]
    r_tax = DR["Tax Rate %"]
    r_div = DR["Dividend Payout Ratio %"]
    r_dep_pct = DR["Depreciation % of Assets"]
    r_cap = DR["CapEx % of Assets"]

    # Inputs row numbers for actuals
    II_ROW  = IS_ROWS["Interest Income"]
    OI_ROW  = IS_ROWS["Other Income"]
    IE_ROW  = IS_ROWS["Interest Expenses"]
    ADV_INP = BS_ROWS["Advances"]
    INV_INP = BS_ROWS["Investments"]
    CSH_INP = BS_ROWS["Cash & Equivalents"]
    DEP_INP = BS_ROWS["Total Deposits"]
    BOR_INP = BS_ROWS["Borrowings"]
    PUC_INP = BS_ROWS["Paid-up Capital"]
    RES_INP = BS_ROWS["Reserves & Surplus"]

    def si(base_ref, bull_ref, bear_ref):
        """Scenario IF selector."""
        return (f'=IF(Inputs!B{SC_ROW}="Bull",{bull_ref},'
                f'IF(Inputs!B{SC_ROW}="Bear",{bear_ref},{base_ref}))')

    # ── INCOME STATEMENT ──────────────────────────────────────────────────────
    section(ws, r, 1, "INCOME STATEMENT  (₹ Millions)", span=6); r += 1

    # Interest Income
    lbl(ws, r, 1, "Interest Income")
    calc(ws, r, 2, f"=Inputs!B{II_ROW}", FMT_INR)
    II_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+Inputs!B{r_ii})",
                f"{pcl}{r}*(1+Inputs!C{r_ii})",
                f"{pcl}{r}*(1+Inputs!D{r_ii})"),
             FMT_INR)
    r += 1

    # Other Income
    lbl(ws, r, 1, "Other Income")
    calc(ws, r, 2, f"=Inputs!B{OI_ROW}", FMT_INR)
    OI_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+Inputs!B{r_oi})",
                f"{pcl}{r}*(1+Inputs!C{r_oi})",
                f"{pcl}{r}*(1+Inputs!D{r_oi})"),
             FMT_INR)
    r += 1

    # Total Income
    lbl(ws, r, 1, "Total Income", bold=True)
    TI_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{II_R}+{cl}{OI_R}", FMT_INR)
    r += 1

    # Interest Expenses
    lbl(ws, r, 1, "Interest Expenses")
    calc(ws, r, 2, f"=Inputs!B{IE_ROW}", FMT_INR)
    IE_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+Inputs!B{r_cod})",
                f"{pcl}{r}*(1+Inputs!C{r_cod})",
                f"{pcl}{r}*(1+Inputs!D{r_cod})"),
             FMT_INR)
    r += 1

    # Net Interest Income
    lbl(ws, r, 1, "Net Interest Income", bold=True)
    NII_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{TI_R}-{cl}{IE_R}", FMT_INR)
    r += 1

    # Operating Expenses
    lbl(ws, r, 1, "Operating Expenses")
    OPEX_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             si(f"{cl}{TI_R}*Inputs!B{r_cir}",
                f"{cl}{TI_R}*Inputs!C{r_cir}",
                f"{cl}{TI_R}*Inputs!D{r_cir}"),
             FMT_INR)
    r += 1

    # Provisions
    lbl(ws, r, 1, "Provisions & Contingencies")
    PROV_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             si(f"{cl}{NII_R}*Inputs!B{r_pro}",
                f"{cl}{NII_R}*Inputs!C{r_pro}",
                f"{cl}{NII_R}*Inputs!D{r_pro}"),
             FMT_INR)
    r += 1

    # Pre-Tax Profit
    lbl(ws, r, 1, "Pre-Tax Profit", bold=True)
    PTP_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{NII_R}-{cl}{OPEX_R}-{cl}{PROV_R}", FMT_INR)
    r += 1

    # Tax
    lbl(ws, r, 1, "Tax (30%)")
    TAX_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             si(f"{cl}{PTP_R}*Inputs!B{r_tax}",
                f"{cl}{PTP_R}*Inputs!C{r_tax}",
                f"{cl}{PTP_R}*Inputs!D{r_tax}"),
             FMT_INR)
    r += 1

    # Net Profit After Tax
    lbl(ws, r, 1, "Net Profit After Tax (NPAT)", bold=True)
    NPAT_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{PTP_R}-{cl}{TAX_R}", FMT_INR)
    r += 2

    # ── Key P&L Ratios ────────────────────────────────────────────────────────
    section(ws, r, 1, "KEY RATIOS", span=6); r += 1
    hdr(ws, r, 1, "Ratio")
    for i, y in enumerate(years):
        hdr(ws, r, YC[i], y)
    r += 1

    lbl(ws, r, 1, "Net Profit Margin %")
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"=IF({cl}{TI_R}=0,0,{cl}{NPAT_R}/{cl}{TI_R})", FMT_PCT)
    r += 1

    lbl(ws, r, 1, "Cost to Income Ratio %")
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"=IF({cl}{TI_R}=0,0,{cl}{OPEX_R}/{cl}{TI_R})", FMT_PCT)
    r += 2

    # ── BALANCE SHEET ─────────────────────────────────────────────────────────
    section(ws, r, 1, "BALANCE SHEET  (₹ Crore)", span=6); r += 1
    hdr(ws, r, 1, "Line Item")
    for i, y in enumerate(years):
        hdr(ws, r, YC[i], y)
    r += 1

    section(ws, r, 1, "ASSETS", span=6); r += 1

    # Advances
    lbl(ws, r, 1, "Advances")
    calc(ws, r, 2, f"=Inputs!B{ADV_INP}", FMT_INR)
    ADV_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+Inputs!B{r_adv})",
                f"{pcl}{r}*(1+Inputs!C{r_adv})",
                f"{pcl}{r}*(1+Inputs!D{r_adv})"),
             FMT_INR)
    r += 1

    # Investments (grow at 65% of advance growth)
    lbl(ws, r, 1, "Investments")
    calc(ws, r, 2, f"=Inputs!B{INV_INP}", FMT_INR)
    INV_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+0.65*Inputs!B{r_adv})",
                f"{pcl}{r}*(1+0.65*Inputs!C{r_adv})",
                f"{pcl}{r}*(1+0.65*Inputs!D{r_adv})"),
             FMT_INR)
    r += 1

    # Cash
    lbl(ws, r, 1, "Cash & Equivalents")
    calc(ws, r, 2, f"=Inputs!B{CSH_INP}", FMT_INR)
    CSH_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci, f"={pcl}{r}*1.05", FMT_INR)
    r += 1

    # Other Assets
    lbl(ws, r, 1, "Other Assets")
    OTH_A_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={cl}{ADV_R}*0.18", FMT_INR)
    r += 1

    # Total Assets
    lbl(ws, r, 1, "TOTAL ASSETS", bold=True)
    TA_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{ADV_R}+{cl}{INV_R}+{cl}{CSH_R}+{cl}{OTH_A_R}", FMT_INR)
    r += 1

    section(ws, r, 1, "LIABILITIES", span=6); r += 1

    # Deposits
    lbl(ws, r, 1, "Total Deposits")
    calc(ws, r, 2, f"=Inputs!B{DEP_INP}", FMT_INR)
    DEP_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci,
             si(f"{pcl}{r}*(1+Inputs!B{r_dep})",
                f"{pcl}{r}*(1+Inputs!C{r_dep})",
                f"{pcl}{r}*(1+Inputs!D{r_dep})"),
             FMT_INR)
    r += 1

    # Borrowings
    lbl(ws, r, 1, "Borrowings")
    calc(ws, r, 2, f"=Inputs!B{BOR_INP}", FMT_INR)
    BOR_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        calc(ws, r, ci, f"={pcl}{r}*1.05", FMT_INR)
    r += 1

    # Other Liabilities
    lbl(ws, r, 1, "Other Liabilities")
    OTH_L_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={cl}{TA_R}*0.06", FMT_INR)
    r += 1

    # Total Liabilities
    lbl(ws, r, 1, "Total Liabilities (excl. Equity)", bold=True)
    TL_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{DEP_R}+{cl}{BOR_R}+{cl}{OTH_L_R}", FMT_INR)
    r += 1

    section(ws, r, 1, "EQUITY", span=6); r += 1

    # Paid-up Capital
    lbl(ws, r, 1, "Paid-up Capital")
    PUC_R = r
    for ci in YC:
        calc(ws, r, ci, f"=Inputs!B{PUC_INP}", FMT_INR)
    r += 1

    # Reserves & Surplus
    lbl(ws, r, 1, "Reserves & Surplus")
    calc(ws, r, 2, f"=Inputs!B{RES_INP}", FMT_INR)
    RES_R = r
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        cl  = get_column_letter(ci)
        retained = si(
            f"({cl}{NPAT_R}/10)*(1-Inputs!B{r_div})",
            f"({cl}{NPAT_R}/10)*(1-Inputs!C{r_div})",
            f"({cl}{NPAT_R}/10)*(1-Inputs!D{r_div})"
        )
        # Replace leading "=" since we're embedding in a sum formula
        retained_expr = retained.lstrip("=")
        calc(ws, r, ci, f"={pcl}{RES_R}+{retained_expr}", FMT_INR)
    r += 1

    # Total Equity
    lbl(ws, r, 1, "Total Equity", bold=True)
    EQ_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{PUC_R}+{cl}{RES_R}", FMT_INR)
    r += 1

    # Balance check
    lbl(ws, r, 1, "Total Liabilities + Equity  (balance check)", bold=True)
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{TL_R}+{cl}{EQ_R}", FMT_INR)
    r += 2

    # ── CASH FLOW STATEMENT ───────────────────────────────────────────────────
    section(ws, r, 1, "CASH FLOW STATEMENT  (₹ Millions)", span=6); r += 1
    hdr(ws, r, 1, "Line Item")
    for i, y in enumerate(years):
        hdr(ws, r, YC[i], y)
    r += 1

    # Operating – Net Profit
    lbl(ws, r, 1, "Net Profit After Tax")
    CFS1_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={'3-Statements Model'}!{cl}{NPAT_R}" if False else
             f"={cl}{NPAT_R}", FMT_INR)
    r += 1

    # Add: Provisions
    lbl(ws, r, 1, "Add: Provisions & Contingencies")
    CFS2_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={cl}{PROV_R}", FMT_INR)
    r += 1

    # Add: Depreciation
    lbl(ws, r, 1, "Add: Depreciation (% of total assets)")
    CFS3_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={cl}{TA_R}*10*Inputs!B{r_dep_pct}", FMT_INR)
    r += 1

    # Change in Advances
    lbl(ws, r, 1, "Less: Increase in Advances (use of funds)")
    CFS4_R = r
    calc(ws, r, 2, "=0", FMT_INR)
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        cl  = get_column_letter(ci)
        calc(ws, r, ci, f"=-({cl}{ADV_R}-{pcl}{ADV_R})*10", FMT_INR)
    r += 1

    # Change in Deposits
    lbl(ws, r, 1, "Add: Increase in Deposits (source of funds)")
    CFS5_R = r
    calc(ws, r, 2, "=0", FMT_INR)
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        cl  = get_column_letter(ci)
        calc(ws, r, ci, f"=({cl}{DEP_R}-{pcl}{DEP_R})*10", FMT_INR)
    r += 1

    lbl(ws, r, 1, "OPERATING CASH FLOW", bold=True)
    OCF_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci,
            f"={cl}{CFS1_R}+{cl}{CFS2_R}+{cl}{CFS3_R}+{cl}{CFS4_R}+{cl}{CFS5_R}",
            FMT_INR)
    r += 1

    # Investing
    lbl(ws, r, 1, "Capital Expenditure")
    CAPEX_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"=-{cl}{TA_R}*10*Inputs!B{r_cap}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "INVESTING CASH FLOW", bold=True)
    ICF_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{CAPEX_R}", FMT_INR)
    r += 1

    # Financing
    lbl(ws, r, 1, "Dividends Paid")
    DIV_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             si(f"-{cl}{NPAT_R}*Inputs!B{r_div}",
                f"-{cl}{NPAT_R}*Inputs!C{r_div}",
                f"-{cl}{NPAT_R}*Inputs!D{r_div}"),
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "FINANCING CASH FLOW", bold=True)
    FCF_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{DIV_R}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "NET CHANGE IN CASH", bold=True)
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{OCF_R}+{cl}{ICF_R}+{cl}{FCF_R}", FMT_INR)
    r += 1

    ROWS = dict(
        NPAT_R=NPAT_R, NII_R=NII_R, TI_R=TI_R, II_R=II_R, OI_R=OI_R,
        IE_R=IE_R, PROV_R=PROV_R, OPEX_R=OPEX_R, PTP_R=PTP_R, TAX_R=TAX_R,
        TA_R=TA_R, DEP_R=DEP_R, ADV_R=ADV_R, EQ_R=EQ_R, CAPEX_R=CAPEX_R,
        OCF_R=OCF_R, ICF_R=ICF_R, FCF_R=FCF_R, CFS3_R=CFS3_R, BOR_R=BOR_R,
    )
    return ws, ROWS


# ── SHEET 3: DCF VALUATION ────────────────────────────────────────────────────
def build_dcf(wb, SR, DR, WACC_R, TGR_ROW, SC_ROW, BS_ROWS):
    ws = wb.create_sheet("DCF Valuation")
    ws.freeze_panes = "B3"

    for col, w in [("A", 42), ("B", 16), ("C", 16), ("D", 16), ("E", 16), ("F", 16)]:
        cw(ws, col, w)

    years = ["FY 2024-25", "FY 2025-26", "FY 2026-27", "FY 2027-28", "FY 2028-29"]
    YC = [2, 3, 4, 5, 6]

    title(ws, 1, 1, "IDBI BANK – DCF VALUATION", span=6)
    r = 2
    hdr(ws, r, 1, "Component")
    for i, y in enumerate(years):
        hdr(ws, r, YC[i], y)
    r += 1

    section(ws, r, 1, "FREE CASH FLOW TO FIRM (FCFF) – ₹ Millions", span=6); r += 1

    lbl(ws, r, 1, "Net Profit After Tax")
    NPAT_D = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"='3-Statements Model'!{cl}{SR['NPAT_R']}", FMT_INR)
    r += 1

    r_tax = DR["Tax Rate %"]
    lbl(ws, r, 1, "Add: Interest Expense × (1 – Tax Rate)")
    IADJ_D = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             f"='3-Statements Model'!{cl}{SR['IE_R']}*(1-Inputs!B{r_tax})",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "Less: Change in Working Capital (net advances – deposits)")
    WC_D = r
    calc(ws, r, 2, "=0", FMT_INR)
    for ci in YC[1:]:
        pcl = get_column_letter(ci - 1)
        cl  = get_column_letter(ci)
        calc(ws, r, ci,
             f"=(-('3-Statements Model'!{cl}{SR['ADV_R']}-'3-Statements Model'!{pcl}{SR['ADV_R']})"
             f"+'3-Statements Model'!{cl}{SR['DEP_R']}-'3-Statements Model'!{pcl}{SR['DEP_R']})*10",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "Less: Capital Expenditure")
    CAPEX_D = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"='3-Statements Model'!{cl}{SR['CAPEX_R']}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "FCFF (Net Profit + Interest×(1-t) + ΔWC + CapEx)", bold=True)
    FCFF_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        res(ws, r, ci,
            f"={cl}{NPAT_D}+{cl}{IADJ_D}+{cl}{WC_D}+{cl}{CAPEX_D}",
            FMT_INR)
    r += 2

    section(ws, r, 1, "WACC & DISCOUNT RATE", span=6); r += 1

    lbl(ws, r, 1, "WACC %  (from Inputs)")
    WACC_DISP = r
    calc(ws, r, 2, f"=Inputs!B{WACC_R}", FMT_PCT)
    r += 1

    lbl(ws, r, 1, "Terminal Growth Rate (g) %  (from Inputs)")
    TGR_DISP = r
    calc(ws, r, 2, f"=Inputs!B{TGR_ROW}", FMT_PCT)
    r += 1

    lbl(ws, r, 1, "Terminal Value  (Year-5 FCFF × (1+g) / (WACC-g))")
    TV_R = r
    res(ws, r, 2,
        f"=F{FCFF_R}*(1+B{TGR_DISP})/(B{WACC_DISP}-B{TGR_DISP})",
        FMT_INR)
    r += 2

    section(ws, r, 1, "PRESENT VALUE OF CASH FLOWS", span=6); r += 1
    hdr(ws, r, 1, "Period")
    for n in range(1, 6):
        hdr(ws, r, n + 1, f"Year {n}")
    r += 1

    lbl(ws, r, 1, "FCFF  (₹ Mn)")
    FCFF2_R = r
    for ci in YC:
        cl = get_column_letter(ci)
        calc(ws, r, ci, f"={cl}{FCFF_R}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "Discount Factor  [1/(1+WACC)^n]")
    DISC_R = r
    for i, ci in enumerate(YC):
        calc(ws, r, ci, f"=1/(1+B{WACC_DISP})^{i+1}", "0.000000")
    r += 1

    lbl(ws, r, 1, "PV of FCFF  (₹ Mn)")
    PV_R = r
    for i, ci in enumerate(YC):
        cl = get_column_letter(ci)
        res(ws, r, ci, f"={cl}{DISC_R}*{cl}{FCFF2_R}", FMT_INR)
    r += 2

    section(ws, r, 1, "ENTERPRISE VALUE → EQUITY VALUE → IMPLIED PRICE", span=6); r += 1

    lbl(ws, r, 1, "Sum of PV of FCFFs  (₹ Mn)")
    SUM_PV_R = r
    calc(ws, r, 2, f"=SUM(B{PV_R}:F{PV_R})", FMT_INR)
    r += 1

    lbl(ws, r, 1, "PV of Terminal Value  (₹ Mn)")
    PV_TV_R = r
    calc(ws, r, 2, f"=B{TV_R}/(1+B{WACC_DISP})^5", FMT_INR)
    r += 1

    lbl(ws, r, 1, "Enterprise Value  (₹ Mn)", bold=True)
    EV_R = r
    res(ws, r, 2, f"=B{SUM_PV_R}+B{PV_TV_R}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "Less: Net Debt  (Borrowings – Cash, ₹ Mn)")
    ND_R = r
    calc(ws, r, 2,
         f"=(Inputs!B{BS_ROWS['Borrowings']}-Inputs!B{BS_ROWS['Cash & Equivalents']})*10",
         FMT_INR)
    r += 1

    lbl(ws, r, 1, "Equity Value  (₹ Mn)", bold=True)
    EQV_R = r
    res(ws, r, 2, f"=B{EV_R}-B{ND_R}", FMT_INR)
    r += 1

    lbl(ws, r, 1, "Shares Outstanding  (Millions)")
    SH_R = r
    calc(ws, r, 2, "=Inputs!B9*1000", FMT_INR)
    r += 1

    lbl(ws, r, 1, "Implied Share Price  (₹)", bold=True)
    ISP_R = r
    res(ws, r, 2, f"=B{EQV_R}/B{SH_R}", FMT_RS)
    r += 1

    lbl(ws, r, 1, "Current Market Price  (₹)")
    CMP_R = r
    calc(ws, r, 2, "=Inputs!B7", FMT_RS)
    r += 1

    lbl(ws, r, 1, "Upside / Downside %", bold=True)
    UDS_R = r
    res(ws, r, 2, f"=(B{ISP_R}-B{CMP_R})/B{CMP_R}", FMT_PCT)
    r += 1

    return ws, dict(
        ISP_R=ISP_R, WACC_DISP=WACC_DISP, TGR_DISP=TGR_DISP,
        EV_R=EV_R, EQV_R=EQV_R, SH_R=SH_R, CMP_R=CMP_R,
        FCFF_R=FCFF_R, PV_R=PV_R, TV_R=TV_R,
    )


# ── SHEET 4: TRADING COMPS ────────────────────────────────────────────────────
def build_comps(wb, PEER_START, IS_ROWS, BS_ROWS):
    ws = wb.create_sheet("Trading Comps")
    ws.freeze_panes = "B3"

    for col, w in [("A", 20), ("B", 14), ("C", 12), ("D", 12),
                   ("E", 14), ("F", 14), ("G", 14),
                   ("H", 11), ("I", 11), ("J", 11), ("K", 11), ("L", 14)]:
        cw(ws, col, w)

    title(ws, 1, 1, "IDBI BANK – TRADING COMPARABLES ANALYSIS", span=12)
    r = 2
    peer_hdrs = [
        "Bank", "Market Cap\n(₹ Cr)", "Stock Price\n(₹)", "Shares\n(Cr)",
        "Net Profit\n(₹ Cr)", "NII / Rev.\n(₹ Cr)", "BV / Share\n(₹)",
        "EV/EBITDA", "EV/Rev.", "P/E", "P/B", "EV  (₹ Cr)"
    ]
    for ci, h in enumerate(peer_hdrs, 1):
        hdr(ws, r, ci, h)
    r += 1

    peer_data = [
        ("HDFC Bank",      "Inputs!B" + str(PEER_START),   "Inputs!C" + str(PEER_START),
                           "Inputs!D" + str(PEER_START),   "Inputs!E" + str(PEER_START),
                           "Inputs!F" + str(PEER_START),   "Inputs!G" + str(PEER_START)),
        ("ICICI Bank",     "Inputs!B" + str(PEER_START+1), "Inputs!C" + str(PEER_START+1),
                           "Inputs!D" + str(PEER_START+1), "Inputs!E" + str(PEER_START+1),
                           "Inputs!F" + str(PEER_START+1), "Inputs!G" + str(PEER_START+1)),
        ("Axis Bank",      "Inputs!B" + str(PEER_START+2), "Inputs!C" + str(PEER_START+2),
                           "Inputs!D" + str(PEER_START+2), "Inputs!E" + str(PEER_START+2),
                           "Inputs!F" + str(PEER_START+2), "Inputs!G" + str(PEER_START+2)),
        ("SBI",            "Inputs!B" + str(PEER_START+3), "Inputs!C" + str(PEER_START+3),
                           "Inputs!D" + str(PEER_START+3), "Inputs!E" + str(PEER_START+3),
                           "Inputs!F" + str(PEER_START+3), "Inputs!G" + str(PEER_START+3)),
        ("Kotak Mahindra", "Inputs!B" + str(PEER_START+4), "Inputs!C" + str(PEER_START+4),
                           "Inputs!D" + str(PEER_START+4), "Inputs!E" + str(PEER_START+4),
                           "Inputs!F" + str(PEER_START+4), "Inputs!G" + str(PEER_START+4)),
    ]

    PEER_ROWS_START = r
    for bank_name, mc, sp, sh, np_, nii, bv in peer_data:
        lbl(ws, r, 1, bank_name)
        ws.cell(row=r, column=1).font = BLUE_FONT
        for ci, ref in enumerate([mc, sp, sh, np_, nii, bv], 2):
            calc(ws, r, ci, f"={ref}", FMT_INR)

        # EV = Market Cap + Net Debt proxy (15% of MC for banking approximation)
        calc(ws, r, 12, f"=B{r}*1.15", FMT_INR)  # EV in col L

        # EV/EBITDA: EV / (NetProfit * 1.3)  ← proxy for EBITDA
        calc(ws, r, 8, f"=IF(E{r}=0,0,L{r}/(E{r}*1.3))", FMT_X)

        # EV/Revenue
        calc(ws, r, 9, f"=IF(F{r}=0,0,L{r}/F{r})", FMT_X)

        # P/E  (= Stock Price × Shares / Net Profit; all in ₹ Crore units)
        calc(ws, r, 10, f"=IF(E{r}=0,0,C{r}*D{r}/E{r})", FMT_X)

        # P/B  (price / BV per share)
        calc(ws, r, 11, f"=IF(G{r}=0,0,C{r}/G{r})", FMT_X)

        r += 1

    PEER_ROWS_END = r - 1
    r += 1

    # Summary statistics
    for stat_name, fn in [("Median", "MEDIAN"), ("Min", "MIN"), ("Max", "MAX")]:
        lbl(ws, r, 1, stat_name, bold=(stat_name == "Median"))
        for ci in [8, 9, 10, 11]:
            cl = get_column_letter(ci)
            f = f"={fn}({cl}{PEER_ROWS_START}:{cl}{PEER_ROWS_END})"
            if stat_name == "Median":
                res(ws, r, ci, f, FMT_X)
            else:
                calc(ws, r, ci, f, FMT_X)
        if stat_name == "Median":
            MED_ROW = r
        r += 1

    r += 1

    # ── IDBI Implied Valuation Range ──────────────────────────────────────────
    section(ws, r, 1, "IDBI BANK – IMPLIED VALUATION (based on Peer Medians)", span=8); r += 1
    hdr(ws, r, 1, "Metric"); hdr(ws, r, 2, "Median Multiple")
    hdr(ws, r, 3, "IDBI Metric  (₹ Cr)"); hdr(ws, r, 4, "Implied MC  (₹ Cr)")
    hdr(ws, r, 5, "Implied Price  (₹)"); hdr(ws, r, 6, "CMP  (₹)"); hdr(ws, r, 7, "Upside %")
    r += 1

    npat_row = IS_ROWS["Net Profit After Tax"]
    nii_row  = IS_ROWS["Net Interest Income"]
    sh_r = "Inputs!B9"   # Shares in Billions

    idbi_comps = [
        ("EV/EBITDA",
         f"=H{MED_ROW}",
         f"=Inputs!B{npat_row}/10*1.3",    # NPAT Mn→Cr, ×1.3 for EBITDA proxy
         "ev"),
        ("EV/Revenue",
         f"=I{MED_ROW}",
         f"=Inputs!B{nii_row}/10",          # NII Mn→Cr
         "ev"),
        ("P/E Ratio",
         f"=J{MED_ROW}",
         "=Inputs!B14",                     # EPS (₹)
         "pe"),
        ("P/B Ratio",
         f"=K{MED_ROW}",
         "=Inputs!B15",                     # Book Value per Share
         "pb"),
    ]

    IMPLIED_ROWS = []
    for metric, mult_f, metric_f, method in idbi_comps:
        lbl(ws, r, 1, metric)
        calc(ws, r, 2, mult_f, FMT_X)
        calc(ws, r, 3, metric_f, FMT_INR2)

        if method == "ev":
            # Implied EV = median × EBITDA/Revenue; MC = EV/1.15
            calc(ws, r, 4, f"=B{r}*C{r}/1.15", FMT_INR)
            # Implied Price (₹) = MC (₹ Cr) × 10 / Shares (Bn)
            res(ws, r, 5, f"=D{r}*10/{sh_r}", FMT_RS)
        else:
            # Implied MC (₹ Cr) = multiple × metric (₹/share) × Shares_Bn × 100
            calc(ws, r, 4, f"=B{r}*C{r}*{sh_r}*100", FMT_INR)
            # Implied Price = multiple × metric (e.g. P/E × EPS, or P/B × BV)
            res(ws, r, 5, f"=B{r}*C{r}", FMT_RS)

        calc(ws, r, 6, "=Inputs!B7", FMT_RS)
        res(ws, r, 7, f"=IF(F{r}=0,0,(E{r}-F{r})/F{r})", FMT_PCT)
        IMPLIED_ROWS.append(r)
        r += 1

    return ws, {"COMPS_ISP_ROWS": IMPLIED_ROWS, "MED_ROW": MED_ROW}


# ── SHEET 5: SCENARIO ANALYSIS ────────────────────────────────────────────────
def build_scenarios(wb, SR, DR, SC_ROW, IS_ROWS, DCF_ROWS):
    ws = wb.create_sheet("Scenario Analysis")
    ws.freeze_panes = "B3"

    for col, w in [("A", 38), ("B", 16), ("C", 16), ("D", 16)]:
        cw(ws, col, w)

    title(ws, 1, 1, "IDBI BANK – SCENARIO ANALYSIS", span=4)
    r = 2

    section(ws, r, 1, "SCENARIO SELECTOR  (driven by Inputs sheet)", span=4); r += 1
    lbl(ws, r, 1, "Active Scenario:")
    sc_disp = ws.cell(row=r, column=2, value=f"=Inputs!B{SC_ROW}")
    sc_disp.font      = Font(name="Calibri", size=11, bold=True, color="1F4E79")
    sc_disp.fill      = PatternFill("solid", fgColor="FFF2CC")
    sc_disp.border    = MED_BORDER
    sc_disp.alignment = Alignment(horizontal="center")
    sc_disp.number_format = "@"
    r += 2

    # ── Key Assumptions Table ─────────────────────────────────────────────────
    section(ws, r, 1, "KEY ASSUMPTIONS BY SCENARIO", span=4); r += 1
    hdr(ws, r, 1, "Assumption"); hdr(ws, r, 2, "Base Case")
    hdr(ws, r, 3, "Bull Case"); hdr(ws, r, 4, "Bear Case")
    r += 1

    asm_keys = [
        "Interest Income Growth %", "NII Growth %",
        "Cost to Income Ratio %", "Advance Growth %",
        "Deposit Growth %", "Provision Coverage %", "Tax Rate %",
    ]
    for k in asm_keys:
        dr = DR[k]
        lbl(ws, r, 1, k)
        calc(ws, r, 2, f"=Inputs!B{dr}", FMT_PCT)
        calc(ws, r, 3, f"=Inputs!C{dr}", FMT_PCT)
        calc(ws, r, 4, f"=Inputs!D{dr}", FMT_PCT)
        r += 1
    r += 1

    # ── P&L Results – computed for FY 2028-29 under each scenario ─────────────
    section(ws, r, 1, "P&L RESULTS – FY 2028-29 PROJECTIONS  (₹ Millions)", span=4); r += 1
    hdr(ws, r, 1, "Metric"); hdr(ws, r, 2, "Base Case")
    hdr(ws, r, 3, "Bull Case"); hdr(ws, r, 4, "Bear Case")
    r += 1

    ii_row  = IS_ROWS["Interest Income"]
    nii_row = IS_ROWS["Net Interest Income"]
    opex_row= IS_ROWS["Operating Expenses"]
    npat_row= IS_ROWS["Net Profit After Tax"]

    r_ii  = DR["Interest Income Growth %"]
    r_nii = DR["NII Growth %"]
    r_cir = DR["Cost to Income Ratio %"]
    r_pro = DR["Provision Coverage %"]
    r_tax = DR["Tax Rate %"]

    sc_cols = [("B", "B"), ("B", "C"), ("B", "D")]  # (inp_col, driver_col)

    lbl(ws, r, 1, "Interest Income  (₹ Mn)")
    for ci, (_, dc) in enumerate(sc_cols, 2):
        calc(ws, r, ci,
             f"=Inputs!B{ii_row}*(1+Inputs!{dc}{r_ii})^4",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "Net Interest Income  (₹ Mn)")
    NII_SC_ROW = r
    for ci, (_, dc) in enumerate(sc_cols, 2):
        calc(ws, r, ci,
             f"=Inputs!B{nii_row}*(1+Inputs!{dc}{r_nii})^4",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "Operating Expenses  (₹ Mn)")
    for ci, (_, dc) in enumerate(sc_cols, 2):
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             f"={cl}{NII_SC_ROW}/Inputs!{dc}{r_nii}*Inputs!{dc}{r_cir}"
             if False else
             f"=Inputs!B{ii_row}*(1+Inputs!{dc}{r_ii})^4*Inputs!{dc}{r_cir}",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "Provisions  (₹ Mn)")
    for ci, (_, dc) in enumerate(sc_cols, 2):
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             f"={cl}{NII_SC_ROW}*Inputs!{dc}{r_pro}",
             FMT_INR)
    PROV_SC_ROW = r
    r += 1

    lbl(ws, r, 1, "Net Profit After Tax  (₹ Mn)", bold=True)
    NPAT_SC = r
    for ci, (_, dc) in enumerate(sc_cols, 2):
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             f"=({cl}{NII_SC_ROW}-Inputs!B{ii_row}*(1+Inputs!{dc}{r_ii})^4"
             f"*Inputs!{dc}{r_cir}-{cl}{PROV_SC_ROW})*(1-Inputs!{dc}{r_tax})",
             FMT_INR, RESULT_FILL)
    r += 2

    # ── Valuation Summary ──────────────────────────────────────────────────────
    section(ws, r, 1, "VALUATION SUMMARY BY SCENARIO", span=4); r += 1
    hdr(ws, r, 1, "Metric"); hdr(ws, r, 2, "Base Case")
    hdr(ws, r, 3, "Bull Case"); hdr(ws, r, 4, "Bear Case")
    r += 1

    r_div = DR["Dividend Payout Ratio %"]
    wacc_ref = f"'DCF Valuation'!B{DCF_ROWS['WACC_DISP']}"
    tgr_ref  = f"'DCF Valuation'!B{DCF_ROWS['TGR_DISP']}"
    sh_ref   = f"'DCF Valuation'!B{DCF_ROWS['SH_R']}"

    lbl(ws, r, 1, "Terminal FCFF  (NPAT FY28-29 × (1+g), ₹ Mn)")
    TV_SC_R = r
    for ci, (_, dc) in enumerate(sc_cols, 2):
        cl = get_column_letter(ci)
        calc(ws, r, ci,
             f"={cl}{NPAT_SC}*(1+Inputs!{dc}{r_ii})",
             FMT_INR)
    r += 1

    lbl(ws, r, 1, "WACC  (%)")
    for ci in [2, 3, 4]:
        calc(ws, r, ci, f"={wacc_ref}", FMT_PCT)
    r += 1

    lbl(ws, r, 1, "Implied Share Price  (₹)", bold=True)
    ISP_SC = r
    for ci in [2, 3, 4]:
        cl = get_column_letter(ci)
        res(ws, r, ci,
            f"=({cl}{TV_SC_R}*(1+{tgr_ref})"
            f"/(MAX({wacc_ref}-{tgr_ref},0.001)))"
            f"/{sh_ref}",
            FMT_RS)
    r += 1

    lbl(ws, r, 1, "Current Market Price  (₹)")
    for ci in [2, 3, 4]:
        calc(ws, r, ci, "=Inputs!B7", FMT_RS)
    r += 1

    lbl(ws, r, 1, "Upside / Downside  (%)", bold=True)
    for ci in [2, 3, 4]:
        cl = get_column_letter(ci)
        res(ws, r, ci, f"=(B{ISP_SC}-Inputs!B7)/Inputs!B7"
            if ci == 2 else
            f"=({cl}{ISP_SC}-Inputs!B7)/Inputs!B7",
            FMT_PCT)
    r += 1

    return ws, {"ISP_SC": ISP_SC, "NPAT_SC": NPAT_SC}


# ── SHEET 6: SENSITIVITY ANALYSIS ────────────────────────────────────────────
def build_sensitivity(wb, DCF_ROWS):
    ws = wb.create_sheet("Sensitivity Analysis")
    ws.freeze_panes = "B4"

    for col, w in [("A", 18), ("B", 12), ("C", 12), ("D", 12),
                   ("E", 12), ("F", 12), ("G", 12), ("H", 12)]:
        cw(ws, col, w)

    title(ws, 1, 1, "SENSITIVITY ANALYSIS – Implied Share Price (₹)", span=8)
    ws.cell(row=2, column=1,
            value="2-Way Table: WACC (columns, %) vs Terminal Growth Rate (rows, %)").font = BLACK_FONT

    WACCS = [0.08, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14]
    TGRS  = [0.020, 0.025, 0.030, 0.035, 0.040, 0.045, 0.050]

    r = 3
    # Corner
    c = ws.cell(row=r, column=1, value="TGR \\ WACC")
    c.font = BOLD_FONT; c.fill = SECTION_FILL; c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="center")

    BASE_WACC = 0.10
    BASE_TGR  = 0.030

    for ci, w in enumerate(WACCS, 2):
        h = ws.cell(row=r, column=ci, value=w)
        h.number_format = FMT_PCT
        h.border = THIN_BORDER
        h.alignment = Alignment(horizontal="center")
        if abs(w - BASE_WACC) < 0.001:
            h.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor="C00000")
        else:
            h.font = HDR_FONT
            h.fill = HDR_FILL
    r += 1

    # Retrieve FCFF and Shares rows from DCF sheet
    FCFF_R = DCF_ROWS["FCFF_R"]
    SH_R   = DCF_ROWS["SH_R"]

    # Precompute PV of FCFFs: reference DCF B{FCFF_R}:F{FCFF_R} and B{SH_R}
    # Each cell = sum of discounted FCFFs + discounted TV
    for tgr in TGRS:
        tgr_c = ws.cell(row=r, column=1, value=tgr)
        tgr_c.number_format = FMT_PCT
        tgr_c.border = THIN_BORDER
        tgr_c.alignment = Alignment(horizontal="center")
        if abs(tgr - BASE_TGR) < 0.001:
            tgr_c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            tgr_c.fill = PatternFill("solid", fgColor="C00000")
        else:
            tgr_c.font = HDR_FONT
            tgr_c.fill = HDR_FILL

        for ci, wacc in enumerate(WACCS, 2):
            # Build formula: sum of PV of 5 FCFFs + PV of Terminal Value
            pv_parts = "+".join(
                f"'DCF Valuation'!{get_column_letter(2+i)}{FCFF_R}/(1+{wacc})^{i+1}"
                for i in range(5)
            )
            tv_part = (
                f"'DCF Valuation'!F{FCFF_R}*(1+{tgr})"
                f"/(MAX({wacc}-{tgr},0.0001))"
                f"/(1+{wacc})^5"
            )
            formula = f"=({pv_parts}+{tv_part})/'DCF Valuation'!B{SH_R}"

            c2 = ws.cell(row=r, column=ci, value=formula)
            c2.number_format = FMT_RS
            c2.border = THIN_BORDER
            c2.alignment = Alignment(horizontal="right")

            if abs(wacc - BASE_WACC) < 0.001 and abs(tgr - BASE_TGR) < 0.001:
                c2.fill = PatternFill("solid", fgColor="FFFF00")
                c2.font = Font(name="Calibri", size=10, bold=True, color="000000")
            elif abs(wacc - BASE_WACC) < 0.001 or abs(tgr - BASE_TGR) < 0.001:
                c2.fill = PatternFill("solid", fgColor="FFE4B5")
                c2.font = BLACK_FONT
            else:
                c2.fill = CALC_FILL
                c2.font = BLACK_FONT
        r += 1

    r += 1
    section(ws, r, 1, "LEGEND", span=4); r += 1
    leg = [
        ("Red Header",   "Current WACC (10%) or TGR (3%) axis"),
        ("Yellow Cell",  "Base Case (WACC 10%, TGR 3%)"),
        ("Peach Cell",   "Single dimension match"),
    ]
    for lbl_txt, desc in leg:
        lbl(ws, r, 1, lbl_txt)
        ws.cell(row=r, column=2, value=desc).border = THIN_BORDER
        r += 1

    return ws


# ── SHEET 7: DASHBOARD ────────────────────────────────────────────────────────
def build_dashboard(wb, SR, DCF_ROWS, COMPS_DATA, SCEN_ROWS, RATIO_ROWS):
    ws = wb.create_sheet("Dashboard")

    for col, w in [("A", 3), ("B", 25), ("C", 18), ("D", 18),
                   ("E", 18), ("F", 18), ("G", 18), ("H", 18), ("I", 18)]:
        cw(ws, col, w)

    title(ws, 1, 2, "IDBI BANK – INVESTMENT ANALYSIS DASHBOARD", span=8)

    r = 3
    # ── Stock Snapshot ────────────────────────────────────────────────────────
    section(ws, r, 2, "STOCK SNAPSHOT  (April 2026)", span=4); r += 1
    snap = [
        ("Current Market Price",  "=Inputs!B7",  FMT_RS),
        ("Market Cap (₹ Cr)",     "=Inputs!B8",  FMT_INR2),
        ("P/E Ratio",             "=Inputs!B12", "0.00"),
        ("P/B Ratio",             "=Inputs!B13", "0.00"),
        ("EPS – ₹",               "=Inputs!B14", FMT_RS),
        ("Book Value / Share – ₹","=Inputs!B15", FMT_RS),
    ]
    for i, (label, formula, fmt) in enumerate(snap):
        col = 2 + (i % 4) * 2
        row_lbl = r + (i // 4) * 2
        lbl(ws, row_lbl, col, label, bold=True)
        c = ws.cell(row=row_lbl + 1, column=col, value=formula)
        c.font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
        c.fill = INPUT_FILL; c.border = MED_BORDER
        c.number_format = fmt
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_lbl + 1].height = 22
    r += 6

    # ── Valuation Summary ─────────────────────────────────────────────────────
    section(ws, r, 2, "VALUATION SUMMARY", span=8); r += 1
    hdr(ws, r, 2, "Valuation Method"); hdr(ws, r, 3, "Implied Price (₹)")
    hdr(ws, r, 4, "CMP (₹)"); hdr(ws, r, 5, "Upside / Downside %")
    r += 1

    ISP_R = DCF_ROWS["ISP_R"]
    ISP_SC= SCEN_ROWS["ISP_SC"]
    # Comps implied rows
    CIR = COMPS_DATA["COMPS_ISP_ROWS"]

    val_items = [
        ("DCF Valuation",               f"='DCF Valuation'!B{ISP_R}"),
        ("Comps – EV/EBITDA",           f"='Trading Comps'!E{CIR[0]}"),
        ("Comps – EV/Revenue",          f"='Trading Comps'!E{CIR[1]}"),
        ("Comps – P/E",                 f"='Trading Comps'!E{CIR[2]}"),
        ("Comps – P/B",                 f"='Trading Comps'!E{CIR[3]}"),
        ("Scenario – Base Case",        f"='Scenario Analysis'!B{ISP_SC}"),
        ("Scenario – Bull Case",        f"='Scenario Analysis'!C{ISP_SC}"),
        ("Scenario – Bear Case",        f"='Scenario Analysis'!D{ISP_SC}"),
    ]
    for lbl_txt, formula in val_items:
        lbl(ws, r, 2, lbl_txt)
        c = ws.cell(row=r, column=3, value=formula)
        c.font = BLACK_FONT; c.fill = RESULT_FILL; c.border = THIN_BORDER
        c.number_format = FMT_RS; c.alignment = Alignment(horizontal="right")
        calc(ws, r, 4, "=Inputs!B7", FMT_RS)
        res(ws, r, 5, f"=IF(D{r}=0,0,(C{r}-D{r})/D{r})", FMT_PCT)
        r += 1
    r += 1

    # ── Key Financial Metrics ─────────────────────────────────────────────────
    section(ws, r, 2, "KEY FINANCIAL METRICS  (FY 2024-25 Actuals)", span=8); r += 1
    roe_r  = RATIO_ROWS["Return on Equity (ROE) %"]
    roa_r  = RATIO_ROWS["Return on Assets (ROA) %"]
    nim_r  = RATIO_ROWS["Net Interest Margin (NIM) %"]
    gnpa_r = RATIO_ROWS["Gross NPA %"]
    nnpa_r = RATIO_ROWS["Net NPA %"]
    car_r  = RATIO_ROWS["Capital Adequacy Ratio %"]

    metrics = [
        ("Interest Income (₹ Mn)",  "='3-Statements Model'!B" + str(SR["II_R"]),  FMT_INR),
        ("NII (₹ Mn)",              "='3-Statements Model'!B" + str(SR["NII_R"]), FMT_INR),
        ("NPAT (₹ Mn)",             "='3-Statements Model'!B" + str(SR["NPAT_R"]),FMT_INR),
        ("Total Assets (₹ Cr)",     "='3-Statements Model'!B" + str(SR["TA_R"]),  FMT_INR),
        ("Advances (₹ Cr)",         "='3-Statements Model'!B" + str(SR["ADV_R"]), FMT_INR),
        ("Deposits (₹ Cr)",         "='3-Statements Model'!B" + str(SR["DEP_R"]), FMT_INR),
        ("ROE %",                   f"=Inputs!B{roe_r}",  FMT_PCT),
        ("ROA %",                   f"=Inputs!B{roa_r}",  FMT_PCT),
        ("NIM %",                   f"=Inputs!B{nim_r}",  FMT_PCT),
        ("Gross NPA %",             f"=Inputs!B{gnpa_r}", FMT_PCT),
        ("Net NPA %",               f"=Inputs!B{nnpa_r}", FMT_PCT),
        ("CAR %",                   f"=Inputs!B{car_r}",  FMT_PCT),
    ]

    for i, (lbl_txt, formula, fmt) in enumerate(metrics):
        col_off = (i % 4) * 2
        row_off = i // 4
        lbl(ws, r + row_off * 2, 2 + col_off, lbl_txt)
        c = ws.cell(row=r + row_off * 2 + 1, column=2 + col_off, value=formula)
        c.font = BLACK_FONT; c.fill = CALC_FILL; c.border = THIN_BORDER
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    r += 8

    # ── 5-Year Projections ────────────────────────────────────────────────────
    section(ws, r, 2, "5-YEAR PROJECTIONS SUMMARY  (₹ Millions)", span=8); r += 1
    years = ["FY 2024-25", "FY 2025-26", "FY 2026-27", "FY 2027-28", "FY 2028-29"]
    hdr(ws, r, 2, "Metric")
    for i, y in enumerate(years):
        hdr(ws, r, 3 + i, y)
    r += 1

    proj = [
        ("Interest Income",  SR["II_R"]),
        ("NII",              SR["NII_R"]),
        ("NPAT",             SR["NPAT_R"]),
    ]
    for lbl_txt, stmt_r in proj:
        lbl(ws, r, 2, lbl_txt)
        for i in range(5):
            cl = get_column_letter(2 + i)
            c = ws.cell(row=r, column=3 + i,
                        value=f"='3-Statements Model'!{cl}{stmt_r}")
            c.font = BLACK_FONT; c.fill = CALC_FILL; c.border = THIN_BORDER
            c.number_format = FMT_INR
            c.alignment = Alignment(horizontal="right")
        r += 1
    r += 2

    # ── Notes ─────────────────────────────────────────────────────────────────
    section(ws, r, 2, "MODEL NOTES", span=8); r += 1
    notes = [
        "1.  P&L figures in ₹ Millions; Balance Sheet in ₹ Crore.  Conversion: ÷10 for Cr→Mn used where needed.",
        "2.  Blue font = hardcoded inputs (Inputs sheet only). Black font = formulas in all other cells.",
        "3.  Scenario selector (Inputs!B) drives all P&L, Cash Flow, and Valuation outputs dynamically.",
        "4.  DCF uses WACC ≈ 10% (Rf 6.5% + β 0.85 × ERP 7%) and terminal growth rate 3%.",
        "5.  Sensitivity Analysis tests WACC 8%–14% vs Terminal Growth 2%–5% in 2-way table.",
        "6.  Trading Comps based on peer bank data entered in Inputs sheet (April 2026 market data).",
    ]
    for note in notes:
        c = ws.cell(row=r, column=2, value=note)
        c.font = Font(name="Calibri", size=9, italic=True, color="595959")
        c.border = THIN_BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        r += 1

    return ws


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default blank sheet

    print("Building Sheet 1: Inputs …")
    (ws_inp, DRIVER_ROWS, SC_ROW, WACC_R, TGR_ROW,
     IS_ROWS, BS_ROWS, rf_r, erp_r, bet_r, kd_r, de_r, PEER_START, RATIO_ROWS) = build_inputs(wb)

    print("Building Sheet 2: 3-Statements Model …")
    ws_3s, STMT_ROWS = build_3statements(wb, DRIVER_ROWS, SC_ROW, IS_ROWS, BS_ROWS)

    print("Building Sheet 3: DCF Valuation …")
    ws_dcf, DCF_ROWS = build_dcf(wb, STMT_ROWS, DRIVER_ROWS, WACC_R, TGR_ROW, SC_ROW, BS_ROWS)

    print("Building Sheet 4: Trading Comps …")
    ws_comps, COMPS_DATA = build_comps(wb, PEER_START, IS_ROWS, BS_ROWS)

    print("Building Sheet 5: Scenario Analysis …")
    ws_scen, SCEN_ROWS = build_scenarios(wb, STMT_ROWS, DRIVER_ROWS, SC_ROW, IS_ROWS, DCF_ROWS)

    print("Building Sheet 6: Sensitivity Analysis …")
    ws_sens = build_sensitivity(wb, DCF_ROWS)

    print("Building Sheet 7: Dashboard …")
    ws_dash = build_dashboard(wb, STMT_ROWS, DCF_ROWS, COMPS_DATA, SCEN_ROWS, RATIO_ROWS)

    # Tab colours
    ws_inp.sheet_properties.tabColor   = "1F4E79"
    ws_3s.sheet_properties.tabColor    = "2E75B6"
    ws_dcf.sheet_properties.tabColor   = "70AD47"
    ws_comps.sheet_properties.tabColor = "ED7D31"
    ws_scen.sheet_properties.tabColor  = "FFC000"
    ws_sens.sheet_properties.tabColor  = "FF0000"
    ws_dash.sheet_properties.tabColor  = "7030A0"

    wb.active = ws_dash

    out = ("/home/runner/work/IDBI-Bank-Financial-Model/"
           "IDBI-Bank-Financial-Model/IDBI_Bank_Financial_Model.xlsx")
    wb.save(out)
    size = os.path.getsize(out)
    print(f"\n✅  Model saved → {out}")
    print(f"    File size : {size:,} bytes ({size/1024:.1f} KB)")


if __name__ == "__main__":
    main()
