# IDBI Bank – Comprehensive Financial Model

A professional investment-grade Excel financial model for **IDBI Bank** built with verified screener data as of **April 2026**.

---

## 📥 Download

**[`IDBI_Bank_Financial_Model.xlsx`](IDBI_Bank_Financial_Model.xlsx)** – Open in Microsoft Excel or LibreOffice Calc

---

## 📊 Model Overview

| Sheet | Description |
|-------|-------------|
| **1 · Inputs** | Single source of truth – all hard-coded data in blue font |
| **2 · 3-Statements Model** | Linked P&L, Balance Sheet, Cash Flow (FY 2024-25 → FY 2028-29) |
| **3 · DCF Valuation** | FCFF-based valuation with WACC and Gordon Growth Terminal Value |
| **4 · Trading Comps** | 5-peer analysis (HDFC, ICICI, Axis, SBI, Kotak) with EV/EBITDA, P/E, P/B |
| **5 · Scenario Analysis** | Base / Bull / Bear with single scenario selector driving all outputs |
| **6 · Sensitivity Analysis** | 2-way table: WACC (8-14%) vs Terminal Growth Rate (2-5%) |
| **7 · Dashboard** | Valuation summary, key metrics, 5-year projections |

---

## 🔢 Verified Financial Data (FY 2024-25 Base Year)

| Metric | Value |
|--------|-------|
| Stock Price (CMP, Apr 2026) | ₹69.55 |
| Market Cap | ₹74,782.96 Crore |
| Shares Outstanding | 10.75 Billion |
| P/E Ratio | 8.05× |
| P/B Ratio | 1.12× |
| EPS (TTM) | ₹8.64 |
| Interest Income | ₹2,89,171 Mn |
| Net Interest Income | ₹1,46,605 Mn |
| Net Profit After Tax | ₹76,307 Mn |
| Total Assets | ₹4,12,962 Crore |
| Advances | ₹2,18,212 Crore |
| Deposits | ₹3,09,975 Crore |
| ROE | 20.15% |
| Gross NPA | 2.98% |

---

## ⚙️ How to Use

1. **Download** the Excel file and open in Microsoft Excel
2. **Navigate to the `Inputs` sheet** – change any blue-font assumption cell
3. **All three statements** (P&L, Balance Sheet, Cash Flow) update automatically
4. **Change the Scenario** in cell `B70` of the Inputs sheet (`Base` / `Bull` / `Bear`)
   - The dropdown selector drives the entire P&L and valuation
5. **Review Scenario Analysis** (Sheet 5) for side-by-side comparison
6. **Review Sensitivity Analysis** (Sheet 6) for implied share price at each WACC/TGR combination
7. **Dashboard** (Sheet 7) shows a one-page valuation summary

---

## 🏗️ Model Architecture

```
Inputs Sheet  <-- Single Source of Truth (all blue-font inputs)
     |
     +---> 3-Statements Model  (P&L -> Balance Sheet -> Cash Flow)
     |              |
     |              +---> DCF Valuation  (FCFF + WACC + Terminal Value -> Implied Price)
     |
     +---> Trading Comps  (Peer multiples -> IDBI implied range)
     |
     +---> Scenario Analysis  (Base/Bull/Bear -> P&L + Implied Price)
     |
     +---> Sensitivity Analysis  (WACC x TGR grid)

Dashboard  <-- References all sheets above for one-page summary
```

---

## 📐 WACC Build-up

| Component | Value |
|-----------|-------|
| Risk-Free Rate (Rf) | 6.5% |
| Equity Risk Premium (ERP) | 7.0% |
| Beta (beta) | 0.85 |
| Cost of Equity (CAPM) | **12.45%** |
| Cost of Debt (pre-tax) | 7.5% |
| Tax Rate | 30% |
| Target D/E Ratio | 6 : 1 |
| **WACC** | **~9.45%** |
| Terminal Growth Rate | 3.0% |

---

## 🔵 Formatting Convention

| Colour | Meaning |
|--------|---------|
| **Blue font** | Hard-coded inputs – edit freely |
| **Black font** | Formula-driven – do not edit |
| Light blue background | Input cells |
| Light yellow background | Calculated cells |
| Light green background | Result / output cells |

---

## 🗂️ Files

| File | Description |
|------|-------------|
| `IDBI_Bank_Financial_Model.xlsx` | The Excel workbook (download and open) |
| `create_model.py` | Python script that generates the workbook |
| `README.md` | This documentation |

---

## 📋 Requirements to Regenerate

```bash
pip install openpyxl
python create_model.py
```

The script produces `IDBI_Bank_Financial_Model.xlsx` in the current directory.

---

*Data sourced from Screener.in and public disclosures as of April 2026.
For educational and analytical purposes only.*
