# Pension Asset Management Project (Excel + VBA)

## Overview
Educational portfolio project simulating pension-asset reporting with Excel tables, formulas, KPIs, charts, cash-flow records, data-quality checks, and VBA automation.

**All holdings, prices, FX rates, benchmark returns, and cash flows are synthetic examples.** This is not real UniCredit/HypoVereinsbank data, a live market report, or investment advice.

## Workbook contents
- `Holdings`: monthly holdings, asset class, region, currency, units, local price, illustrative FX rates, market value, previous value, P/L, returns, benchmark, excess return, and data-check status.
- `Dashboard`: reporting-date selector, market value, previous value, change in value, simplified portfolio return, data-check count, period cash flows, asset-class chart, and portfolio-value trend.
- `Cash_Flows`: illustrative contributions, pension benefit payments, and fees.
- `Instructions`: guidance and calculation limitations.

## VBA automation
Import `Pension_Asset_Management_Automation.bas` in Excel VBA (`Alt+F11` → File → Import File) after saving a copy of the workbook as `.xlsm`.

Macros:
1. `RefreshPensionDashboard` — requests workbook refresh and recalculates formulas.
2. `ValidatePensionHoldings` — checks key fields and flags rows needing review.
3. `ExportDashboardToPDF` — exports the Dashboard worksheet to a one-page landscape PDF.
4. `CreateMonthlySnapshot` — appends current dashboard KPI values to a `Snapshot_Log` worksheet.

To add buttons: Developer → Insert → Button (Form Control) → draw the button → assign a macro. Enable macros only for code you trust and have reviewed.

## Calculations
- Market Value (€) = Units × Price (Local) ÷ FX Rate (Local per EUR).
- Change in Value (€) = Current Market Value − Previous Value.
- Holding Return (%) = Change in Value ÷ Previous Value.
- Excess Return (%) = Holding Return − Benchmark Return.
- Dashboard portfolio return is a simplified change/previous-value ratio. It does not adjust for contributions, withdrawals, or cash-flow timing, so it is not a formal time-weighted or money-weighted return.

## Skills demonstrated
Excel tables and formulas, `SUMIFS`, `COUNTIF`, data validation, KPI reporting, charts, basic VBA, error handling, data-quality validation, PDF export, and snapshot logging.

## Limitations
Synthetic data only; FX and benchmark values are illustrative. The validation macro is a basic demonstration, not a full financial control framework. Future enhancements could include reconciliation checks, formal cash-flow-adjusted performance calculations, audit logging, and refresh-completion checks.
