# Check results

Workbook recalculated in LibreOffice: 4842 formulas, 0 errors.

Independent recalculation in pandas (`tools/check_workbook.py`):

- PASS: Real price, every quarter (to £0.01)
- PASS: Fall from peak, every quarter
- PASS: Annual cash price
- PASS: Annual real price
- PASS: All Checks sheet tests PASS
- PASS: Monthly repayment = annuity formula
