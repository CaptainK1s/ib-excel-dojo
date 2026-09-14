# IB Excel Dojo

Practice file for investment-banking Excel.

I built this to get faster at the actual job: moving around a model without a mouse, keeping inputs obvious, and making sure the statements still balance after you change a driver.

Blue font = hardcoded input  
Black font = formula  
Green font = link to another sheet  
Yellow fill = the cell you are supposed to toggle

Not a course. Not investment advice. Figures are hypothetical and belong to a fake company (Apex Industrial).

## Files

| | |
|---|---|
| [IB_Excel_Dojo.xlsx](IB_Excel_Dojo.xlsx) | The workbook. Start on tab `00_Start`. |
| `index.html` | Simple landing page if Pages is on. |
| `src/build_ib_excel_dojo.py` | Regenerates the xlsx. |

## Tabs

- `00_Start` — color rules and how the file is meant to be used
- `01_Shortcuts` — the shortcuts that matter, marked by week
- `02_Format_Gym` — messy extract on the left, target look on the right
- `03_Formula_Dojo` — 20 drills. Yellow cells are yours
- `04_3Statement` — mini IS / BS / CFS for Apex. Change growth and watch it move
- `05_Sensitivity` — WACC × exit multiple, plus a downside / base / upside switch
- `06_Comps` — price → equity → EV → implied price
- `07_Checks` — if this tab is red, the file does not leave the desk
- `08_Plan` — two weeks of short sessions

## How I use it

Shortcuts and the format gym first. Then break tab `04` on purpose (cut growth, jack capex, push the payout over 100%) and see which checks go red. That is the point.

## License

Free to copy for practice. Leave the disclaimer on it. Do not treat Apex as a real name.
