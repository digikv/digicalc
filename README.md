# DigiCalc

![DigiCalc](digicalc.png)

**A spreadsheet in character mode, in the manner of Lotus 1-2-3**, written in
[Harbour](https://harbour.github.io/) on the [Harbour-CUI](https://github.com/digikv/Harbour-CUI)
library. Windows (gtwvt) and Linux (X11, gtxwc).

- the 1-2-3 screen: the control panel (`A1: (F2) [W12] @SUM(B1..B5)`, the entry line,
  the mode: READY, VALUE, LABEL, EDIT, POINT, MENU), the `/` menu, the cell pointer;
  and a menu bar for the mouse
- formulas in **both syntaxes**: Excel `=SUM(A1:A5)`, `AND(A1>0,B1<5)` and
  1-2-3 `@SUM(A1..A5)`, `+A1+B1`, `A1>0#AND#B1<5` — type either; they are shown the
  Excel or the 1-2-3 way (`/ Worksheet Global Syntax`)
- POINT: while a formula is typed, the arrows point at a cell, `.` or `:` anchors a range
- recalculation of just the cells that depend on a change; circular references (`#CIRC!`)
- `.xlsx` (read and write, several sheets, number formats, alignment, column widths)
  through the [XLSX class](https://github.com/digikv/xlsx); `.csv` import and export
- Copy and Move with the addresses following, insert / delete rows and columns,
  column widths, formats, names of ranges, sort, fill

DigiCalc is written from scratch. It takes nothing from Lotus 1-2-3 but the way it
works on the screen; "1-2-3" is a trademark of its owner.

## Keys

| | |
|---|---|
| typing | a number, a text or a formula; Enter or an arrow stores it (the arrow then moves) |
| `'` `"` `^` | a text left / right / centered (1-2-3 label prefixes) |
| `/` | the 1-2-3 menu: Worksheet Range Copy Move File Data Quit |
| F10, Alt + letter | the menu bar |
| F2 | edit the cell |
| F5 | go to a cell |
| F9 | recalculate (with `/ Worksheet Global Recalculation Manual`) |
| Del | erase the cell or the selected range |
| arrows, PgUp, PgDn | move; Shift + arrows select a range |
| Tab, Shift-Tab | a screen right / left |
| Home, End + arrow, End Home | A1, the edge of the data, the last cell |
| Ctrl-PgDn, Ctrl-PgUp | the next / previous sheet |
| Esc | drop the entry, a menu level back |
| Alt-X | exit |

## Formulas

Operators: `+ - * / ^ & %`, `= <> < > <= >=`, `#AND# #OR# #NOT#` (1-2-3).
Addresses: `A1`, `$A$1`, `A1:B5` / `A1..B5`, names of ranges.

Functions (the 1-2-3 names too: `@AVG`, `@STRING`, `@ISERR`, `@N`, `@S` …):
SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK PRODUCT STDEV STDEVP VAR VARP
SUMIF COUNTIF AVERAGEIF · IF IFERROR AND OR NOT CHOOSE ISERROR ISNA ISNUMBER ISTEXT
ISBLANK NA · ABS INT ROUND ROUNDUP ROUNDDOWN TRUNC MOD SQRT POWER EXP LN LOG LOG10
SIN COS TAN ASIN ACOS ATAN ATAN2 PI RAND SIGN · LEN LEFT RIGHT MID UPPER LOWER PROPER
TRIM CONCATENATE CONCAT REPT EXACT FIND SEARCH SUBSTITUTE CHAR CODE VALUE TEXT N S ·
VLOOKUP HLOOKUP INDEX MATCH ROWS COLUMNS · TODAY NOW DATE YEAR MONTH DAY WEEKDAY HOUR
MINUTE SECOND · PMT PV FV NPER NPV.

The 1-2-3 functions whose arguments differ from Excel's are converted when typed:
`@VLOOKUP(x,r,0)` (offsets from 0) is `VLOOKUP(x,r,1,TRUE)`, `@INDEX(r,col,row)` is
`INDEX(r,row+1,col+1)`, `@PMT(principal,rate,term)` is `-PMT(rate,term,principal)`.

## Building

Next to each other in one folder:

```
Harbour-CUI/     https://github.com/digikv/Harbour-CUI  (the hbtui library built)
xlsx/            https://github.com/digikv/xlsx          (for .xlsx; without it only .csv)
digicalc/        this repository
```

Harbour 3.2.1 (or newer) with hbmk2.

- Windows: `build.bat` (set `HBMK2` when hbmk2 is not in `D:\tools\harbour-core\bin\win\mingw`)
- Linux: `./build.sh` (X11 development files: `sudo apt install libx11-dev`; when Harbour
  was built without contrib libraries, `HBLIB` points at a folder with `hbct`, `hbmzip`,
  `hbziparc`)

The tests of the core (formulas, recalculation, structure, files):
`cd tests && hbmk2 dctest.hbp && ./dctest`

## License

Copyright (c) 2026 Srđan Dragojlović.
Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial license from the
author — see LICENSE.
