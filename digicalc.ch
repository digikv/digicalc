/* DigiCalc - common definitions */

#ifndef DIGICALC_CH
#define DIGICALC_CH

#define DC_VERSION   "0.1"

#define DC_MAXCOL    16384          /* XFD, as Excel */
#define DC_MAXROW    1048576
#define DC_KEYMUL    32768          /* cell key = nRow * DC_KEYMUL + nCol */

/* a cell: an array */
#define CELL_INPUT   1              /* what is shown in the entry line: "=SUM(A1:A5)", "'Text", "12" */
#define CELL_VALUE   2              /* number, text, logical, DCErr, NIL */
#define CELL_AST     3              /* the formula, NIL for a constant */
#define CELL_FMT     4              /* "" = the sheet's; "F2" "C2" ",2" "P1" "G" "D1".."D5" "T" "H" */
#define CELL_ALIGN   5              /* "" = numbers right, text left; "L" "R" "C" */
#define CELL_RAW     6              /* a formula from a file that DigiCalc does not parse
                                       (Sheet2!A1 ...): kept as text, written back; the
                                       value is the one in the file */
#define CELL_LEN     6

#endif
