/** @file digicalc.prg
 * DigiCalc - a spreadsheet in character mode, in the manner of Lotus 1-2-3,
 * on Harbour-CUI. Formulas in Excel (=SUM(A1:A5)) or 1-2-3 (@SUM(A1..A5))
 * syntax; .xlsx through the XLSX class, .csv.
 *
 *   digicalc [file.xlsx | file.csv]
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "hbclass.ch"
#include "hbtui.ch"
#include "inkey.ch"
#include "hbgtinfo.ch"
#include "digicalc.ch"

REQUEST HB_CODEPAGE_UTF8EX

STATIC s_oApp, s_oMain, s_oGrid, s_oStat, s_oBar
STATIC s_aSheets := {}              /* the workbook */
STATIC s_nSheet := 1
STATIC s_cFile := ""
STATIC s_aPos := {}                 /* the pointer of each sheet: { nCol, nRow, nTop, nLeft } */
STATIC s_aClip := NIL               /* Edit Copy: { oSheet, aArea } */

PROCEDURE Main( ... )

   LOCAL cArg

   hb_cdpSelect( "UTF8EX" )

   s_oApp := HTApplication():new()
   ht_cuiSetupDisplay( s_oApp:rows(), s_oApp:cols() )
   hb_gtInfo( HB_GTI_WINTITLE, "DigiCalc" )
   hb_gtInfo( HB_GTI_ICONRES, 1 )                     /* digicalc.ico (digicalc.rc), Windows */

   s_oMain := DCMain():new()
   s_oMain:setWindowTitle( "DigiCalc" )
   s_oMain:show()

   s_oGrid := HTSheet():new( s_oMain )
   s_oGrid:bStatus := {|| dc_Status() }
   s_oGrid:bSlash := {|| DC_MenuOpen( s_oGrid, dc_SlashMenu() ) }
   s_oGrid:bSheet := {| n | dc_GoSheet( s_nSheet + n ) }
   s_oStat := HTStatusBar():new( s_oMain )
   s_oStat:addSection( "" )
   dc_SetBook( { DCSheet():new( "Sheet1" ) }, "" )
   dc_Arrange()
   dc_BuildMenu()

   FOR EACH cArg IN hb_AParams()
      IF !( Left( cArg, 1 ) == "-" )
         dc_Open( cArg )
         EXIT
      ENDIF
   NEXT

   s_oGrid:setFocus()
   dc_Status()
   s_oApp:exec()

RETURN

/*------------------------------------------------------------ the main window */

CLASS DCMain FROM HTMainWindow
   METHOD keyEvent( keyEvent )
   METHOD closeEvent( closeEvent )
   METHOD screenResized()
ENDCLASS

METHOD PROCEDURE keyEvent( keyEvent ) CLASS DCMain
   IF keyEvent:key == K_ALT_X
      keyEvent:accept()
      dc_Exit()
      RETURN
   ENDIF
   ::super:keyEvent( keyEvent )
RETURN

METHOD PROCEDURE closeEvent( closeEvent ) CLASS DCMain
   IF dc_AskSave()
      ::super:closeEvent( closeEvent )
   ENDIF
RETURN

METHOD PROCEDURE screenResized() CLASS DCMain
   ::super:screenResized()
   dc_Arrange()
RETURN

/* the sheet from under the menu bar down to the status line */
STATIC PROCEDURE dc_Arrange()
   s_oGrid:setGeometry( 0, 0, s_oMain:width, Max( 6, s_oMain:height - 2 ) )
   s_oStat:setGeometry( 0, s_oMain:height - 2, s_oMain:width, 1 )
RETURN

STATIC FUNCTION dc_Sheet()
RETURN s_aSheets[ s_nSheet ]

STATIC FUNCTION dc_Modified()
   LOCAL o
   FOR EACH o IN s_aSheets
      IF o:lModified
         RETURN .T.
      ENDIF
   NEXT
RETURN .F.

STATIC PROCEDURE dc_Status()
   LOCAL aA := s_oGrid:selArea(), c, cRight
   c := " " + IIF( Empty( s_cFile ), "(new)", hb_FNameNameExt( s_cFile ) ) + IIF( dc_Modified(), " *", "" ) + ;
        "   " + dc_Sheet():cName + IIF( Len( s_aSheets ) > 1, " (" + hb_ntos( s_nSheet ) + "/" + hb_ntos( Len( s_aSheets ) ) + ")", "" )
   IF ! HB_ISNIL( s_oGrid:nAncCol )
      c += "   " + DC_AreaText( aA, s_oGrid:lLotus )
   ENDIF
   cRight := IIF( s_oGrid:lEnd, "END  ", "" ) + IIF( dc_Sheet():lAutoCalc, "", "CALC  " ) + ;
             IIF( s_oGrid:lLotus, "1-2-3", "Excel" ) + "   / menu   F1 help "
   s_oStat:setSection( 1, PadR( c, Max( 1, s_oMain:width - Len( cRight ) ) ) + cRight )
   hb_gtInfo( HB_GTI_WINTITLE, "DigiCalc" + IIF( Empty( s_cFile ), "", " - " + hb_FNameNameExt( s_cFile ) ) )
RETURN

/* a new workbook in the window */
STATIC PROCEDURE dc_SetBook( aSheets, cFile )
   s_aSheets := aSheets
   s_cFile := cFile
   s_aPos := {}
   AEval( aSheets, {|| AAdd( s_aPos, { 1, 1, 1, 1 } ) } )
   s_nSheet := 1
   s_oGrid:oSheet := s_aSheets[ 1 ]
   s_oGrid:nCol := s_oGrid:nRow := s_oGrid:nTop := s_oGrid:nLeft := 1
   s_oGrid:clearSel()
   s_oGrid:cancel()
RETURN

/* Ctrl-PgDn / Ctrl-PgUp: the next / previous sheet, each keeps its pointer */
STATIC PROCEDURE dc_GoSheet( n )
   IF n < 1 .OR. n > Len( s_aSheets ) .OR. n == s_nSheet
      RETURN
   ENDIF
   s_aPos[ s_nSheet ] := { s_oGrid:nCol, s_oGrid:nRow, s_oGrid:nTop, s_oGrid:nLeft }
   s_nSheet := n
   s_oGrid:oSheet := s_aSheets[ n ]
   s_oGrid:nCol := s_aPos[ n ][ 1 ]
   s_oGrid:nRow := s_aPos[ n ][ 2 ]
   s_oGrid:nTop := s_aPos[ n ][ 3 ]
   s_oGrid:nLeft := s_aPos[ n ][ 4 ]
   s_oGrid:clearSel()
RETURN

/*------------------------------------------------------------ menus */

/* the / menu: the same commands as the menu bar, the 1-2-3 way */
STATIC FUNCTION dc_SlashMenu()

   LOCAL aFmt := { ;
      { "Fixed",    "A fixed number of decimal places: 1234.50",      {|| dc_Format( "F" ) } }, ;
      { "Sci",      "Scientific: 1.23E+03",                           {|| dc_Format( "S" ) } }, ;
      { "Currency", "The currency sign, thousands, decimals",          {|| dc_Format( "C" ) } }, ;
      { ",",        "Thousands separated, decimals: 1,234.50",        {|| dc_Format( "," ) } }, ;
      { "General",  "As the number is",                               {|| dc_SetFormat( "G" ) } }, ;
      { "Percent",  "Times 100 with %: 12.5%",                         {|| dc_Format( "P" ) } }, ;
      { "Date",     "1 dd.mm.yyyy  2 dd-mmm-yy  3 dd.mm.  4 mm.yyyy  5 hh:mm", {|| dc_DateFormat() } }, ;
      { "Text",     "Formulas shown as text",                         {|| dc_SetFormat( "T" ) } }, ;
      { "Hidden",   "The cells show nothing",                         {|| dc_SetFormat( "H" ) } }, ;
      { "Reset",    "Back to the global format",                      {|| dc_SetFormat( "" ) } } }

RETURN { ;
   { "Worksheet", "Insert  Delete  Column  Global  Erase  Sheet", { ;
      { "Insert", "Insert columns or rows", { ;
         { "Column", "Columns before the selected ones", {|| dc_InsDel( .F., .T. ) } }, ;
         { "Row",    "Rows above the selected ones",     {|| dc_InsDel( .T., .T. ) } } } }, ;
      { "Delete", "Delete columns or rows", { ;
         { "Column", "Delete the selected columns", {|| dc_InsDel( .F., .F. ) } }, ;
         { "Row",    "Delete the selected rows",    {|| dc_InsDel( .T., .F. ) } } } }, ;
      { "Column", "Set-Width  Reset-Width  Fit", { ;
         { "Set-Width",   "The width of the selected columns", {|| dc_ColWidth() } }, ;
         { "Reset-Width", "Back to the global width",          {|| dc_ColReset() } }, ;
         { "Fit",         "As wide as the longest value",      {|| dc_ColFit() } } } }, ;
      { "Global", "Format  Column-Width  Recalculation  Syntax", { ;
         { "Format",        "The format of cells without their own", {|| DC_MenuOpen( s_oGrid, { { "Format", "Global format", dc_GlobalFmtMenu() } } ) } }, ;
         { "Column-Width",  "The width of columns without their own", {|| dc_GlobalWidth() } }, ;
         { "Recalculation", "Automatic  Manual", { ;
            { "Automatic", "Formulas follow every change",    {|| dc_Sheet():lAutoCalc := .T., dc_Sheet():recalcAll() } }, ;
            { "Manual",    "Only F9 recalculates (CALC)",     {|| dc_Sheet():lAutoCalc := .F. } } } }, ;
         { "Syntax", "Excel  1-2-3: how formulas are shown", { ;
            { "Excel", "=SUM(A1:A5)",  {|| s_oGrid:lLotus := .F. } }, ;
            { "1-2-3", "@SUM(A1..A5)", {|| s_oGrid:lLotus := .T. } } } } } }, ;
      { "Erase", "Erase the whole sheet", { ;
         { "No",  "Keep it",              {|| NIL } }, ;
         { "Yes", "Erase every cell",     {|| dc_EraseSheet() } } } }, ;
      { "Sheet", "Insert  Delete  Rename  Next  Previous", { ;
         { "Insert",   "A new sheet after this one",  {|| dc_SheetInsert() } }, ;
         { "Delete",   "Delete this sheet",           {|| dc_SheetDelete() } }, ;
         { "Rename",   "The name of this sheet",      {|| dc_SheetRename() } }, ;
         { "Next",     "Ctrl-PgDn",                   {|| dc_GoSheet( s_nSheet + 1 ) } }, ;
         { "Previous", "Ctrl-PgUp",                   {|| dc_GoSheet( s_nSheet - 1 ) } } } } } }, ;
   { "Range", "Format  Label  Erase  Name", { ;
      { "Format", "Fixed  Sci  Currency  ,  General  Percent  Date  Text  Hidden  Reset", aFmt }, ;
      { "Label", "Left  Right  Center", { ;
         { "Left",   "Labels to the left",  {|| dc_Align( "L" ) } }, ;
         { "Right",  "Labels to the right", {|| dc_Align( "R" ) } }, ;
         { "Center", "Labels centered",     {|| dc_Align( "C" ) } } } }, ;
      { "Erase", "Erase a range", {|| dc_EraseRange() } }, ;
      { "Name", "Create  Delete", { ;
         { "Create", "A name for a range: =SUM(PRICES)", {|| dc_NameCreate() } }, ;
         { "Delete", "Delete a name",                    {|| dc_NameDelete() } } } } } }, ;
   { "Copy", "Copy a range; relative addresses follow", {|| dc_Copy() } }, ;
   { "Move", "Move a range; every address of it follows", {|| dc_Move() } }, ;
   { "File", "Retrieve  Save  Import  Export  New", { ;
      { "Retrieve", "Open a workbook (.xlsx, .csv)",  {|| dc_CmdOpen() } }, ;
      { "Save",     "Save the workbook (.xlsx)",      {|| dc_Save( .F. ) } }, ;
      { "Import",   "A .csv file as a new workbook",  {|| dc_CmdOpen( "*.csv;*.txt" ) } }, ;
      { "Export",   "This sheet as .csv",             {|| dc_Export() } }, ;
      { "New",      "An empty workbook",              {|| dc_New() } } } }, ;
   { "Data", "Fill  Sort", { ;
      { "Fill", "Numbers in a range: start, step", {|| dc_Fill() } }, ;
      { "Sort", "Ascending  Descending", { ;
         { "Ascending",  "Sort the rows of a range", {|| dc_Sort( .F. ) } }, ;
         { "Descending", "Sort the rows of a range", {|| dc_Sort( .T. ) } } } } } }, ;
   { "Quit", "End DigiCalc", { ;
      { "No",  "Back to the sheet", {|| NIL } }, ;
      { "Yes", "End DigiCalc",      {|| dc_Exit() } } } } }

STATIC FUNCTION dc_GlobalFmtMenu()
RETURN { ;
   { "Fixed",    "",  {|| dc_GlobalFmt( "F" ) } }, ;
   { "Sci",      "",  {|| dc_GlobalFmt( "S" ) } }, ;
   { "Currency", "",  {|| dc_GlobalFmt( "C" ) } }, ;
   { ",",        "",  {|| dc_GlobalFmt( "," ) } }, ;
   { "General",  "",  {|| dc_Sheet():cDefFmt := "G" } }, ;
   { "Percent",  "",  {|| dc_GlobalFmt( "P" ) } } }

STATIC PROCEDURE dc_BuildMenu()

   LOCAL oMenu

   s_oBar := s_oMain:menuBar()
   s_oBar:setBarRow( 0 )

   oMenu := s_oBar:addMenu( "&File" )
   dc_Item( oMenu, "&New", "", {|| dc_New() } )
   dc_Item( oMenu, "&Open...", "", {|| dc_CmdOpen() } )
   dc_Item( oMenu, "&Save", "", {|| dc_Save( .F. ) } )
   dc_Item( oMenu, "Save &as...", "", {|| dc_Save( .T. ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "&Import .csv...", "", {|| dc_CmdOpen( "*.csv;*.txt" ) } )
   dc_Item( oMenu, "&Export .csv...", "", {|| dc_Export() } )
   oMenu:addSeparator()
   dc_Item( oMenu, "E&xit", "Alt-X", {|| dc_Exit() } )

   oMenu := s_oBar:addMenu( "&Edit" )
   dc_Item( oMenu, "&Edit cell", "F2", {|| s_oGrid:editCell() } )
   dc_Item( oMenu, "&Copy", "", {|| s_aClip := { dc_Sheet(), s_oGrid:selArea() }, s_oGrid:message( "Copied: " + DC_AreaText( s_oGrid:selArea() ) ) } )
   dc_Item( oMenu, "&Paste", "", {|| dc_Paste() } )
   dc_Item( oMenu, "Copy &to...", "/C", {|| dc_Copy() } )
   dc_Item( oMenu, "&Move to...", "/M", {|| dc_Move() } )
   dc_Item( oMenu, "E&rase", "Del", {|| DC_EraseArea( dc_Sheet(), s_oGrid:selArea() ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "Insert &rows", "", {|| dc_InsDel( .T., .T. ) } )
   dc_Item( oMenu, "Insert co&lumns", "", {|| dc_InsDel( .F., .T. ) } )
   dc_Item( oMenu, "&Delete rows", "", {|| dc_InsDel( .T., .F. ) } )
   dc_Item( oMenu, "Delete col&umns", "", {|| dc_InsDel( .F., .F. ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "&Go to", "F5", {|| s_oGrid:keyEvent( HTKeyEvent():new( K_F5 ) ) } )
   dc_Item( oMenu, "Recalc&ulate", "F9", {|| dc_Sheet():recalcAll() } )

   oMenu := s_oBar:addMenu( "F&ormat" )
   dc_Item( oMenu, "&Fixed...", "", {|| dc_Format( "F" ) } )
   dc_Item( oMenu, "&Comma (1,234.50)...", "", {|| dc_Format( "," ) } )
   dc_Item( oMenu, "C&urrency...", "", {|| dc_Format( "C" ) } )
   dc_Item( oMenu, "&Percent...", "", {|| dc_Format( "P" ) } )
   dc_Item( oMenu, "&Scientific...", "", {|| dc_Format( "S" ) } )
   dc_Item( oMenu, "&Date...", "", {|| dc_DateFormat() } )
   dc_Item( oMenu, "&General", "", {|| dc_SetFormat( "G" ) } )
   dc_Item( oMenu, "&Reset", "", {|| dc_SetFormat( "" ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "Align &left", "", {|| dc_Align( "L" ) } )
   dc_Item( oMenu, "Align ri&ght", "", {|| dc_Align( "R" ) } )
   dc_Item( oMenu, "Ce&nter", "", {|| dc_Align( "C" ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "Column &width...", "", {|| dc_ColWidth() } )
   dc_Item( oMenu, "&AutoFit columns", "", {|| dc_ColFit() } )

   oMenu := s_oBar:addMenu( "&Data" )
   dc_Item( oMenu, "Sort &ascending...", "", {|| dc_Sort( .F. ) } )
   dc_Item( oMenu, "Sort &descending...", "", {|| dc_Sort( .T. ) } )
   dc_Item( oMenu, "&Fill...", "", {|| dc_Fill() } )
   dc_Item( oMenu, "&Name a range...", "", {|| dc_NameCreate() } )

   oMenu := s_oBar:addMenu( "&View" )
   dc_Item( oMenu, "Formulas as &1-2-3", "", {|| s_oGrid:lLotus := .T. } )
   dc_Item( oMenu, "Formulas as &Excel", "", {|| s_oGrid:lLotus := .F. } )
   oMenu:addSeparator()
   dc_Item( oMenu, "&Next sheet", "Ctrl-PgDn", {|| dc_GoSheet( s_nSheet + 1 ) } )
   dc_Item( oMenu, "&Previous sheet", "Ctrl-PgUp", {|| dc_GoSheet( s_nSheet - 1 ) } )
   dc_Item( oMenu, "&Insert sheet", "", {|| dc_SheetInsert() } )
   dc_Item( oMenu, "&Rename sheet...", "", {|| dc_SheetRename() } )

   oMenu := s_oBar:addMenu( "&Help" )
   dc_Item( oMenu, "&Keys", "F1", {|| dc_Help() }, K_F1 )
   dc_Item( oMenu, "&About", "", {|| dc_About() } )

RETURN

STATIC PROCEDURE dc_Item( oMenu, cText, cKey, bAction, nKey )
   LOCAL oAct := oMenu:addAction( PadR( cText, IIF( "&" $ cText, 23, 22 ) ) + PadL( cKey, 10 ) )
   IF ! HB_ISNIL( nKey )
      oAct:setShortcut( nKey )
   ENDIF
   oAct:onTriggered := {|| Eval( bAction ), s_oGrid:setFocus(), s_oGrid:refresh() }
RETURN

/*------------------------------------------------------------ asking */

/* a range in the entry line (the selection offered; the arrows point) */
STATIC PROCEDURE dc_AskRange( cPrompt, bDone )
   s_oGrid:prompt( cPrompt + ": ", DC_AreaText( s_oGrid:selArea(), s_oGrid:lLotus ), ;
      {| c, a | a := DC_ParseArea( c ), IIF( HB_ISNIL( a ), s_oGrid:message( "Not a range: " + c ), Eval( bDone, a ) ) }, .T. )
RETURN

STATIC PROCEDURE dc_AskNum( cPrompt, nDefault, nMin, nMax, bDone )
   s_oGrid:prompt( cPrompt + ": ", hb_ntos( nDefault ), ;
      {| c | IIF( ! DC_IsNumText( c ) .OR. DC_Val( c ) < nMin .OR. DC_Val( c ) > nMax, ;
                  s_oGrid:message( "A number from " + hb_ntos( nMin ) + " to " + hb_ntos( nMax ) + ", please" ), ;
                  Eval( bDone, DC_Val( c ) ) ) } )
RETURN

/*------------------------------------------------------------ commands */

/* a format with decimals: F , C P S */
STATIC PROCEDURE dc_Format( cKind )
   dc_AskNum( "Decimal places (0-15)", 2, 0, 15, {| n | dc_SetFormat( cKind + hb_ntos( Int( n ) ) ) } )
RETURN

STATIC PROCEDURE dc_DateFormat()
   dc_AskNum( "Date format 1 dd.mm.yyyy 2 dd-mmm-yy 3 dd.mm. 4 mm.yyyy 5 hh:mm", 1, 1, 5, ;
      {| n | dc_SetFormat( "D" + hb_ntos( Int( n ) ) ) } )
RETURN

STATIC PROCEDURE dc_SetFormat( cFmt )
   dc_AskRange( "Range to format", {| a | dc_ForArea( a, {| nC, nR | dc_Sheet():setFormat( nC, nR, cFmt ) } ) } )
RETURN

/* bCell( nCol, nRow ) for every cell of the range */
STATIC PROCEDURE dc_ForArea( a, bCell )
   LOCAL nC, nR
   FOR nR := a[ 2 ] TO a[ 4 ]
      FOR nC := a[ 1 ] TO a[ 3 ]
         Eval( bCell, nC, nR )
      NEXT
   NEXT
RETURN

STATIC PROCEDURE dc_GlobalFmt( cKind )
   dc_AskNum( "Decimal places (0-15)", 2, 0, 15, {| n | dc_Sheet():cDefFmt := cKind + hb_ntos( Int( n ) ) } )
RETURN

STATIC PROCEDURE dc_Align( cAlign )
   dc_AskRange( "Range of labels", {| a | dc_ForArea( a, {| nC, nR | dc_Sheet():setAlign( nC, nR, cAlign ) } ) } )
RETURN

STATIC PROCEDURE dc_EraseRange()
   dc_AskRange( "Range to erase", {| a | DC_EraseArea( dc_Sheet(), a ) } )
RETURN

STATIC PROCEDURE dc_EraseSheet()
   LOCAL oS := dc_Sheet()
   oS:hCells := { => }
   oS:rebuildDeps()
   oS:lModified := .T.
RETURN

/* insert / delete as many rows or columns as are selected */
STATIC PROCEDURE dc_InsDel( lRows, lInsert )
   LOCAL a := s_oGrid:selArea(), n
   n := IIF( lRows, a[ 4 ] - a[ 2 ] + 1, a[ 3 ] - a[ 1 ] + 1 )
   dc_Sheet():shiftCells( lRows, IIF( lRows, a[ 2 ], a[ 1 ] ), IIF( lInsert, n, -n ) )
   s_oGrid:clearSel()
   s_oGrid:message( IIF( lInsert, "Inserted ", "Deleted " ) + hb_ntos( n ) + IIF( lRows, " row(s)", " column(s)" ) )
RETURN

STATIC PROCEDURE dc_ColWidth()
   LOCAL a := s_oGrid:selArea()
   dc_AskNum( "Column width (1-240)", dc_Sheet():colWidth( s_oGrid:nCol ), 1, 240, ;
      {| n | dc_SetWidths( a, Int( n ) ) } )
RETURN

STATIC PROCEDURE dc_SetWidths( a, n )
   dc_ForArea( { a[ 1 ], 1, a[ 3 ], 1 }, {| nC | dc_Sheet():setColWidth( nC, n ) } )
RETURN

STATIC PROCEDURE dc_ColReset()
   LOCAL a := s_oGrid:selArea(), n
   FOR n := a[ 1 ] TO a[ 3 ]
      dc_Sheet():setColWidth( n, NIL )
   NEXT
RETURN

/* AutoFit: the widest value of each selected column (as shown) */
STATIC PROCEDURE dc_ColFit()
   LOCAL a := s_oGrid:selArea(), oS := dc_Sheet(), nC, k, nW, aC, c
   FOR nC := a[ 1 ] TO a[ 3 ]
      nW := 3
      FOR EACH k IN hb_HKeys( oS:hCells )
         IF DC_KeyCol( k ) == nC
            aC := oS:hCells[ k ]
            c := s_oGrid:cellText( nC, DC_KeyRow( k ), 200 )
            nW := Max( nW, Len( AllTrim( c ) ) + IIF( HB_ISNUMERIC( aC[ CELL_VALUE ] ), 2, 1 ) )
         ENDIF
      NEXT
      oS:setColWidth( nC, Min( nW, 60 ) )
   NEXT
RETURN

STATIC PROCEDURE dc_GlobalWidth()
   dc_AskNum( "Global column width (1-240)", dc_Sheet():nDefWidth, 1, 240, {| n | dc_Sheet():nDefWidth := Int( n ) } )
RETURN

STATIC PROCEDURE dc_Copy()
   dc_AskRange( "Copy what", {| aFrom | s_oGrid:clearSel(), dc_CopyTo( aFrom ) } )
RETURN

/* a one-cell source fills the whole target range; else the target is its top left cell */
STATIC PROCEDURE dc_CopyTo( aFrom )
   dc_AskRange( "To where", {| aTo | ;
      dc_Sheet():copyArea( aFrom, IIF( aFrom[ 1 ] == aFrom[ 3 ] .AND. aFrom[ 2 ] == aFrom[ 4 ], aTo, aTo[ 1 ] ), aTo[ 2 ] ), ;
      s_oGrid:message( "Copied " + DC_AreaText( aFrom ) + " to " + DC_AreaText( aTo ) ) } )
RETURN

STATIC PROCEDURE dc_Paste()
   IF HB_ISNIL( s_aClip )
      s_oGrid:message( "Nothing copied: Edit - Copy first" )
      RETURN
   ENDIF
   IF s_aClip[ 1 ] == dc_Sheet()
      dc_Sheet():copyArea( s_aClip[ 2 ], s_oGrid:nCol, s_oGrid:nRow )
   ELSE
      dc_PasteFrom( s_aClip[ 1 ], s_aClip[ 2 ], s_oGrid:nCol, s_oGrid:nRow )
   ENDIF
RETURN

/* from another sheet: the values and formulas (moved), the same way */
STATIC PROCEDURE dc_PasteFrom( oFrom, aA, nToC, nToR )
   LOCAL nC, nR, aC, oS := dc_Sheet()
   FOR nR := aA[ 2 ] TO aA[ 4 ]
      FOR nC := aA[ 1 ] TO aA[ 3 ]
         aC := oFrom:cell( nC, nR )
         IF HB_ISNIL( aC )
            oS:clear( nToC + nC - aA[ 1 ], nToR + nR - aA[ 2 ] )
         ELSEIF ! HB_ISNIL( aC[ CELL_AST ] )
            oS:setInput( nToC + nC - aA[ 1 ], nToR + nR - aA[ 2 ], "=" + DC_AstText( DC_AstShift( aC[ CELL_AST ], nToC - aA[ 1 ], nToR - aA[ 2 ] ) ) )
         ELSE
            oS:setCell( nToC + nC - aA[ 1 ], nToR + nR - aA[ 2 ], aC[ CELL_VALUE ] )
         ENDIF
      NEXT
   NEXT
   oS:recalcAll()
RETURN

STATIC PROCEDURE dc_Move()
   dc_AskRange( "Move what", {| aFrom | s_oGrid:clearSel(), dc_MoveTo( aFrom ) } )
RETURN

STATIC PROCEDURE dc_MoveTo( aFrom )
   dc_AskRange( "To where", {| aTo | ;
      dc_Sheet():moveArea( aFrom, aTo[ 1 ], aTo[ 2 ] ), ;
      s_oGrid:message( "Moved " + DC_AreaText( aFrom ) + " to " + DC_CellName( aTo[ 1 ], aTo[ 2 ] ) ) } )
RETURN

STATIC PROCEDURE dc_Fill()
   dc_AskRange( "Range to fill", {| a | dc_FillStart( a ) } )
RETURN

STATIC PROCEDURE dc_FillStart( a )
   dc_AskNum( "Start", 1, -999999999999, 999999999999, {| nStart | dc_FillStep( a, nStart ) } )
RETURN

STATIC PROCEDURE dc_FillStep( a, nStart )
   dc_AskNum( "Step", 1, -999999999999, 999999999999, {| nStep | dc_Sheet():fillArea( a, nStart, nStep ) } )
RETURN

STATIC PROCEDURE dc_Sort( lDesc )
   dc_AskRange( "Rows to sort (the data range)", {| a | dc_SortKey( a, lDesc ) } )
RETURN

STATIC PROCEDURE dc_SortKey( a, lDesc )
   s_oGrid:prompt( "Key column (a letter, or a cell in it): ", DC_ColName( a[ 1 ] ), ;
      {| c, k | k := DC_ColNum( StrTran( Upper( AllTrim( c ) ), "$", "" ) ), ;
                IIF( k == 0, k := IIF( HB_ISNIL( DC_ParseRef( c ) ), 0, DC_ParseRef( c )[ 1 ] ), NIL ), ;
                IIF( k < a[ 1 ] .OR. k > a[ 3 ], s_oGrid:message( "The key column must be in the range" ), ;
                     dc_Sheet():sortArea( a, k, lDesc ) ) } )
RETURN

STATIC PROCEDURE dc_NameCreate()
   s_oGrid:prompt( "Name: ", "", {| cName | ;
      IIF( Empty( cName ) .OR. ! IsAlpha( Left( cName, 1 ) ) .OR. ! HB_ISNIL( DC_ParseRef( cName ) ), ;
           s_oGrid:message( "A name starts with a letter and is not a cell address" ), ;
           dc_NameRange( cName ) ) } )
RETURN

STATIC PROCEDURE dc_NameRange( cName )
   dc_AskRange( "Range of " + Upper( AllTrim( cName ) ), {| a | dc_Sheet():setName( cName, a ) } )
RETURN

STATIC PROCEDURE dc_NameDelete()
   LOCAL aNames := hb_HKeys( dc_Sheet():hNames )
   s_oGrid:prompt( "Delete name (" + IIF( Empty( aNames ), "none", dc_Join( aNames ) ) + "): ", "", ;
      {| c | dc_Sheet():setName( c, NIL ) } )
RETURN

STATIC FUNCTION dc_Join( a )
   LOCAL c := ""
   AEval( a, {| x, i | c += IIF( i > 1, " ", "" ) + x } )
RETURN c

STATIC PROCEDURE dc_SheetInsert()
   LOCAL n := Len( s_aSheets ) + 1, cName
   DO WHILE AScan( s_aSheets, {| o | Upper( o:cName ) == "SHEET" + hb_ntos( n ) } ) > 0
      n++
   ENDDO
   cName := "Sheet" + hb_ntos( n )
   hb_AIns( s_aSheets, s_nSheet + 1, DCSheet():new( cName ), .T. )
   hb_AIns( s_aPos, s_nSheet + 1, { 1, 1, 1, 1 }, .T. )
   s_aSheets[ s_nSheet + 1 ]:lModified := .T.
   dc_GoSheet( s_nSheet + 1 )
RETURN

STATIC PROCEDURE dc_SheetDelete()
   IF Len( s_aSheets ) == 1
      s_oGrid:message( "The only sheet cannot be deleted" )
      RETURN
   ENDIF
   IF ! HTMessageBox():question( "DigiCalc", "Delete the sheet " + dc_Sheet():cName + "?" )
      RETURN
   ENDIF
   hb_ADel( s_aSheets, s_nSheet, .T. )
   hb_ADel( s_aPos, s_nSheet, .T. )
   s_aSheets[ 1 ]:lModified := .T.
   s_nSheet := Min( s_nSheet, Len( s_aSheets ) )
   s_oGrid:oSheet := dc_Sheet()
   s_oGrid:nCol := s_aPos[ s_nSheet ][ 1 ]
   s_oGrid:nRow := s_aPos[ s_nSheet ][ 2 ]
   s_oGrid:nTop := s_aPos[ s_nSheet ][ 3 ]
   s_oGrid:nLeft := s_aPos[ s_nSheet ][ 4 ]
RETURN

STATIC PROCEDURE dc_SheetRename()
   s_oGrid:prompt( "Sheet name: ", dc_Sheet():cName, {| c | ;
      IIF( Empty( c ), NIL, ( dc_Sheet():cName := Left( AllTrim( c ), 31 ), dc_Sheet():lModified := .T. ) ) } )
RETURN

/*------------------------------------------------------------ files */

STATIC PROCEDURE dc_New()
   IF dc_AskSave()
      dc_SetBook( { DCSheet():new( "Sheet1" ) }, "" )
   ENDIF
RETURN

STATIC PROCEDURE dc_CmdOpen( cMasks )
   LOCAL cFile
   IF ! dc_AskSave()
      RETURN
   ENDIF
   cFile := DC_FileDialog( "Open", s_cFile, hb_defaultValue( cMasks, "*.xlsx;*.csv;*.txt" ) )
   IF ! Empty( cFile )
      dc_Open( cFile )
   ENDIF
RETURN

STATIC PROCEDURE dc_Open( cFile )
   LOCAL cErr := "", aBook, oS
   IF ! hb_FileExists( cFile )
      HTMessageBox():error( "DigiCalc", cFile + Chr( 10 ) + "was not found." )
      RETURN
   ENDIF
   IF Lower( hb_FNameExt( cFile ) ) == ".xlsx"
      aBook := DC_LoadXlsx( cFile, @cErr )
   ELSE
      oS := DC_LoadCsv( cFile, @cErr )
      aBook := IIF( HB_ISNIL( oS ), NIL, { oS } )
      cFile := hb_FNameExtSet( cFile, ".xlsx" )         /* Save writes a workbook next to it */
   ENDIF
   IF HB_ISNIL( aBook )
      HTMessageBox():error( "DigiCalc", cErr )
      RETURN
   ENDIF
   dc_SetBook( aBook, IIF( hb_FileExists( cFile ) .OR. Lower( hb_FNameExt( cFile ) ) == ".xlsx", cFile, "" ) )
RETURN

/* lAs: ask for the name (always when there is none) */
STATIC FUNCTION dc_Save( lAs )
   LOCAL cFile := s_cFile, cErr := ""
   IF lAs .OR. Empty( cFile )
      cFile := DC_FileDialog( "Save as (.xlsx)", IIF( Empty( cFile ), hb_cwd() + "book1.xlsx", cFile ), "*.xlsx" )
      IF Empty( cFile )
         RETURN .F.
      ENDIF
      IF Empty( hb_FNameExt( cFile ) )
         cFile += ".xlsx"
      ENDIF
      IF hb_FileExists( cFile ) .AND. ! ( cFile == s_cFile ) .AND. ;
         ! HTMessageBox():question( "DigiCalc", hb_FNameNameExt( cFile ) + " exists. Replace it?" )
         RETURN .F.
      ENDIF
   ENDIF
   IF ! DC_SaveXlsx( cFile, s_aSheets, @cErr )
      HTMessageBox():error( "DigiCalc", cErr )
      RETURN .F.
   ENDIF
   s_cFile := cFile
   s_oGrid:message( "Saved: " + cFile )
RETURN .T.

STATIC PROCEDURE dc_Export()
   LOCAL cFile, cErr := ""
   cFile := DC_FileDialog( "Export this sheet (.csv)", hb_FNameExtSet( IIF( Empty( s_cFile ), hb_cwd() + dc_Sheet():cName, s_cFile ), ".csv" ), "*.csv" )
   IF Empty( cFile )
      RETURN
   ENDIF
   IF ! DC_SaveCsv( cFile, dc_Sheet(), ";", @cErr )
      HTMessageBox():error( "DigiCalc", cErr )
   ELSE
      dc_Sheet():lModified := .T.                    /* .csv keeps only the values: the workbook is still unsaved */
      s_oGrid:message( "Exported: " + cFile )
   ENDIF
RETURN

STATIC PROCEDURE dc_Exit()
   IF dc_AskSave()
      s_oApp:quit()
   ENDIF
RETURN

/* .T. when it may go on: nothing changed, saved, or not wanted */
STATIC FUNCTION dc_AskSave()
   LOCAL n
   IF ! dc_Modified()
      RETURN .T.
   ENDIF
   n := HTMessageBox():choose( "DigiCalc", "The workbook has been changed. Save it?", { "Save", "Don't save", "Cancel" } )
   DO CASE
   CASE n == 1
      RETURN dc_Save( .F. )
   CASE n == 2
      RETURN .T.
   ENDCASE
RETURN .F.

STATIC PROCEDURE dc_About()
   HTMessageBox():information( "DigiCalc", "DigiCalc " + DC_VERSION + " - spreadsheet in character mode" + Chr( 10 ) + ;
      "Formulas: Excel (=SUM(A1:A5)) and 1-2-3 (@SUM(A1..A5))" + Chr( 10 ) + ;
      "Files: .xlsx (github.com/digikv/xlsx), .csv" + Chr( 10 ) + ;
      "Copyright (c) 2026 Srđan Dragojlović - GPL 3 or commercial license" )
RETURN

STATIC PROCEDURE dc_Help()
   HTMessageBox():information( "DigiCalc - keys", ;
      "Type a number, text or formula, then Enter or an arrow." + Chr( 10 ) + ;
      "Formula: =A1+B1  =SUM(A1:A5)   or 1-2-3: +A1+B1  @SUM(A1..A5)" + Chr( 10 ) + ;
      "Text: 'left  " + '"' + "right  ^center (1-2-3 prefixes)" + Chr( 10 ) + ;
      "In a formula the arrows point at a cell; . or : anchors a range." + Chr( 10 ) + ;
      "/ the 1-2-3 menu   F10 the menu bar   Alt-X exit" + Chr( 10 ) + ;
      "F2 edit   F5 go to   F9 recalculate   Del erase" + Chr( 10 ) + ;
      "End + arrow: edge of the data   End Home: last cell   Home: A1" + Chr( 10 ) + ;
      "Shift + arrows: select   Tab / Shift-Tab: a screen right / left" + Chr( 10 ) + ;
      "Ctrl-PgDn / Ctrl-PgUp: next / previous sheet" )
RETURN
