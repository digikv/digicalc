/** @file dcfile.prg
 * DigiCalc - files: .xlsx through the XLSX class (github.com/digikv/xlsx),
 * and .csv.
 *
 *   DC_LoadXlsx( cFile, @cErr )             -> { oSheet, ... } or NIL
 *   DC_SaveXlsx( cFile, aSheets, @cErr )    -> lOk
 *   DC_LoadCsv( cFile, @cErr )              -> oSheet or NIL
 *   DC_SaveCsv( cFile, oSheet, cSep, @cErr ) -> lOk
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "digicalc.ch"

#define DC_DATE0   0d18991230

REQUEST HB_CODEPAGE_HRWIN        /* Windows-1250 (Latin 2): an old .csv */

/*------------------------------------------------------------ .xlsx */

#ifdef DC_NO_XLSX
/* built without ../xlsx/XLSXClass.prg */
FUNCTION DC_LoadXlsx( cFile, cErr )
   HB_SYMBOL_UNUSED( cFile )
   cErr := "DigiCalc was built without the XLSX class (github.com/digikv/xlsx)"
RETURN NIL

FUNCTION DC_SaveXlsx( cFile, aSheets, cErr )
   HB_SYMBOL_UNUSED( cFile )
   HB_SYMBOL_UNUSED( aSheets )
   cErr := "DigiCalc was built without the XLSX class (github.com/digikv/xlsx)"
RETURN .F.
#else
FUNCTION DC_LoadXlsx( cFile, cErr )

   LOCAL oBook := WorkBook():New(), aSheets := {}, cName, oWs, oS, nR, nC, x, cF, nSt, h, a, nW, i
   LOCAL aCalc, hFmt := { => }

   cErr := ""
   oBook:lReadDates := .F.                  /* dates as serial numbers, with their format */
   IF ! oBook:Read( cFile )
      cErr := oBook:cError
      RETURN NIL
   ENDIF
   FOR EACH cName IN oBook:SheetNames()
      oWs := oBook:GetSheet( cName )
      oS := DCSheet():new( cName )
      oS:lAutoCalc := .F.
      aCalc := {}
      FOR nR := 1 TO oWs:LastRow()
         FOR nC := 1 TO Min( oWs:LastCol(), DC_MAXCOL )
            x := oWs:ReadCell( { nR, nC } )
            cF := oWs:CellFormula( { nR, nC } )
            nSt := oWs:CellStyle( { nR, nC } )
            IF HB_ISNIL( x ) .AND. HB_ISNIL( cF ) .AND. nSt == 0
               LOOP
            ENDIF
            IF HB_ISSTRING( x ) .AND. x == Chr( 0 )     /* the rest of a merged cell */
               LOOP
            ENDIF
            IF HB_ISDATETIME( x )
               x := ( hb_TToD( x ) - DC_DATE0 ) + ( x - hb_DToT( hb_TToD( x ) ) )
            ELSEIF HB_ISDATE( x )
               x := x - DC_DATE0
            ENDIF
            IF ! HB_ISNIL( cF )
               cF := SubStr( cF, 2 )
               IF HB_ISSTRING( x ) .AND. Left( x, 1 ) == "=" .AND. SubStr( x, 2 ) == cF
                  x := NIL                              /* no value in the file: computed below */
                  AAdd( aCalc, DC_Key( nC, nR ) )
               ENDIF
            ENDIF
            IF ! HB_ISNIL( x ) .OR. ! HB_ISNIL( cF )
               oS:setCell( nC, nR, x, cF )
            ENDIF
            IF nSt > 0
               IF ! hb_HHasKey( hFmt, nSt )
                  h := oBook:StyleInfo( nSt )
                  hFmt[ nSt ] := { DC_FmtFromXlsx( h[ "numformat" ] ), ;
                                   IIF( h[ "halign" ] == "left", "L", IIF( h[ "halign" ] == "right", "R", IIF( h[ "halign" ] == "center", "C", "" ) ) ) }
               ENDIF
               IF ! Empty( hFmt[ nSt ][ 1 ] )
                  oS:setFormat( nC, nR, hFmt[ nSt ][ 1 ] )
               ENDIF
               IF ! Empty( hFmt[ nSt ][ 2 ] ) .AND. ! HB_ISNIL( oS:cell( nC, nR ) )
                  oS:setAlign( nC, nR, hFmt[ nSt ][ 2 ] )
               ENDIF
            ENDIF
         NEXT
      NEXT
      FOR EACH a IN oWs:ColumnWidths()
         nW := Max( 1, Round( a[ 3 ], 0 ) )
         FOR i := Max( 1, a[ 1 ] ) TO Min( a[ 2 ], 1000 )
            oS:setColWidth( i, nW )
         NEXT
      NEXT
      oS:lAutoCalc := .T.
      IF ! Empty( aCalc )
         oS:recalc( aCalc )
      ENDIF
      oS:lModified := .F.
      AAdd( aSheets, oS )
   NEXT
   IF Empty( aSheets )
      cErr := hb_FNameNameExt( cFile ) + ": the workbook has no worksheets"
      RETURN NIL
   ENDIF

RETURN aSheets

FUNCTION DC_SaveXlsx( cFile, aSheets, cErr )

   LOCAL oBook, oWs, oS, k, aC, nC, nR, nSt, hStyle := { => }, cKey, x, aKeys

   cErr := ""
   oBook := WorkBook():New( cFile )
   oBook:Application := "DigiCalc"
   FOR EACH oS IN aSheets
      oWs := oBook:WorkSheet( oS:cName )
      aKeys := ASort( hb_HKeys( oS:hCells ) )
      FOR EACH k IN aKeys
         aC := oS:hCells[ k ]
         nC := DC_KeyCol( k )
         nR := DC_KeyRow( k )
         nSt := NIL
         cKey := aC[ CELL_FMT ] + "|" + aC[ CELL_ALIGN ]
         IF !( cKey == "|" ) .AND. !( cKey == "G|" )
            IF ! hb_HHasKey( hStyle, cKey )
               hStyle[ cKey ] := oBook:NewStyle( , , , , ;
                  IIF( aC[ CELL_ALIGN ] == "L", 1, IIF( aC[ CELL_ALIGN ] == "C", 2, IIF( aC[ CELL_ALIGN ] == "R", 3, NIL ) ) ), ;
                  IIF( Empty( DC_FmtToXlsx( aC[ CELL_FMT ] ) ), NIL, oBook:NewFormat( DC_FmtToXlsx( aC[ CELL_FMT ] ) ) ) )
            ENDIF
            nSt := hStyle[ cKey ]
         ENDIF
         x := aC[ CELL_VALUE ]
         DO CASE
         CASE ! HB_ISNIL( aC[ CELL_AST ] )
            oWs:Formula( { nR, nC }, DC_AstText( aC[ CELL_AST ] ), nSt )
         CASE ! Empty( aC[ CELL_RAW ] )
            oWs:Formula( { nR, nC }, aC[ CELL_RAW ], nSt )
         CASE HB_ISSTRING( x ) .OR. HB_ISNUMERIC( x ) .OR. HB_ISLOGICAL( x )
            oWs:Cell( { nR, nC }, x, nSt )
         CASE DC_IsErr( x )
            oWs:Cell( { nR, nC }, x:cCode, nSt )
         OTHERWISE
            IF ! HB_ISNIL( nSt )
               oWs:CellStyle( { nR, nC }, nSt )
            ENDIF
         ENDCASE
      NEXT
      FOR EACH k IN ASort( hb_HKeys( oS:hColW ) )
         oWs:ColumnsWidth( k, k, oS:hColW[ k ] )
      NEXT
   NEXT
   IF ! oBook:Save()
      cErr := oBook:cError
      RETURN .F.
   ENDIF
   AEval( aSheets, {| o | o:lModified := .F. } )

RETURN .T.
#endif

/* an Excel number format -> a DigiCalc one ("" = General) */
FUNCTION DC_FmtFromXlsx( cCode )

   LOCAL c := Lower( StrTran( hb_defaultValue( cCode, "" ), "\", "" ) ), nDec := 0, n

   IF Empty( c ) .OR. c == "general" .OR. c == "@"
      RETURN ""
   ENDIF
   IF c == ";;;"
      RETURN "H"
   ENDIF
   /* the decimals: the 0s after the first point */
   n := At( ".", c )
   IF n > 0
      DO WHILE n + nDec + 1 <= Len( c ) .AND. SubStr( c, n + nDec + 1, 1 ) $ "0#"
         nDec++
      ENDDO
   ENDIF
   DO CASE
   CASE "%" $ c
      RETURN "P" + hb_ntos( nDec )
   CASE "e+" $ c
      RETURN "S" + hb_ntos( nDec )
   CASE ( "h" $ c .AND. ":" $ c ) .AND. !( "d" $ c )
      RETURN "D5"
   CASE "mmm" $ c
      RETURN "D2"
   CASE "y" $ c .AND. "d" $ c
      RETURN "D1"
   CASE "y" $ c .AND. "m" $ c
      RETURN "D4"
   CASE "d" $ c .AND. "m" $ c
      RETURN "D3"
   CASE ( "$" $ c .OR. "€" $ c .OR. "din" $ c .OR. "rsd" $ c ) .AND. "0" $ c
      RETURN "C" + hb_ntos( nDec )
   CASE "#,##0" $ c .OR. "#,#" $ c
      RETURN "," + hb_ntos( nDec )
   CASE "0" $ c
      RETURN "F" + hb_ntos( nDec )
   ENDCASE

RETURN ""

/* a DigiCalc format -> an Excel number format ("" = General) */
FUNCTION DC_FmtToXlsx( cFmt )

   LOCAL cK := Left( cFmt, 1 ), nDec := Val( SubStr( cFmt, 2 ) ), cDec

   cDec := IIF( nDec > 0, "." + Replicate( "0", nDec ), "" )
   SWITCH cK
   CASE "F" ; RETURN "0" + cDec
   CASE "," ; RETURN "#,##0" + cDec
   CASE "C" ; RETURN IIF( Empty( DC_Currency() ), "", '"' + DC_Currency() + '"' ) + "#,##0" + cDec
   CASE "P" ; RETURN "0" + cDec + "%"
   CASE "S" ; RETURN "0" + cDec + "E+00"
   CASE "H" ; RETURN ";;;"
   CASE "D"
      SWITCH cFmt
      CASE "D2" ; RETURN "dd\-mmm\-yy"
      CASE "D3" ; RETURN "dd\.mm\."
      CASE "D4" ; RETURN "mm\.yyyy"
      CASE "D5" ; RETURN "hh:mm"
      ENDSWITCH
      RETURN "dd\.mm\.yyyy"
   ENDSWITCH

RETURN ""

/*------------------------------------------------------------ .csv */

/* separator found in the first line: ; , or Tab; "1,5" is a number with ; */
FUNCTION DC_LoadCsv( cFile, cErr )

   LOCAL cText, aRows, aRow, oS, nR, nC, c, cSep

   cErr := ""
   IF ! hb_FileExists( cFile )
      cErr := cFile + ": not found"
      RETURN NIL
   ENDIF
   cText := hb_MemoRead( cFile )
   IF Left( cText, 3 ) == e"\xEF\xBB\xBF"
      cText := SubStr( cText, 4 )
   ELSEIF ! hb_StrIsUTF8( cText )
      cText := hb_Translate( cText, "HRWIN", "UTF8EX" )   /* an old Windows file */
   ENDIF
   cSep := CsvSep( cText )
   aRows := CsvSplit( cText, cSep )
   oS := DCSheet():new( hb_FNameName( cFile ) )
   oS:lAutoCalc := .F.
   FOR EACH aRow IN aRows
      nR := aRow:__enumIndex()
      FOR EACH c IN aRow
         nC := c:__enumIndex()
         IF Empty( c ) .OR. nC > DC_MAXCOL .OR. nR > DC_MAXROW
            LOOP
         ENDIF
         IF cSep == ";" .AND. DC_IsNumText( StrTran( c, ",", "." ) ) .AND. !( "." $ c )
            c := StrTran( c, ",", "." )
         ENDIF
         DO CASE
         CASE DC_IsNumText( c )
            oS:setCell( nC, nR, DC_Val( c ) )
         CASE Left( c, 1 ) == "="
            oS:setCell( nC, nR, NIL, SubStr( c, 2 ) )
         OTHERWISE
            oS:setCell( nC, nR, c )
         ENDCASE
      NEXT
   NEXT
   oS:lAutoCalc := .T.
   oS:recalcAll()
   oS:lModified := .F.

RETURN oS

STATIC FUNCTION CsvSep( cText )
   LOCAL n := At( Chr( 10 ), cText ), cLine, nSemi, nComma, nTab
   cLine := IIF( n > 0, Left( cText, n ), cText )
   nSemi := hb_TokenCount( cLine, ";" ) - 1
   nComma := hb_TokenCount( cLine, "," ) - 1
   nTab := hb_TokenCount( cLine, Chr( 9 ) ) - 1
   DO CASE
   CASE nTab > nSemi .AND. nTab > nComma
      RETURN Chr( 9 )
   CASE nSemi >= nComma .AND. nSemi > 0
      RETURN ";"
   ENDCASE
RETURN ","

/* rows of fields; "..." may hold the separator, "" and line ends */
STATIC FUNCTION CsvSplit( cText, cSep )
   LOCAL aRows := {}, aRow := {}, cField := "", i := 1, n := Len( cText ), c, lQ := .F.
   DO WHILE i <= n
      c := SubStr( cText, i, 1 )
      IF lQ
         IF c == '"'
            IF SubStr( cText, i + 1, 1 ) == '"'
               cField += '"'
               i++
            ELSE
               lQ := .F.
            ENDIF
         ELSE
            cField += c
         ENDIF
      ELSE
         DO CASE
         CASE c == '"'
            lQ := .T.
         CASE c == cSep
            AAdd( aRow, cField )
            cField := ""
         CASE c == Chr( 10 )
            AAdd( aRow, cField )
            AAdd( aRows, aRow )
            aRow := {}
            cField := ""
         CASE c == Chr( 13 )
         OTHERWISE
            cField += c
         ENDCASE
      ENDIF
      i++
   ENDDO
   IF ! Empty( cField ) .OR. ! Empty( aRow )
      AAdd( aRow, cField )
      AAdd( aRows, aRow )
   ENDIF
RETURN aRows

/* the values (not the formulas), numbers with a point; UTF-8 with a BOM (Excel) */
FUNCTION DC_SaveCsv( cFile, oS, cSep, cErr )

   LOCAL nLastR := oS:lastRow(), nLastC := oS:lastCol(), nR, nC, x, c, cOut := e"\xEF\xBB\xBF", aLine

   cErr := ""
   cSep := hb_defaultValue( cSep, "," )
   FOR nR := 1 TO nLastR
      aLine := {}
      FOR nC := 1 TO nLastC
         x := oS:value( nC, nR )
         c := IIF( HB_ISNUMERIC( x ), DC_NumText( x ), DC_ToText( x ) )
         IF cSep $ c .OR. '"' $ c .OR. Chr( 10 ) $ c .OR. Chr( 13 ) $ c
            c := '"' + StrTran( c, '"', '""' ) + '"'
         ENDIF
         AAdd( aLine, c )
      NEXT
      DO WHILE Len( aLine ) > 0 .AND. Empty( ATail( aLine ) )
         ASize( aLine, Len( aLine ) - 1 )
      ENDDO
      c := ""
      AEval( aLine, {| f, i | c += IIF( i > 1, cSep, "" ) + f } )
      cOut += c + Chr( 13 ) + Chr( 10 )
   NEXT
   IF ! hb_MemoWrit( cFile, cOut )
      cErr := cFile + ": cannot be written"
      RETURN .F.
   ENDIF
   oS:lModified := .F.

RETURN .T.
