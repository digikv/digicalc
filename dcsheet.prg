/** @file dcsheet.prg
 * DigiCalc - a worksheet: the cells, what each formula depends on, and the
 * recalculation of just the cells that depend on a change.
 *
 *   oS := DCSheet():new( "Sheet1" )
 *   oS:setInput( nCol, nRow, "=SUM(A1:A5)", @cErr )     also "@SUM(A1..A5)", "'Text", "12"
 *   oS:value( nCol, nRow )  oS:input( nCol, nRow, lLotus )
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "hbclass.ch"
#include "digicalc.ch"

CREATE CLASS DCSheet

   VAR cName      INIT "Sheet1"
   VAR hCells                       /* key -> cell (digicalc.ch CELL_*) */
   VAR hDeps                        /* key -> { key of a formula that uses it => .T. } */
   VAR aAreaDeps                    /* { { nKey, nC1, nR1, nC2, nR2 } }  formulas that use an area */
   VAR hNames                       /* name -> { "R"|"A", ... } (named ranges) */
   VAR hColW                        /* nCol -> width */
   VAR nDefWidth  INIT 9            /* as 1-2-3 */
   VAR cDefFmt    INIT "G"          /* General */
   VAR lAutoCalc  INIT .T.
   VAR lModified  INIT .F.

   VAR hDirty     PROTECTED         /* during a recalculation: still to compute */
   VAR hBusy      PROTECTED         /* being computed (a circle when met again) */

   METHOD new( cName )
   METHOD cell( nCol, nRow )
   METHOD value( nCol, nRow )
   METHOD input( nCol, nRow, lLotus )
   METHOD setInput( nCol, nRow, cText, cErr )
   METHOD setCell( nCol, nRow, xValue, cFormula )
   METHOD clear( nCol, nRow )
   METHOD setFormat( nCol, nRow, cFmt )
   METHOD setAlign( nCol, nRow, cAlign )
   METHOD colWidth( nCol )
   METHOD setColWidth( nCol, nWidth )
   METHOD lastRow()
   METHOD lastCol()

   METHOD recalc( aKeys )
   METHOD recalcAll()
   METHOD evalCell( nCol, nRow )
   METHOD evalName( cName )
   METHOD areaValues( a1, a2 )
   METHOD areaMatrix( a1, a2 )

   METHOD unlink( nKey )
   METHOD link( nKey, xAst )
   METHOD dependents( nKey )
   METHOD compute( nKey )

ENDCLASS

METHOD new( cName ) CLASS DCSheet
   ::cName := hb_defaultValue( cName, "Sheet1" )
   ::hCells := { => }
   ::hDeps := { => }
   ::aAreaDeps := {}
   ::hNames := { => }
   ::hColW := { => }
   hb_HCaseMatch( ::hNames, .F. )
RETURN Self

METHOD cell( nCol, nRow ) CLASS DCSheet
RETURN hb_HGetDef( ::hCells, DC_Key( nCol, nRow ), NIL )

METHOD value( nCol, nRow ) CLASS DCSheet
   LOCAL aC := hb_HGetDef( ::hCells, DC_Key( nCol, nRow ), NIL )
RETURN IIF( HB_ISNIL( aC ), NIL, aC[ CELL_VALUE ] )

/* the entry line: a formula in Excel (=SUM(A1:A5)) or 1-2-3 (@SUM(A1..A5)) form */
METHOD input( nCol, nRow, lLotus ) CLASS DCSheet
   LOCAL aC := ::cell( nCol, nRow ), c
   IF HB_ISNIL( aC )
      RETURN ""
   ENDIF
   IF ! HB_ISNIL( aC[ CELL_AST ] )
      IF hb_defaultValue( lLotus, .F. )
         c := DC_AstText( aC[ CELL_AST ], .T. )
         /* 1-2-3: a formula starts with + when it would read as a label */
         RETURN IIF( Left( c, 1 ) $ "+-(@.0123456789$", c, "+" + c )
      ENDIF
      RETURN "=" + DC_AstText( aC[ CELL_AST ], .F. )
   ENDIF
RETURN aC[ CELL_INPUT ]

/* what was typed into a cell:
     =...                   a formula (Excel)
     + - ( @ . digit        a number, or a formula (1-2-3)
     ' " ^                  a label: left / right / centered (1-2-3 prefix)
     anything else          a label
   "" empties the cell. .F. (and cErr) when the formula does not parse. */
METHOD setInput( nCol, nRow, cText, cErr ) CLASS DCSheet

   LOCAL nKey := DC_Key( nCol, nRow ), aC, cFirst, xAst := NIL, xVal, cAlign := "", cFmt := ""

   cErr := ""
   aC := hb_HGetDef( ::hCells, nKey, NIL )
   IF ! HB_ISNIL( aC )
      cFmt := aC[ CELL_FMT ]
      cAlign := aC[ CELL_ALIGN ]
   ENDIF
   IF Empty( cText )
      ::clear( nCol, nRow )
      RETURN .T.
   ENDIF

   cFirst := Left( cText, 1 )
   DO CASE
   CASE cFirst == "="
      xAst := DC_Parse( SubStr( cText, 2 ), @cErr )
      IF HB_ISNIL( xAst )
         RETURN .F.
      ENDIF
   CASE DC_IsNumText( cText )
      xVal := DC_Val( cText )
   CASE cFirst $ "+-(@.0123456789$#"
      xAst := DC_Parse( cText, @cErr )
      IF HB_ISNIL( xAst )
         RETURN .F.
      ENDIF
   CASE cFirst $ "'" + '"^'
      xVal := SubStr( cText, 2 )
      cAlign := IIF( cFirst == "'", "L", IIF( cFirst == '"', "R", "C" ) )
   OTHERWISE
      xVal := cText
   ENDCASE

   ::unlink( nKey )
   aC := Array( CELL_LEN )
   aC[ CELL_FMT ] := cFmt
   aC[ CELL_ALIGN ] := cAlign
   IF ! HB_ISNIL( xAst )
      aC[ CELL_AST ] := xAst
      aC[ CELL_INPUT ] := "=" + DC_AstText( xAst )
      ::hCells[ nKey ] := aC
      ::link( nKey, xAst )
   ELSE
      aC[ CELL_VALUE ] := xVal
      aC[ CELL_INPUT ] := IIF( HB_ISSTRING( xVal ), cText, DC_NumText( xVal ) )
      ::hCells[ nKey ] := aC
   ENDIF
   ::lModified := .T.
   IF ::lAutoCalc
      ::recalc( { nKey } )
   ENDIF

RETURN .T.

/* a value (and a formula in Excel form, without =) from a file */
METHOD setCell( nCol, nRow, xValue, cFormula ) CLASS DCSheet

   LOCAL nKey := DC_Key( nCol, nRow ), aC := Array( CELL_LEN ), xAst := NIL, cErr

   ::unlink( nKey )
   aC[ CELL_FMT ] := ""
   aC[ CELL_ALIGN ] := ""
   aC[ CELL_VALUE ] := xValue
   IF ! Empty( cFormula )
      xAst := DC_Parse( cFormula, @cErr )
   ENDIF
   IF ! HB_ISNIL( xAst )
      aC[ CELL_AST ] := xAst
      aC[ CELL_INPUT ] := "=" + DC_AstText( xAst )
      ::link( nKey, xAst )
   ELSE
      aC[ CELL_INPUT ] := IIF( HB_ISSTRING( xValue ), xValue, DC_ToText( xValue ) )
   ENDIF
   ::hCells[ nKey ] := aC

RETURN NIL

METHOD clear( nCol, nRow ) CLASS DCSheet
   LOCAL nKey := DC_Key( nCol, nRow )
   IF hb_HHasKey( ::hCells, nKey )
      ::unlink( nKey )
      hb_HDel( ::hCells, nKey )
      ::lModified := .T.
      IF ::lAutoCalc
         ::recalc( { nKey } )
      ENDIF
   ENDIF
RETURN NIL

METHOD setFormat( nCol, nRow, cFmt ) CLASS DCSheet
   LOCAL aC := ::cell( nCol, nRow )
   IF HB_ISNIL( aC )
      aC := Array( CELL_LEN )
      aC[ CELL_INPUT ] := ""
      aC[ CELL_ALIGN ] := ""
      ::hCells[ DC_Key( nCol, nRow ) ] := aC
   ENDIF
   aC[ CELL_FMT ] := cFmt
   ::lModified := .T.
RETURN NIL

METHOD setAlign( nCol, nRow, cAlign ) CLASS DCSheet
   LOCAL aC := ::cell( nCol, nRow )
   IF ! HB_ISNIL( aC )
      aC[ CELL_ALIGN ] := cAlign
      ::lModified := .T.
   ENDIF
RETURN NIL

METHOD colWidth( nCol ) CLASS DCSheet
RETURN hb_HGetDef( ::hColW, nCol, ::nDefWidth )

METHOD setColWidth( nCol, nWidth ) CLASS DCSheet
   IF HB_ISNIL( nWidth )
      hb_HDel( ::hColW, nCol )
   ELSE
      ::hColW[ nCol ] := Max( 1, Min( 240, nWidth ) )
   ENDIF
   ::lModified := .T.
RETURN NIL

METHOD lastRow() CLASS DCSheet
   LOCAL n := 0, k
   FOR EACH k IN hb_HKeys( ::hCells )
      n := Max( n, DC_KeyRow( k ) )
   NEXT
RETURN n

METHOD lastCol() CLASS DCSheet
   LOCAL n := 0, k
   FOR EACH k IN hb_HKeys( ::hCells )
      n := Max( n, DC_KeyCol( k ) )
   NEXT
RETURN n

/*------------------------------------------------------------ dependencies */

/* forget what the formula in nKey used */
METHOD unlink( nKey ) CLASS DCSheet
   LOCAL aC := hb_HGetDef( ::hCells, nKey, NIL ), x, k, i
   IF HB_ISNIL( aC ) .OR. HB_ISNIL( aC[ CELL_AST ] )
      RETURN NIL
   ENDIF
   FOR EACH x IN DC_AstRefs( aC[ CELL_AST ] )
      IF x[ 1 ] == "R"
         k := DC_Key( x[ 2 ][ 1 ], x[ 2 ][ 2 ] )
         IF hb_HHasKey( ::hDeps, k ) .AND. hb_HHasKey( ::hDeps[ k ], nKey )
            hb_HDel( ::hDeps[ k ], nKey )
         ENDIF
      ENDIF
   NEXT
   FOR i := Len( ::aAreaDeps ) TO 1 STEP -1
      IF ::aAreaDeps[ i ][ 1 ] == nKey
         hb_ADel( ::aAreaDeps, i, .T. )
      ENDIF
   NEXT
RETURN NIL

/* remember what the formula in nKey uses */
METHOD link( nKey, xAst ) CLASS DCSheet
   LOCAL x, k
   FOR EACH x IN DC_AstRefs( xAst )
      DO CASE
      CASE x[ 1 ] == "R"
         k := DC_Key( x[ 2 ][ 1 ], x[ 2 ][ 2 ] )
         IF ! hb_HHasKey( ::hDeps, k )
            ::hDeps[ k ] := { => }
         ENDIF
         ::hDeps[ k ][ nKey ] := .T.
      CASE x[ 1 ] == "A"
         AAdd( ::aAreaDeps, { nKey, x[ 2 ][ 1 ], x[ 2 ][ 2 ], x[ 3 ][ 1 ], x[ 3 ][ 2 ] } )
      ENDCASE
   NEXT
RETURN NIL

/* the formulas that use nKey directly */
METHOD dependents( nKey ) CLASS DCSheet
   LOCAL aOut := {}, a, nCol := DC_KeyCol( nKey ), nRow := DC_KeyRow( nKey )
   IF hb_HHasKey( ::hDeps, nKey )
      AEval( hb_HKeys( ::hDeps[ nKey ] ), {| k | AAdd( aOut, k ) } )
   ENDIF
   FOR EACH a IN ::aAreaDeps
      IF nCol >= a[ 2 ] .AND. nCol <= a[ 4 ] .AND. nRow >= a[ 3 ] .AND. nRow <= a[ 5 ]
         AAdd( aOut, a[ 1 ] )
      ENDIF
   NEXT
RETURN aOut

/*------------------------------------------------------------ recalculation */

/* the changed cells: they and every formula that depends on them, in order */
METHOD recalc( aKeys ) CLASS DCSheet

   LOCAL aQueue := AClone( aKeys ), nKey, k, i := 1

   ::hDirty := { => }
   DO WHILE i <= Len( aQueue )
      nKey := aQueue[ i ]
      i++
      IF hb_HHasKey( ::hDirty, nKey )
         LOOP
      ENDIF
      IF hb_HHasKey( ::hCells, nKey ) .AND. ! HB_ISNIL( ::hCells[ nKey ][ CELL_AST ] )
         ::hDirty[ nKey ] := .T.
      ELSE
         ::hDirty[ nKey ] := .F.            /* a constant: only its dependents */
      ENDIF
      FOR EACH k IN ::dependents( nKey )
         AAdd( aQueue, k )
      NEXT
   ENDDO
   FOR EACH k IN hb_HKeys( ::hDirty )
      IF ! ::hDirty[ k ]
         hb_HDel( ::hDirty, k )
      ENDIF
   NEXT

   ::hBusy := { => }
   DO WHILE Len( ::hDirty ) > 0
      ::compute( hb_HKeyAt( ::hDirty, 1 ) )
   ENDDO
   ::hDirty := NIL
   ::hBusy := NIL

RETURN NIL

METHOD recalcAll() CLASS DCSheet
   LOCAL aKeys := {}, k
   FOR EACH k IN hb_HKeys( ::hCells )
      IF ! HB_ISNIL( ::hCells[ k ][ CELL_AST ] )
         AAdd( aKeys, k )
      ENDIF
   NEXT
   ::recalc( aKeys )
RETURN NIL

METHOD compute( nKey ) CLASS DCSheet
   LOCAL aC := ::hCells[ nKey ]
   IF hb_HHasKey( ::hBusy, nKey )          /* the formula uses itself */
      aC[ CELL_VALUE ] := DC_Err( "#CIRC!" )
      hb_HDel( ::hDirty, nKey )
      RETURN NIL
   ENDIF
   ::hBusy[ nKey ] := .T.
   aC[ CELL_VALUE ] := DC_SafeEval( aC[ CELL_AST ], Self )
   /* NaN is not equal to itself; infinity is beyond any double */
   IF HB_ISNUMERIC( aC[ CELL_VALUE ] ) .AND. ( !( aC[ CELL_VALUE ] == aC[ CELL_VALUE ] ) .OR. Abs( aC[ CELL_VALUE ] ) > 10 ^ 308 )
      aC[ CELL_VALUE ] := DC_Err( "#NUM!" )
   ENDIF
   hb_HDel( ::hBusy, nKey )
   IF hb_HHasKey( ::hDirty, nKey )
      hb_HDel( ::hDirty, nKey )
   ENDIF
RETURN NIL

/* a run time error in a formula (a bug, a strange argument) is #VALUE! in the
   cell, never the end of the program */
STATIC FUNCTION DC_SafeEval( xAst, oCtx )
   LOCAL xVal, bOld := ErrorBlock( {| e | Break( e ) } )
   BEGIN SEQUENCE
      xVal := DC_Eval( xAst, oCtx )
   RECOVER
      xVal := DC_Err( "#VALUE!" )
   END SEQUENCE
   ErrorBlock( bOld )
RETURN xVal

/* the value of a cell for a formula: computed first when it is still to compute */
METHOD evalCell( nCol, nRow ) CLASS DCSheet
   LOCAL nKey := DC_Key( nCol, nRow ), aC
   IF ! HB_ISNIL( ::hDirty ) .AND. hb_HHasKey( ::hDirty, nKey )
      IF hb_HHasKey( ::hBusy, nKey )
         RETURN DC_Err( "#CIRC!" )
      ENDIF
      ::compute( nKey )
   ELSEIF ! HB_ISNIL( ::hBusy ) .AND. hb_HHasKey( ::hBusy, nKey )
      RETURN DC_Err( "#CIRC!" )
   ENDIF
   aC := hb_HGetDef( ::hCells, nKey, NIL )
RETURN IIF( HB_ISNIL( aC ), NIL, aC[ CELL_VALUE ] )

METHOD evalName( cName ) CLASS DCSheet
   LOCAL x := hb_HGetDef( ::hNames, cName, NIL )
   IF HB_ISNIL( x )
      RETURN DC_Err( "#NAME?" )
   ENDIF
RETURN DC_Eval( x, Self )

/* the values of the filled cells of an area, row by row (empty cells left out) */
METHOD areaValues( a1, a2 ) CLASS DCSheet
   LOCAL aOut := {}, nCol, nRow, k, nCells
   nCells := ( a2[ 1 ] - a1[ 1 ] + 1 ) * ( a2[ 2 ] - a1[ 2 ] + 1 )
   IF nCells <= Len( ::hCells ) * 4
      FOR nRow := a1[ 2 ] TO a2[ 2 ]
         FOR nCol := a1[ 1 ] TO a2[ 1 ]
            IF hb_HHasKey( ::hCells, DC_Key( nCol, nRow ) )
               AAdd( aOut, ::evalCell( nCol, nRow ) )
            ENDIF
         NEXT
      NEXT
   ELSE                                    /* a big area on a small sheet */
      FOR EACH k IN ASort( hb_HKeys( ::hCells ) )
         nCol := DC_KeyCol( k )
         nRow := DC_KeyRow( k )
         IF nCol >= a1[ 1 ] .AND. nCol <= a2[ 1 ] .AND. nRow >= a1[ 2 ] .AND. nRow <= a2[ 2 ]
            AAdd( aOut, ::evalCell( nCol, nRow ) )
         ENDIF
      NEXT
   ENDIF
RETURN aOut

/* the area as rows of values, empty cells as NIL (VLOOKUP, INDEX ...) */
METHOD areaMatrix( a1, a2 ) CLASS DCSheet
   LOCAL aRows := {}, aRow, nCol, nRow
   FOR nRow := a1[ 2 ] TO Min( a2[ 2 ], a1[ 2 ] + 100000 )
      aRow := {}
      FOR nCol := a1[ 1 ] TO a2[ 1 ]
         AAdd( aRow, ::evalCell( nCol, nRow ) )
      NEXT
      AAdd( aRows, aRow )
   NEXT
RETURN aRows
