/** @file dcgrid.prg
 * DigiCalc - HTSheet: the worksheet on the screen, in the manner of 1-2-3:
 *
 *   row 0   A1: (F2) [W12] @SUM(B1..B5)                            READY
 *   row 1   the entry line (VALUE / LABEL / EDIT / POINT), a prompt or the / menu
 *   row 2   a message, or the description of the menu entry
 *   row 3       A        B        C        ...      (column letters)
 *   row 4..   1 cells ...
 *
 * Typing starts an entry: + - ( @ . = $ # or a digit a value or formula (VALUE),
 * anything else a label (LABEL). Enter or an arrow stores it (an arrow then
 * moves); Esc drops it. In a formula the arrows point at a cell (POINT), '.'
 * or ':' anchors a range there. F2 edits the cell, F5 goes to a cell, F9
 * recalculates, Del erases, End + arrow jumps to the edge of the data,
 * Shift + arrows select a range.
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

#define PANEL_ROWS   4               /* the 3 lines of the control panel + the column letters */

CREATE CLASS HTSheet FROM HTWidget

   EXPORTED:

   VAR oSheet
   VAR nCol      INIT 1               /* the cell pointer */
   VAR nRow      INIT 1
   VAR nTop      INIT 1               /* the first row / column on the screen */
   VAR nLeft     INIT 1
   VAR nAncCol                        /* the other corner of a selected range, NIL = none */
   VAR nAncRow
   VAR cMode     INIT "READY"         /* READY VALUE LABEL EDIT POINT MENU ERROR WAIT, or a prompt's */
   VAR cBuf      INIT ""              /* the entry line */
   VAR nPos      INIT 1               /* the cursor in it (a character) */
   VAR cPrompt   INIT ""              /* "Address to go to: " ... */
   VAR bPrompt                        /* {| cText | } when a prompt is answered */
   VAR cMsg      INIT ""              /* row 2 */
   VAR lEnd      INIT .F.             /* End was pressed: the next arrow jumps */
   VAR lLotus    INIT .F.             /* formulas shown the 1-2-3 way (@SUM(A1..A5)) */
   VAR lPoint    INIT .F.             /* pointing at a cell for the formula */
   VAR nPtCol
   VAR nPtRow
   VAR nPtAncCol
   VAR nPtAncRow
   VAR nPtStart                       /* where the pointed address starts in cBuf */
   VAR cEditMode INIT ""              /* the mode the entry started in (VALUE / LABEL / EDIT) */
   VAR bStatus                        /* {|| } after every change: the status line */
   VAR bSlash                         /* {|| } the / menu */
   VAR bMenuLine                      /* {| oSheetWidget, nRow | } paints rows 1-2 while the / menu is open */
   VAR lMenu     INIT .F.
   VAR bMenuKey                       /* {| nKey | } a key while the / menu is open */
   VAR bSheet                         /* {| nStep | } Ctrl-PgDn / Ctrl-PgUp: the next / previous sheet */
   VAR lRangePrompt INIT .F.          /* the prompt asks for a range: the arrows point */
   VAR lDrag     INIT .F.             /* the left button is down: moving selects */

   METHOD new( parent )
   METHOD paintEvent( event )
   METHOD keyEvent( keyEvent )
   METHOD mouseEvent( eventMouse )
   METHOD mouseDragWanted() INLINE .T.
   METHOD cursorPos()

   METHOD gridRows() INLINE Max( 1, ::Fheight - PANEL_ROWS )
   METHOD gutter()
   METHOD visibleCols()
   METHOD ensureVisible()
   METHOD moveTo( nCol, nRow, lExtend )
   METHOD jump( nKey, lExtend )
   METHOD selArea()
   METHOD clearSel() INLINE ::nAncCol := ::nAncRow := NIL
   METHOD refresh()

   METHOD startEntry( cFirst )
   METHOD editCell()
   METHOD commit( nThenKey )
   METHOD cancel()
   METHOD prompt( cPrompt, cDefault, bDone, lRange )
   METHOD message( cText )
   METHOD entryKey( nKey, lShift )
   METHOD wantsPoint()
   METHOD pointMove( nKey, lShift )
   METHOD pointText()

   METHOD cellText( nCol, nRow, nWidth, nAvail )
   METHOD paintPanel()
   METHOD paintHeader( nY, aCols )
   METHOD paintRow( nY, nRow, aCols )

ENDCLASS

METHOD new( parent ) CLASS HTSheet
   ::super:new( parent )
   ::setFocusPolicy( HT_FOCUS_STRONG )
   ::FisVisible := .T.
RETURN self

/*------------------------------------------------------------ geometry */

/* the width of the row numbers */
METHOD gutter() CLASS HTSheet
RETURN Max( 4, Len( hb_ntos( ::nTop + ::gridRows() ) ) + 1 )

/* { { nCol, nX, nWidth }, ... } of the columns on the screen from ::nLeft */
METHOD visibleCols() CLASS HTSheet
   LOCAL aCols := {}, nX := ::gutter(), nCol := ::nLeft, nW
   DO WHILE nX < ::Fwidth .AND. nCol <= DC_MAXCOL
      nW := ::oSheet:colWidth( nCol )
      AAdd( aCols, { nCol, nX, Min( nW, ::Fwidth - nX ) } )
      nX += nW
      nCol++
   ENDDO
RETURN aCols

METHOD ensureVisible() CLASS HTSheet
   LOCAL nCol := IIF( ::lPoint, ::nPtCol, ::nCol ), nRow := IIF( ::lPoint, ::nPtRow, ::nRow )
   LOCAL aCols, nX
   IF nRow < ::nTop
      ::nTop := nRow
   ELSEIF nRow >= ::nTop + ::gridRows()
      ::nTop := nRow - ::gridRows() + 1
   ENDIF
   IF nCol < ::nLeft
      ::nLeft := nCol
   ELSE
      /* the whole column on the screen */
      DO WHILE .T.
         aCols := ::visibleCols()
         nX := AScan( aCols, {| a | a[ 1 ] == nCol } )
         IF nX > 0 .AND. ( nX < Len( aCols ) .OR. aCols[ nX ][ 3 ] == ::oSheet:colWidth( nCol ) ) .OR. ::nLeft >= nCol
            EXIT
         ENDIF
         ::nLeft++
      ENDDO
   ENDIF
RETURN NIL

METHOD moveTo( nCol, nRow, lExtend ) CLASS HTSheet
   nCol := Max( 1, Min( DC_MAXCOL, nCol ) )
   nRow := Max( 1, Min( DC_MAXROW, nRow ) )
   IF hb_defaultValue( lExtend, .F. )
      IF HB_ISNIL( ::nAncCol )
         ::nAncCol := ::nCol
         ::nAncRow := ::nRow
      ENDIF
   ELSE
      ::clearSel()
   ENDIF
   ::nCol := nCol
   ::nRow := nRow
   ::ensureVisible()
RETURN NIL

/* the selected range { nC1, nR1, nC2, nR2 }; the cell pointer when none */
METHOD selArea() CLASS HTSheet
   IF HB_ISNIL( ::nAncCol )
      RETURN { ::nCol, ::nRow, ::nCol, ::nRow }
   ENDIF
RETURN { Min( ::nCol, ::nAncCol ), Min( ::nRow, ::nAncRow ), Max( ::nCol, ::nAncCol ), Max( ::nRow, ::nAncRow ) }

/* End + arrow (and Ctrl + arrow): to the edge of the filled cells, as 1-2-3 / Excel */
METHOD jump( nKey, lExtend ) CLASS HTSheet
   LOCAL nDC := 0, nDR := 0, nC := ::nCol, nR := ::nRow, lFilled, nLastR, nLastC
   SWITCH nKey
   CASE K_UP    ; nDR := -1 ; EXIT
   CASE K_DOWN  ; nDR := 1  ; EXIT
   CASE K_LEFT  ; nDC := -1 ; EXIT
   CASE K_RIGHT ; nDC := 1  ; EXIT
   CASE K_HOME                         /* End Home: the last cell with data */
      ::moveTo( Max( 1, ::oSheet:lastCol() ), Max( 1, ::oSheet:lastRow() ), lExtend )
      RETURN NIL
   OTHERWISE
      RETURN NIL
   ENDSWITCH
   nLastR := Max( ::oSheet:lastRow(), 1 )
   nLastC := Max( ::oSheet:lastCol(), 1 )
   lFilled := {| c, r, a | a := ::oSheet:cell( c, r ), ;
                           ! HB_ISNIL( a ) .AND. ( ! HB_ISNIL( a[ CELL_VALUE ] ) .OR. ! HB_ISNIL( a[ CELL_AST ] ) ) }
   IF Eval( lFilled, nC, nR ) .AND. Eval( lFilled, nC + nDC, nR + nDR )
      /* in a block: to its last filled cell */
      DO WHILE nC + nDC >= 1 .AND. nR + nDR >= 1 .AND. Eval( lFilled, nC + nDC, nR + nDR )
         nC += nDC
         nR += nDR
      ENDDO
   ELSE
      /* to the next filled cell; none: to the edge */
      nC += nDC
      nR += nDR
      DO WHILE nC >= 1 .AND. nR >= 1 .AND. nC <= nLastC .AND. nR <= nLastR .AND. ! Eval( lFilled, nC, nR )
         nC += nDC
         nR += nDR
      ENDDO
      IF nC < 1 .OR. nR < 1 .OR. nC > nLastC .OR. nR > nLastR
         nC := IIF( nDC < 0, 1, IIF( nDC > 0, Max( nLastC, ::nCol ), ::nCol ) )
         nR := IIF( nDR < 0, 1, IIF( nDR > 0, Max( nLastR, ::nRow ), ::nRow ) )
      ENDIF
   ENDIF
   ::moveTo( nC, nR, lExtend )
RETURN NIL

METHOD refresh() CLASS HTSheet
   IF ! HB_ISNIL( ::parent() ) .AND. ! HB_ISNIL( ::parent():windowId )
      ::parent():repaintChild( self )
   ENDIF
   IF HB_ISEVALITEM( ::bStatus )
      Eval( ::bStatus )
   ENDIF
RETURN NIL

/*------------------------------------------------------------ painting */

METHOD PROCEDURE paintEvent( event ) CLASS HTSheet

   LOCAL aCols, nY

   HB_SYMBOL_UNUSED( event )
   IF HB_ISNIL( ::oSheet )
      RETURN
   ENDIF
   ::paintPanel()
   aCols := ::visibleCols()
   ::paintHeader( PANEL_ROWS - 1, aCols )
   FOR nY := 0 TO ::gridRows() - 1
      ::paintRow( PANEL_ROWS + nY, ::nTop + nY, aCols )
   NEXT

RETURN

METHOD paintPanel() CLASS HTSheet

   LOCAL cPanel := HTTheme():getColor( HT_CLR_WINDOW )
   LOCAL cEntry := HTTheme():getColor( HT_CLR_LINEEDIT_FOCUSED )
   LOCAL cInd := HTTheme():getColor( HT_CLR_MENU_BAR_SEL )
   LOCAL nW := ::Fwidth, aC, c, cFmt, nCol, nRow, cMode, nOff

   nCol := IIF( ::lPoint, ::nPtCol, ::nCol )
   nRow := IIF( ::lPoint, ::nPtRow, ::nRow )
   aC := ::oSheet:cell( ::nCol, ::nRow )

   /* row 0: A1: (F2) [W12] contents                         MODE */
   c := DC_CellName( nCol, nRow ) + ":"
   IF ! ::lPoint .AND. ! HB_ISNIL( aC )
      cFmt := aC[ CELL_FMT ]
      IF ! Empty( cFmt )
         c += " (" + cFmt + ")"
      ENDIF
      IF hb_HHasKey( ::oSheet:hColW, ::nCol )
         c += " [W" + hb_ntos( ::oSheet:colWidth( ::nCol ) ) + "]"
      ENDIF
      c += " " + DC_PanelText( ::oSheet, ::nCol, ::nRow, ::lLotus )
   ELSEIF ! ::lPoint .AND. hb_HHasKey( ::oSheet:hColW, ::nCol )
      c += " [W" + hb_ntos( ::oSheet:colWidth( ::nCol ) ) + "]"
   ENDIF
   cMode := " " + IIF( ::lMenu, "MENU", ::cMode ) + " "
   DispOutAt( 0, 0, PadR( c, nW - Len( cMode ) ), cPanel )
   DispOutAt( 0, nW - Len( cMode ), cMode, cInd )

   /* rows 1-2: the / menu paints them itself */
   IF ::lMenu .AND. HB_ISEVALITEM( ::bMenuLine )
      Eval( ::bMenuLine, self, 1 )
      RETURN NIL
   ENDIF

   /* row 1: the entry line or a prompt */
   IF ::cMode == "READY" .OR. ::cMode == "ERROR"
      DispOutAt( 1, 0, Space( nW ), cPanel )
   ELSE
      nOff := Max( 0, Len( ::cPrompt ) + ::nPos - nW + 1 )
      DispOutAt( 1, 0, Space( nW ), cEntry )
      DispOutAt( 1, 0, SubStr( ::cPrompt + ::cBuf, nOff + 1, nW ), cEntry )
   ENDIF

   /* row 2: the message */
   DispOutAt( 2, 0, PadR( ::cMsg, nW ), cPanel )

RETURN NIL

METHOD paintHeader( nY, aCols ) CLASS HTSheet
   LOCAL cHead := HTTheme():getColor( HT_CLR_TAB_INACTIVE ), cCur := HTTheme():getColor( HT_CLR_TAB_ACTIVE )
   LOCAL a, nCur := IIF( ::lPoint, ::nPtCol, ::nCol ), aSel := ::selArea()
   DispOutAt( nY, 0, Space( ::Fwidth ), cHead )
   FOR EACH a IN aCols
      DispOutAt( nY, a[ 2 ], PadC( DC_ColName( a[ 1 ] ), a[ 3 ] ), ;
         IIF( a[ 1 ] == nCur .OR. ( a[ 1 ] >= aSel[ 1 ] .AND. a[ 1 ] <= aSel[ 3 ] .AND. ! HB_ISNIL( ::nAncCol ) ), cCur, cHead ) )
   NEXT
RETURN NIL

METHOD paintRow( nY, nRow, aCols ) CLASS HTSheet

   LOCAL cCell := HTTheme():getColor( HT_CLR_BROWSE_NORMAL )
   LOCAL cPtr := HTTheme():getColor( HT_CLR_BROWSE_SELECTED )
   LOCAL cSel := HTTheme():getColor( HT_CLR_LIST_SEL_UNFOC )
   LOCAL cHead := HTTheme():getColor( HT_CLR_TAB_INACTIVE ), cCurH := HTTheme():getColor( HT_CLR_TAB_ACTIVE )
   LOCAL nGut := ::gutter(), i, a, c, nW, nAvail, j, aSel := ::selArea(), lSel, cColor
   LOCAL nCurC := IIF( ::lPoint, ::nPtCol, ::nCol ), nCurR := IIF( ::lPoint, ::nPtRow, ::nRow )
   LOCAL aPt := IIF( ::lPoint .AND. ! HB_ISNIL( ::nPtAncCol ), ;
      { Min( ::nPtCol, ::nPtAncCol ), Min( ::nPtRow, ::nPtAncRow ), Max( ::nPtCol, ::nPtAncCol ), Max( ::nPtRow, ::nPtAncRow ) }, NIL )

   DispOutAt( nY, 0, PadL( hb_ntos( nRow ) + " ", nGut ), IIF( nRow == nCurR, cCurH, cHead ) )
   DispOutAt( nY, nGut, Space( ::Fwidth - nGut ), cCell )

   i := 1
   DO WHILE i <= Len( aCols )
      a := aCols[ i ]
      nW := a[ 3 ]
      c := ::cellText( a[ 1 ], nRow, nW, @nAvail )
      lSel := ( ! HB_ISNIL( ::nAncCol ) .AND. a[ 1 ] >= aSel[ 1 ] .AND. a[ 1 ] <= aSel[ 3 ] .AND. nRow >= aSel[ 2 ] .AND. nRow <= aSel[ 4 ] ) .OR. ;
              ( ! HB_ISNIL( aPt ) .AND. a[ 1 ] >= aPt[ 1 ] .AND. a[ 1 ] <= aPt[ 3 ] .AND. nRow >= aPt[ 2 ] .AND. nRow <= aPt[ 4 ] )
      cColor := IIF( a[ 1 ] == nCurC .AND. nRow == nCurR, cPtr, IIF( lSel, cSel, cCell ) )
      IF nAvail > nW
         /* a long label runs over the empty cells to its right */
         j := i + 1
         DO WHILE j <= Len( aCols ) .AND. HB_ISNIL( ::oSheet:cell( aCols[ j ][ 1 ], nRow ) ) .AND. ;
               ! ( aCols[ j ][ 1 ] == nCurC .AND. nRow == nCurR ) .AND. Len( c ) > aCols[ j ][ 2 ] - a[ 2 ]
            j++
         ENDDO
         nAvail := IIF( j <= Len( aCols ), aCols[ j ][ 2 ], ::Fwidth ) - a[ 2 ]
         DispOutAt( nY, a[ 2 ], PadR( Left( c, nW ), nW ), cColor )
         IF nAvail > nW
            DispOutAt( nY, a[ 2 ] + nW, PadR( SubStr( c, nW + 1, nAvail - nW ), nAvail - nW ), cCell )
         ENDIF
         i := j
         LOOP
      ENDIF
      DispOutAt( nY, a[ 2 ], c, cColor )
      i++
   ENDDO

RETURN NIL

/* the text of a cell in its column (nWidth wide). nAvail: the length of a label
   that wants to run over the next cells (> nWidth), else nWidth */
METHOD cellText( nCol, nRow, nWidth, nAvail ) CLASS HTSheet

   LOCAL aC := ::oSheet:cell( nCol, nRow ), x, c, cFmt, cAlign

   nAvail := nWidth
   IF HB_ISNIL( aC )
      RETURN Space( nWidth )
   ENDIF
   x := aC[ CELL_VALUE ]
   cFmt := IIF( Empty( aC[ CELL_FMT ] ), ::oSheet:cDefFmt, aC[ CELL_FMT ] )
   cAlign := aC[ CELL_ALIGN ]
   IF cFmt == "H"
      RETURN Space( nWidth )
   ENDIF
   IF cFmt == "T" .AND. ! HB_ISNIL( aC[ CELL_AST ] )
      c := ::oSheet:input( nCol, nRow, ::lLotus )
      nAvail := Max( nWidth, Len( c ) )
      RETURN IIF( Len( c ) > nWidth, c, PadR( c, nWidth ) )
   ENDIF
   DO CASE
   CASE HB_ISNIL( x )
      RETURN Space( nWidth )
   CASE HB_ISSTRING( x )
      IF Len( x ) > nWidth .AND. ( Empty( cAlign ) .OR. cAlign == "L" )
         nAvail := Len( x )
         RETURN x
      ENDIF
      x := Left( x, nWidth )
      RETURN IIF( cAlign == "R", PadL( x, nWidth ), IIF( cAlign == "C", PadC( x, nWidth ), PadR( x, nWidth ) ) )
   CASE HB_ISNUMERIC( x )
      c := DC_FmtValue( x, cFmt, nWidth - 1 )
      IF Len( c ) > nWidth - 1
         c := Replicate( "*", nWidth - 1 )
      ENDIF
      RETURN IIF( cAlign == "L", PadR( c, nWidth ), IIF( cAlign == "C", PadC( c, nWidth ), PadL( c + " ", nWidth ) ) )
   ENDCASE
   c := IIF( DC_IsErr( x ), x:cCode, IIF( x, "TRUE", "FALSE" ) )
RETURN PadC( Left( c, nWidth ), nWidth )

/* the hardware cursor: in the entry line while typing */
METHOD cursorPos() CLASS HTSheet
   LOCAL nOff
   IF ! ::hasFocus() .OR. ::cMode == "READY" .OR. ::cMode == "ERROR" .OR. ::lMenu .OR. ::lPoint
      RETURN NIL
   ENDIF
   nOff := Max( 0, Len( ::cPrompt ) + ::nPos - ::Fwidth + 1 )
RETURN { 1, Len( ::cPrompt ) + ::nPos - 1 - nOff }

/*------------------------------------------------------------ keys */

METHOD PROCEDURE keyEvent( keyEvent ) CLASS HTSheet

   LOCAL nKey := keyEvent:key, nShifts := hb_gtInfo( HB_GTI_KBDSHIFTS )
   LOCAL lShift := hb_bitAnd( nShifts, HB_GTI_KBD_SHIFT ) != 0
   LOCAL cChar, lEnd := ::lEnd

   IF HB_ISNIL( ::oSheet )
      RETURN
   ENDIF
   ::lEnd := .F.

   IF ::lMenu .AND. HB_ISEVALITEM( ::bMenuKey )
      keyEvent:accept()
      Eval( ::bMenuKey, nKey )
      ::refresh()
      RETURN
   ENDIF

   /* typing, editing, pointing, a prompt */
   IF !( ::cMode == "READY" .OR. ::cMode == "ERROR" )
      IF ::entryKey( nKey, lShift )
         keyEvent:accept()
         ::refresh()
      ENDIF
      RETURN
   ENDIF

   IF ::cMode == "ERROR"
      ::cMode := "READY"
      ::cMsg := ""
   ENDIF
   ::cMsg := ""

   SWITCH nKey
   CASE K_UP
   CASE K_DOWN
   CASE K_LEFT
   CASE K_RIGHT
      IF lEnd .OR. hb_bitAnd( nShifts, HB_GTI_KBD_CTRL ) != 0 .AND. ( nKey == K_UP .OR. nKey == K_DOWN )
         ::jump( nKey, lShift )
      ELSE
         ::moveTo( ::nCol + IIF( nKey == K_RIGHT, 1, IIF( nKey == K_LEFT, -1, 0 ) ), ;
                   ::nRow + IIF( nKey == K_DOWN, 1, IIF( nKey == K_UP, -1, 0 ) ), lShift )
      ENDIF
      EXIT
   CASE K_CTRL_LEFT                         /* a screen to the left / right (1-2-3) */
   CASE K_SH_TAB
      ::moveTo( Max( 1, ::nCol - Max( 1, Len( ::visibleCols() ) - 1 ) ), ::nRow, lShift )
      EXIT
   CASE K_CTRL_RIGHT
   CASE K_TAB
      ::moveTo( ::nCol + Max( 1, Len( ::visibleCols() ) - 1 ), ::nRow, lShift .AND. nKey != K_TAB )
      EXIT
   CASE K_PGUP
      ::nTop := Max( 1, ::nTop - ::gridRows() )
      ::moveTo( ::nCol, ::nRow - ::gridRows(), lShift )
      EXIT
   CASE K_PGDN
      ::nTop := Min( DC_MAXROW, ::nTop + ::gridRows() )
      ::moveTo( ::nCol, ::nRow + ::gridRows(), lShift )
      EXIT
   CASE K_HOME
      IF lEnd
         ::jump( K_HOME, lShift )
      ELSE
         ::nTop := 1
         ::nLeft := 1
         ::moveTo( 1, 1, lShift )
      ENDIF
      EXIT
   CASE K_CTRL_HOME
      ::nTop := 1
      ::nLeft := 1
      ::moveTo( 1, 1, lShift )
      EXIT
   CASE K_CTRL_PGDN
   CASE K_CTRL_PGUP
      IF HB_ISEVALITEM( ::bSheet )
         Eval( ::bSheet, IIF( nKey == K_CTRL_PGDN, 1, -1 ) )
      ENDIF
      EXIT
   CASE K_END
      ::lEnd := .T.
      EXIT
   CASE K_CTRL_END
      ::jump( K_HOME, lShift )
      EXIT
   CASE K_F2
      ::editCell()
      EXIT
   CASE K_F5
      ::prompt( "Address to go to: ", DC_CellName( ::nCol, ::nRow ), ;
         {| c, a | a := DC_ParseRef( StrTran( Upper( AllTrim( c ) ), "$", "" ) ), ;
                   IIF( HB_ISNIL( a ), ::message( "Not a cell address: " + c ), ;
                        ( ::nTop := Max( 1, a[ 2 ] - 2 ), ::moveTo( a[ 1 ], a[ 2 ] ) ) ) } )
      EXIT
   CASE K_F9
      ::oSheet:recalcAll()
      EXIT
   CASE K_DEL
      DC_EraseArea( ::oSheet, ::selArea() )
      EXIT
   CASE K_BS
      ::oSheet:clear( ::nCol, ::nRow )
      EXIT
   CASE K_ESC
      ::clearSel()
      EXIT
   CASE K_ENTER
      ::moveTo( ::nCol, ::nRow + 1 )
      EXIT
   OTHERWISE
      cChar := hb_keyChar( nKey )
      IF cChar == "/" .AND. HB_ISEVALITEM( ::bSlash )
         keyEvent:accept()
         Eval( ::bSlash )
         ::refresh()
         RETURN
      ENDIF
      IF Len( cChar ) == 1 .AND. hb_UCode( cChar ) >= 32
         ::startEntry( cChar )
         EXIT
      ENDIF
      RETURN                                  /* not ours: the menu bar, Alt-X ... */
   ENDSWITCH

   keyEvent:accept()
   ::refresh()

RETURN

/* a key while typing; .F. when it is not used */
METHOD entryKey( nKey, lShift ) CLASS HTSheet

   LOCAL cChar, bDone

   /* a prompt (F5, file names ...): only the line */
   IF ! Empty( ::cPrompt )
      SWITCH nKey
      CASE K_ENTER
         cChar := ::cBuf
         bDone := ::bPrompt
         ::cancel()                          /* first: the block may ask the next question */
         Eval( bDone, cChar )
         RETURN .T.
      CASE K_ESC
         ::cancel()
         RETURN .T.
      ENDSWITCH
   ENDIF

   /* pointing at a cell for the formula */
   IF ::lPoint
      SWITCH nKey
      CASE K_UP
      CASE K_DOWN
      CASE K_LEFT
      CASE K_RIGHT
      CASE K_PGUP
      CASE K_PGDN
      CASE K_HOME
         ::pointMove( nKey, lShift )
         RETURN .T.
      CASE K_ESC                              /* back: the address is dropped */
         ::cBuf := Left( ::cBuf, ::nPtStart - 1 )
         ::nPos := Len( ::cBuf ) + 1
         ::lPoint := .F.
         ::cMode := ::cEditMode
         RETURN .T.
      ENDSWITCH
      cChar := hb_keyChar( nKey )
      IF cChar == "." .OR. cChar == ":"       /* anchor a range here */
         IF HB_ISNIL( ::nPtAncCol )
            ::nPtAncCol := ::nPtCol
            ::nPtAncRow := ::nPtRow
            ::cBuf := Left( ::cBuf, ::nPtStart - 1 ) + ::pointText()
            ::nPos := Len( ::cBuf ) + 1
         ENDIF
         RETURN .T.
      ENDIF
      /* anything else: the address stays, typing goes on */
      ::lPoint := .F.
      ::cMode := ::cEditMode
   ENDIF

   SWITCH nKey
   CASE K_ENTER
      ::commit( 0 )
      RETURN .T.
   CASE K_ESC
      ::cancel()
      RETURN .T.
   CASE K_UP
   CASE K_DOWN
   CASE K_PGUP
   CASE K_PGDN
   CASE K_TAB
   CASE K_SH_TAB
      IF ( Empty( ::cPrompt ) .AND. ::wantsPoint() ) .OR. ::lRangePrompt
         ::pointMove( nKey, lShift )
         RETURN .T.
      ENDIF
      IF Empty( ::cPrompt )
         ::commit( nKey )
      ENDIF
      RETURN .T.
   CASE K_LEFT
   CASE K_RIGHT
      IF ( Empty( ::cPrompt ) .AND. ::wantsPoint() ) .OR. ::lRangePrompt
         ::pointMove( nKey, lShift )
         RETURN .T.
      ENDIF
      IF ::cMode == "EDIT" .OR. ! Empty( ::cPrompt )
         ::nPos := Max( 1, Min( Len( ::cBuf ) + 1, ::nPos + IIF( nKey == K_LEFT, -1, 1 ) ) )
      ELSE
         ::commit( nKey )                      /* 1-2-3: an arrow stores and moves */
      ENDIF
      RETURN .T.
   CASE K_HOME
      ::nPos := 1
      RETURN .T.
   CASE K_END
      ::nPos := Len( ::cBuf ) + 1
      RETURN .T.
   CASE K_BS
      IF ::nPos > 1
         ::cBuf := Left( ::cBuf, ::nPos - 2 ) + SubStr( ::cBuf, ::nPos )
         ::nPos--
      ENDIF
      RETURN .T.
   CASE K_DEL
      ::cBuf := Left( ::cBuf, ::nPos - 1 ) + SubStr( ::cBuf, ::nPos + 1 )
      RETURN .T.
   CASE K_F2                                  /* VALUE / LABEL -> EDIT: the arrows move in the line */
      IF Empty( ::cPrompt )
         ::cMode := "EDIT"
         ::cEditMode := "EDIT"
      ENDIF
      RETURN .T.
   ENDSWITCH

   cChar := hb_keyChar( nKey )
   IF Len( cChar ) == 1 .AND. hb_UCode( cChar ) >= 32
      ::cBuf := Left( ::cBuf, ::nPos - 1 ) + cChar + SubStr( ::cBuf, ::nPos )
      ::nPos++
      RETURN .T.
   ENDIF

RETURN .F.

/* a formula waits for an operand: the arrows point at a cell */
METHOD wantsPoint() CLASS HTSheet
   LOCAL c := RTrim( Left( ::cBuf, ::nPos - 1 ) )
   IF ::cMode == "LABEL" .OR. ::nPos <= Len( ::cBuf )
      RETURN .F.
   ENDIF
   IF !( Left( ::cBuf, 1 ) $ "=+-(@" )
      RETURN .F.
   ENDIF
RETURN Right( c, 1 ) $ "=+-*/^(,;<>&:" .OR. Right( c, 1 ) == "." .AND. Len( c ) > 1 .OR. c == "@"

METHOD pointMove( nKey, lShift ) CLASS HTSheet
   IF ! ::lPoint                     /* from the cell being typed; the arrow moves at once */
      ::lPoint := .T.
      ::cMode := "POINT"
      ::nPtCol := ::nCol
      ::nPtRow := ::nRow
      ::nPtAncCol := ::nPtAncRow := NIL
      /* a range prompt: the pointed range replaces the text */
      ::nPtStart := IIF( ::lRangePrompt, 1, Len( ::cBuf ) + 1 )
   ENDIF
   IF nKey != 0
      IF lShift .AND. HB_ISNIL( ::nPtAncCol )
         ::nPtAncCol := ::nPtCol
         ::nPtAncRow := ::nPtRow
      ENDIF
      SWITCH nKey
      CASE K_UP    ; ::nPtRow := Max( 1, ::nPtRow - 1 ) ; EXIT
      CASE K_DOWN  ; ::nPtRow := Min( DC_MAXROW, ::nPtRow + 1 ) ; EXIT
      CASE K_LEFT  ; ::nPtCol := Max( 1, ::nPtCol - 1 ) ; EXIT
      CASE K_RIGHT ; ::nPtCol := Min( DC_MAXCOL, ::nPtCol + 1 ) ; EXIT
      CASE K_PGUP  ; ::nPtRow := Max( 1, ::nPtRow - ::gridRows() ) ; EXIT
      CASE K_PGDN  ; ::nPtRow := Min( DC_MAXROW, ::nPtRow + ::gridRows() ) ; EXIT
      CASE K_HOME  ; ::nPtRow := ::nPtCol := 1 ; EXIT
      ENDSWITCH
   ENDIF
   ::cBuf := Left( ::cBuf, ::nPtStart - 1 ) + ::pointText()
   ::nPos := Len( ::cBuf ) + 1
   ::ensureVisible()
RETURN NIL

/* the pointed cell or range, as the entry is typed (Excel A1:B5, 1-2-3 A1..B5) */
METHOD pointText() CLASS HTSheet
   LOCAL lLotus := IIF( ::lRangePrompt, ::lLotus, !( Left( ::cBuf, 1 ) == "=" ) )
   IF HB_ISNIL( ::nPtAncCol )
      RETURN DC_CellName( ::nPtCol, ::nPtRow )
   ENDIF
RETURN DC_CellName( Min( ::nPtCol, ::nPtAncCol ), Min( ::nPtRow, ::nPtAncRow ) ) + IIF( lLotus, "..", ":" ) + ;
       DC_CellName( Max( ::nPtCol, ::nPtAncCol ), Max( ::nPtRow, ::nPtAncRow ) )

/*------------------------------------------------------------ entries */

METHOD startEntry( cFirst ) CLASS HTSheet
   ::clearSel()
   ::cBuf := cFirst
   ::nPos := 2
   ::cPrompt := ""
   ::cMode := IIF( cFirst $ "+-(@.=$#0123456789", "VALUE", "LABEL" )
   ::cEditMode := ::cMode
   ::cMsg := ""
RETURN NIL

METHOD editCell() CLASS HTSheet
   ::clearSel()
   ::cBuf := ::oSheet:input( ::nCol, ::nRow, ::lLotus )
   ::nPos := Len( ::cBuf ) + 1
   ::cPrompt := ""
   ::cMode := "EDIT"
   ::cEditMode := "EDIT"
   ::cMsg := ""
RETURN NIL

/* store the entry; nThenKey: the arrow / key that moves afterwards (0 = stay) */
METHOD commit( nThenKey ) CLASS HTSheet
   LOCAL cErr := ""
   IF ! ::oSheet:setInput( ::nCol, ::nRow, ::cBuf, @cErr )
      /* as 1-2-3: the entry stays, to be corrected */
      ::cMode := "EDIT"
      ::cEditMode := "EDIT"
      ::lPoint := .F.
      ::cMsg := "Error: " + cErr
      RETURN .F.
   ENDIF
   ::cMode := "READY"
   ::lPoint := .F.
   ::cBuf := ""
   ::cMsg := ""
   SWITCH nThenKey
   CASE K_UP    ; ::moveTo( ::nCol, ::nRow - 1 ) ; EXIT
   CASE K_DOWN  ; ::moveTo( ::nCol, ::nRow + 1 ) ; EXIT
   CASE K_LEFT  ; ::moveTo( ::nCol - 1, ::nRow ) ; EXIT
   CASE K_RIGHT ; ::moveTo( ::nCol + 1, ::nRow ) ; EXIT
   CASE K_TAB   ; ::moveTo( ::nCol + 1, ::nRow ) ; EXIT
   CASE K_SH_TAB ; ::moveTo( ::nCol - 1, ::nRow ) ; EXIT
   CASE K_PGUP  ; ::moveTo( ::nCol, ::nRow - ::gridRows() ) ; EXIT
   CASE K_PGDN  ; ::moveTo( ::nCol, ::nRow + ::gridRows() ) ; EXIT
   ENDSWITCH
RETURN .T.

METHOD cancel() CLASS HTSheet
   ::lRangePrompt := .F.
   ::cMode := "READY"
   ::lPoint := .F.
   ::cBuf := ""
   ::cPrompt := ""
   ::bPrompt := NIL
   ::nPos := 1
RETURN NIL

/* a question in the entry line: bDone( cAnswer ) after Enter, nothing after Esc */
METHOD prompt( cPrompt, cDefault, bDone, lRange ) CLASS HTSheet
   ::lRangePrompt := hb_defaultValue( lRange, .F. )
   ::cPrompt := cPrompt
   ::cBuf := hb_defaultValue( cDefault, "" )
   ::nPos := Len( ::cBuf ) + 1
   ::bPrompt := bDone
   ::cMode := IIF( ::lRangePrompt, "POINT", "EDIT" )
   ::cEditMode := ::cMode
   ::cMsg := IIF( ::lRangePrompt, "Type a range, or point at it with the arrows ( . anchors), Enter", "" )
RETURN NIL

METHOD message( cText ) CLASS HTSheet
   ::cMsg := cText
RETURN NIL

/*------------------------------------------------------------ mouse */

METHOD PROCEDURE mouseEvent( eventMouse ) CLASS HTSheet

   LOCAL nR := eventMouse:mouseRow, nC := eventMouse:mouseCol, aCols, a, nRow, nCol := 0

   IF HB_ISNIL( ::oSheet )
      RETURN
   ENDIF
   SWITCH eventMouse:nKey
   CASE K_MWFORWARD
      ::nTop := Max( 1, ::nTop - 3 )
      ::refresh()
      RETURN
   CASE K_MWBACKWARD
      ::nTop := Min( DC_MAXROW, ::nTop + 3 )
      ::refresh()
      RETURN
   CASE K_LBUTTONUP
      ::lDrag := .F.
      RETURN
   CASE K_LBUTTONDOWN
   CASE K_MOUSEMOVE
      IF eventMouse:nKey == K_MOUSEMOVE .AND. ! ::lDrag
         RETURN
      ENDIF
      IF nR < PANEL_ROWS .OR. nC < ::gutter()
         RETURN
      ENDIF
      IF eventMouse:nKey == K_LBUTTONDOWN
         ::lDrag := .T.
      ENDIF
      nRow := ::nTop + nR - PANEL_ROWS
      aCols := ::visibleCols()
      FOR EACH a IN aCols
         IF nC >= a[ 2 ] .AND. nC < a[ 2 ] + a[ 3 ]
            nCol := a[ 1 ]
         ENDIF
      NEXT
      IF nCol == 0
         RETURN
      ENDIF
      IF ::lPoint .OR. ( ::cMode $ "VALUE|EDIT" .AND. Empty( ::cPrompt ) .AND. ::wantsPoint() )
         IF ! ::lPoint
            ::pointMove( 0, .F. )
         ENDIF
         ::nPtCol := nCol
         ::nPtRow := nRow
         ::cBuf := Left( ::cBuf, ::nPtStart - 1 ) + ::pointText()
         ::nPos := Len( ::cBuf ) + 1
      ELSEIF ::cMode == "READY"
         ::moveTo( nCol, nRow, eventMouse:nKey == K_MOUSEMOVE .OR. hb_bitAnd( hb_gtInfo( HB_GTI_KBDSHIFTS ), HB_GTI_KBD_SHIFT ) != 0 )
      ENDIF
      ::refresh()
      RETURN
   ENDSWITCH

RETURN

/*------------------------------------------------------------ helpers */

/* row 0 of the panel: what is in the cell (the formula, or the label with its prefix) */
FUNCTION DC_PanelText( oSheet, nCol, nRow, lLotus )
   LOCAL aC := oSheet:cell( nCol, nRow ), x
   IF HB_ISNIL( aC )
      RETURN ""
   ENDIF
   IF ! HB_ISNIL( aC[ CELL_AST ] )
      RETURN oSheet:input( nCol, nRow, lLotus )
   ENDIF
   x := aC[ CELL_VALUE ]
   IF HB_ISSTRING( x )
      RETURN IIF( aC[ CELL_ALIGN ] == "R", '"', IIF( aC[ CELL_ALIGN ] == "C", "^", "'" ) ) + x
   ENDIF
RETURN aC[ CELL_INPUT ]

FUNCTION DC_EraseArea( oSheet, aA )
   LOCAL nCol, nRow
   FOR nRow := aA[ 2 ] TO aA[ 4 ]
      FOR nCol := aA[ 1 ] TO aA[ 3 ]
         oSheet:clear( nCol, nRow )
      NEXT
   NEXT
RETURN NIL

/* a number in a format: G general, Fn fixed, ,n comma, Cn currency, Pn percent,
   D1 dd.mm.yyyy, D2 dd-mmm-yy, D3 dd.mm., D4 mm.yyyy, D5 hh:mm, S scientific */
FUNCTION DC_FmtValue( x, cFmt, nWidth )
   LOCAL cK := Left( cFmt, 1 ), nDec := Val( SubStr( cFmt, 2 ) ), d, c, n
   SWITCH cK
   CASE "F"
      RETURN DC_FmtNumber( x, nDec, .F. )
   CASE ","
      RETURN DC_FmtNumber( x, nDec, .T. )
   CASE "C"
      RETURN IIF( x < 0, "-", "" ) + DC_Currency() + DC_FmtNumber( Abs( x ), nDec, .T. )
   CASE "P"
      RETURN DC_FmtNumber( x * 100, nDec, .F. ) + "%"
   CASE "S"
      RETURN DC_Sci( x, Max( 0, nDec ) )
   CASE "D"
      IF cFmt == "D5"
         n := Round( ( x - Int( x ) ) * 1440, 0 )
         RETURN StrZero( Int( n / 60 ) % 24, 2 ) + ":" + StrZero( n % 60, 2 )
      ENDIF
      d := DC_SerialDate( x )
      SWITCH cFmt
      CASE "D2" ; RETURN StrZero( Day( d ), 2 ) + "-" + Left( CMonth( d ), 3 ) + "-" + Right( StrZero( Year( d ), 4 ), 2 )
      CASE "D3" ; RETURN StrZero( Day( d ), 2 ) + "." + StrZero( Month( d ), 2 ) + "."
      CASE "D4" ; RETURN StrZero( Month( d ), 2 ) + "." + StrZero( Year( d ), 4 )
      ENDSWITCH
      RETURN StrZero( Day( d ), 2 ) + "." + StrZero( Month( d ), 2 ) + "." + StrZero( Year( d ), 4 )
   ENDSWITCH
   /* General: as it is, shorter when it does not fit */
   c := DC_NumText( x )
   IF Len( c ) > nWidth
      IF Abs( x ) >= 1 .AND. "." $ c .AND. At( ".", c ) <= nWidth
         c := DC_NumText( Round( x, Max( 0, nWidth - At( ".", c ) ) ) )
      ENDIF
      IF Len( c ) > nWidth
         c := DC_Sci( x, Max( 0, nWidth - 7 ) )
      ENDIF
   ENDIF
RETURN c

STATIC FUNCTION DC_Sci( x, nDec )
   LOCAL nExp := 0, m := Abs( x )
   IF m != 0
      nExp := Floor( Log10( m ) )
      m := m / 10 ^ nExp
      IF Round( m, nDec ) >= 10
         m /= 10
         nExp++
      ENDIF
   ENDIF
RETURN IIF( x < 0, "-", "" ) + LTrim( Str( m, 20, nDec ) ) + "E" + IIF( nExp < 0, "-", "+" ) + StrZero( Abs( nExp ), 2 )

