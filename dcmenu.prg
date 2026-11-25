/** @file dcmenu.prg
 * DigiCalc - the / menu in the control panel, in the manner of 1-2-3:
 * row 1 the entries of the level, row 2 what the marked one does (or the
 * entries under it). Left / Right mark, Enter or the first letter takes an
 * entry, Esc goes a level back.
 *
 *   an entry: { cName, cDescription, bAction | aSubEntries }
 *
 * And the file dialog (DC_FileDialog).
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "hbclass.ch"
#include "hbtui.ch"
#include "inkey.ch"

STATIC s_aStack := {}               /* { { aEntries, nMarked }, ... } */
STATIC s_oGrid

PROCEDURE DC_MenuOpen( oGrid, aRoot )
   s_oGrid := oGrid
   s_aStack := { { aRoot, 1 } }
   oGrid:lMenu := .T.
   oGrid:bMenuKey := {| n | DC_MenuKey( n ) }
   oGrid:bMenuLine := {| o | DC_MenuPaint( o ) }
   oGrid:message( "" )
RETURN

STATIC PROCEDURE DC_MenuClose()
   s_oGrid:lMenu := .F.
   s_aStack := {}
RETURN

STATIC PROCEDURE DC_MenuKey( nKey )

   LOCAL aTop := ATail( s_aStack ), aItems := aTop[ 1 ], n := aTop[ 2 ], c, i

   SWITCH nKey
   CASE K_LEFT
      aTop[ 2 ] := IIF( n > 1, n - 1, Len( aItems ) )
      RETURN
   CASE K_RIGHT
   CASE K_SPACE
      aTop[ 2 ] := IIF( n < Len( aItems ), n + 1, 1 )
      RETURN
   CASE K_HOME
      aTop[ 2 ] := 1
      RETURN
   CASE K_END
      aTop[ 2 ] := Len( aItems )
      RETURN
   CASE K_ENTER
      DC_MenuTake( aItems[ n ] )
      RETURN
   CASE K_ESC
   CASE K_BS
      hb_ADel( s_aStack, Len( s_aStack ), .T. )
      IF Empty( s_aStack )
         DC_MenuClose()
      ENDIF
      RETURN
   ENDSWITCH
   c := Upper( hb_keyChar( nKey ) )
   IF Len( c ) == 1
      FOR i := 1 TO Len( aItems )
         IF Upper( Left( aItems[ i ][ 1 ], 1 ) ) == c
            aTop[ 2 ] := i
            DC_MenuTake( aItems[ i ] )
            RETURN
         ENDIF
      NEXT
   ENDIF
RETURN

STATIC PROCEDURE DC_MenuTake( aItem )
   IF HB_ISARRAY( aItem[ 3 ] )
      AAdd( s_aStack, { aItem[ 3 ], 1 } )
   ELSE
      DC_MenuClose()
      s_oGrid:refresh()               /* the menu off the screen before a dialog opens */
      Eval( aItem[ 3 ] )
   ENDIF
RETURN

STATIC PROCEDURE DC_MenuPaint( oGrid )

   LOCAL cPanel := HTTheme():getColor( HT_CLR_WINDOW ), cSel := HTTheme():getColor( HT_CLR_MENU_BAR_SEL )
   LOCAL nW := oGrid:width, aTop, aItems, n, nX := 0, a, c := ""

   IF Empty( s_aStack )
      RETURN
   ENDIF
   aTop := ATail( s_aStack )
   aItems := aTop[ 1 ]
   n := aTop[ 2 ]
   DispOutAt( 1, 0, Space( nW ), cPanel )
   FOR EACH a IN aItems
      IF nX + Len( a[ 1 ] ) <= nW
         DispOutAt( 1, nX, a[ 1 ], IIF( a:__enumIndex() == n, cSel, cPanel ) )
      ENDIF
      nX += Len( a[ 1 ] ) + 2
   NEXT
   IF HB_ISARRAY( aItems[ n ][ 3 ] )
      AEval( aItems[ n ][ 3 ], {| x, i | c += IIF( i > 1, "  ", "" ) + x[ 1 ] } )
   ELSE
      c := aItems[ n ][ 2 ]
   ENDIF
   DispOutAt( 2, 0, PadR( c, nW ), cPanel )

RETURN

/*------------------------------------------------------------ the file dialog */

/* a file to open or a name to save under; cMasks: "*.xlsx;*.csv" */
FUNCTION DC_FileDialog( cTitle, cStart, cMasks )

   LOCAL dlg, oLbl, oName, oList, nW, nH, cResult := ""
   LOCAL cDir := hb_FNameDir( cStart ), cName := hb_FNameNameExt( cStart ), bTake

   IF Empty( cDir ) .OR. ! hb_DirExists( cDir )
      cDir := hb_cwd()
   ENDIF
   cDir := hb_DirSepAdd( cDir )
   nW := Min( 76, ht_ScreenMaxCol() - 3 )
   nH := Min( 22, ht_ScreenMaxRow() - 2 )

   dlg := HTDialog():new()
   dlg:setWindowTitle( cTitle )
   dlg:resize( nW, nH )
   dlg:move( Int( ( ht_ScreenMaxCol() + 1 - nW ) / 2 ), Int( ( ht_ScreenMaxRow() + 1 - nH ) / 2 ) )

   oLbl := HTLabel():new( "", dlg )
   oLbl:setGeometry( 1, 0, nW - 4, 1 )
   oName := DCPathEdit():new( dlg )
   oName:setGeometry( 1, 1, nW - 4, 1 )
   oName:setText( cName )
   oList := HTListBox():new( dlg )
   oList:setGeometry( 1, 3, nW - 4, nH - 7 )

   /* a folder: into it; a file: done */
   bTake := {| cTxt, cPath |
      cTxt := AllTrim( cTxt )
      IF Empty( cTxt )
         RETURN NIL
      ENDIF
      cPath := IIF( DC_IsAbsolute( cTxt ), cTxt, cDir + cTxt )
      IF hb_DirExists( cPath )
         cDir := IIF( cTxt == "..", DC_ParentDir( cDir ), hb_DirSepAdd( hb_DirSepToOS( cPath ) ) )
         DC_FillDir( oList, oLbl, cDir, cMasks )
         oName:setText( "" )
         dlg:repaint()
      ELSE
         cResult := cPath
         dlg:accept()
      ENDIF
      RETURN NIL
   }
   oName:bEnter := {|| Eval( bTake, oName:text ) }
   oList:onActivated := {|| Eval( bTake, oList:currentData() ) }
   oList:onChanged := {|| IIF( HB_ISSTRING( oList:currentData() ) .AND. ! hb_DirExists( cDir + oList:currentData() ), ;
                               ( oName:setText( oList:currentData() ), dlg:repaintChild( oName ) ), NIL ) }

   DC_FillDir( oList, oLbl, cDir, cMasks )
   dlg:addButtonBar( "OK", "Cancel" )

   IF dlg:exec() == HT_DIALOG_ACCEPTED .AND. Empty( cResult ) .AND. ! Empty( oName:text )
      cResult := IIF( DC_IsAbsolute( AllTrim( oName:text ) ), AllTrim( oName:text ), cDir + AllTrim( oName:text ) )
      IF hb_DirExists( cResult )
         cResult := ""
      ENDIF
   ENDIF

RETURN cResult

STATIC PROCEDURE DC_FillDir( oList, oLbl, cDir, cMasks )

   LOCAL aDirs := {}, aFiles := {}, a, cMask, aMasks := hb_ATokens( hb_defaultValue( cMasks, "*.*" ), ";" )

   FOR EACH a IN Directory( cDir + hb_osFileMask(), "D" )
      IF a[ 1 ] == "." .OR. a[ 1 ] == ".."
         LOOP
      ENDIF
      IF "D" $ a[ 5 ]
         AAdd( aDirs, a[ 1 ] )
      ELSE
         FOR EACH cMask IN aMasks
            IF hb_WildMatchI( AllTrim( cMask ), a[ 1 ], .T. )
               AAdd( aFiles, { a[ 1 ], a[ 2 ], a[ 3 ] } )
               EXIT
            ENDIF
         NEXT
      ENDIF
   NEXT
   ASort( aDirs, , , {| x, y | Upper( x ) < Upper( y ) } )
   ASort( aFiles, , , {| x, y | Upper( x[ 1 ] ) < Upper( y[ 1 ] ) } )

   oList:clear()
   IF ! DC_IsRoot( cDir )
      oList:addItem( "[..]", ".." )
   ENDIF
   FOR EACH a IN aDirs
      oList:addItem( "[" + a + "]", a )
   NEXT
   FOR EACH a IN aFiles
      oList:addItem( ht_cellPadR( a[ 1 ], Max( 10, oList:width - 24 ) ) + Str( a[ 2 ], 12 ) + "  " + DToC( a[ 3 ] ), a[ 1 ] )
   NEXT
   oList:setCurrentIndex( 1 )
   oLbl:setText( cDir )

RETURN

/* an input line of a dialog: Enter runs bEnter (not the next field) */
CLASS DCPathEdit FROM HTLineEdit
   DATA bEnter
   METHOD enterMovesFocus() INLINE .F.
   METHOD keyEvent( keyEvent )
ENDCLASS

METHOD PROCEDURE keyEvent( keyEvent ) CLASS DCPathEdit
   IF keyEvent:key == K_ENTER .AND. ! HB_ISNIL( ::bEnter )
      keyEvent:accept()
      Eval( ::bEnter )
      RETURN
   ENDIF
   ::super:keyEvent( keyEvent )
RETURN

FUNCTION DC_IsAbsolute( c )
RETURN Left( c, 1 ) $ "\/" .OR. SubStr( c, 2, 1 ) == ":"

STATIC FUNCTION DC_IsRoot( cDir )
   cDir := hb_DirSepDel( cDir )
RETURN Empty( cDir ) .OR. cDir == "/" .OR. ( Len( cDir ) == 2 .AND. SubStr( cDir, 2, 1 ) == ":" )

/* "C:\a\b\" -> "C:\a\" */
STATIC FUNCTION DC_ParentDir( cDir )
   LOCAL c := hb_DirSepDel( cDir ), n := RAt( hb_ps(), c )
RETURN IIF( n > 0, Left( c, n ), cDir )
