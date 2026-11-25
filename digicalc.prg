/** @file digicalc.prg
 * DigiCalc - a spreadsheet in character mode, in the manner of Lotus 1-2-3,
 * on Harbour-CUI. Formulas in Excel (=SUM(A1:A5)) or 1-2-3 (@SUM(A1..A5))
 * syntax; .xlsx through the XLSX class.
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
STATIC s_oSheet
STATIC s_cFile := ""

PROCEDURE Main( ... )

   hb_cdpSelect( "UTF8EX" )

   s_oApp := HTApplication():new()
   ht_cuiSetupDisplay( s_oApp:rows(), s_oApp:cols() )
   hb_gtInfo( HB_GTI_WINTITLE, "DigiCalc" )

   s_oMain := DCMain():new()
   s_oMain:setWindowTitle( "DigiCalc" )
   s_oMain:show()

   s_oSheet := DCSheet():new( "Sheet1" )
   s_oGrid := HTSheet():new( s_oMain )
   s_oGrid:oSheet := s_oSheet
   s_oGrid:bStatus := {|| dc_Status() }
   s_oStat := HTStatusBar():new( s_oMain )
   s_oStat:addSection( "" )
   dc_Arrange()
   dc_BuildMenu()

   s_oGrid:setFocus()
   dc_Status()
   s_oApp:exec()

RETURN

/*------------------------------------------------------------ the main window */

CLASS DCMain FROM HTMainWindow
   METHOD keyEvent( keyEvent )
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

METHOD PROCEDURE screenResized() CLASS DCMain
   ::super:screenResized()
   dc_Arrange()
RETURN

/* the sheet from under the menu bar down to the status line */
STATIC PROCEDURE dc_Arrange()
   s_oGrid:setGeometry( 0, 0, s_oMain:width, Max( 6, s_oMain:height - 2 ) )
   s_oStat:setGeometry( 0, s_oMain:height - 2, s_oMain:width, 1 )
RETURN

STATIC PROCEDURE dc_Status()
   LOCAL aA := s_oGrid:selArea(), c
   c := " " + IIF( Empty( s_cFile ), "(new)", hb_FNameNameExt( s_cFile ) ) + IIF( s_oSheet:lModified, " *", "" ) + ;
        "   " + s_oSheet:cName
   IF ! HB_ISNIL( s_oGrid:nAncCol )
      c += "   " + DC_CellName( aA[ 1 ], aA[ 2 ] ) + ":" + DC_CellName( aA[ 3 ], aA[ 4 ] )
   ENDIF
   c := PadR( c, Max( 1, s_oMain:width - 30 ) ) + PadL( IIF( s_oGrid:lEnd, "END  ", "" ) + ;
        IIF( s_oGrid:lLotus, "1-2-3 ", "Excel " ) + "/ menu  F1 help ", 30 )
   s_oStat:setSection( 1, c )
RETURN

/*------------------------------------------------------------ the menu bar */

STATIC PROCEDURE dc_BuildMenu()

   LOCAL oMenu

   s_oBar := s_oMain:menuBar()
   s_oBar:setBarRow( 0 )

   oMenu := s_oBar:addMenu( "&File" )
   dc_Item( oMenu, "&New", "", NIL, {|| dc_New() } )
   oMenu:addSeparator()
   dc_Item( oMenu, "E&xit", "Alt-X", NIL, {|| dc_Exit() } )

   oMenu := s_oBar:addMenu( "&Edit" )
   dc_Item( oMenu, "&Edit cell", "F2", NIL, {|| s_oGrid:editCell() } )
   dc_Item( oMenu, "&Go to", "F5", NIL, {|| dc_Key( K_F5 ) } )
   dc_Item( oMenu, "E&rase", "Del", NIL, {|| DC_EraseArea( s_oSheet, s_oGrid:selArea() ) } )
   oMenu:addSeparator()
   dc_Item( oMenu, "Re&calculate", "F9", NIL, {|| s_oSheet:recalcAll() } )

   oMenu := s_oBar:addMenu( "&View" )
   dc_Item( oMenu, "Formulas as &1-2-3", "", NIL, {|| s_oGrid:lLotus := .T. } )
   dc_Item( oMenu, "Formulas as &Excel", "", NIL, {|| s_oGrid:lLotus := .F. } )

   oMenu := s_oBar:addMenu( "&Help" )
   dc_Item( oMenu, "&Keys", "F1", K_F1, {|| dc_Help() } )
   dc_Item( oMenu, "&About", "", NIL, {|| dc_About() } )

RETURN

STATIC PROCEDURE dc_Item( oMenu, cText, cKey, nKey, bAction )
   LOCAL oAct := oMenu:addAction( PadR( cText, IIF( "&" $ cText, 21, 20 ) ) + PadL( cKey, 8 ) )
   IF ! HB_ISNIL( nKey )
      oAct:setShortcut( nKey )
   ENDIF
   oAct:onTriggered := {|| Eval( bAction ), s_oGrid:setFocus(), s_oGrid:refresh() }
RETURN

/* a key as if it was pressed in the sheet (menu entries of F5 ...) */
STATIC PROCEDURE dc_Key( nKey )
   LOCAL oEv := HTKeyEvent():new( nKey )
   s_oGrid:keyEvent( oEv )
RETURN

/*------------------------------------------------------------ commands */

STATIC PROCEDURE dc_New()
   IF ! dc_AskSave()
      RETURN
   ENDIF
   s_oSheet := DCSheet():new( "Sheet1" )
   s_oGrid:oSheet := s_oSheet
   s_oGrid:nCol := s_oGrid:nRow := s_oGrid:nTop := s_oGrid:nLeft := 1
   s_oGrid:clearSel()
   s_cFile := ""
RETURN

STATIC PROCEDURE dc_Exit()
   IF dc_AskSave()
      s_oApp:quit()
   ENDIF
RETURN

/* .T. when it may go on (nothing to save, saved, or not wanted) */
STATIC FUNCTION dc_AskSave()
   IF ! s_oSheet:lModified
      RETURN .T.
   ENDIF
RETURN HTMessageBox():question( "DigiCalc", "The worksheet has been changed and not saved." + Chr( 10 ) + "Leave it anyway?" )

STATIC PROCEDURE dc_About()
   HTMessageBox():information( "DigiCalc", "DigiCalc " + DC_VERSION + " - spreadsheet in character mode" + Chr( 10 ) + ;
      "Formulas: Excel (=SUM(A1:A5)) and 1-2-3 (@SUM(A1..A5))" + Chr( 10 ) + ;
      "Copyright (c) 2026 Srđan Dragojlović - GPL 3 or commercial license" )
RETURN

STATIC PROCEDURE dc_Help()
   HTMessageBox():information( "DigiCalc - keys", ;
      "Type a number, text or formula, then Enter or an arrow." + Chr( 10 ) + ;
      "Formula: =A1+B1  =SUM(A1:A5)   or 1-2-3: +A1+B1  @SUM(A1..A5)" + Chr( 10 ) + ;
      "Text: 'left  " + '"' + "right  ^center (1-2-3 prefixes)" + Chr( 10 ) + ;
      "In a formula the arrows point at a cell; . or : anchors a range." + Chr( 10 ) + ;
      "F2 edit   F5 go to   F9 recalculate   Del erase   / menu" + Chr( 10 ) + ;
      "End + arrow: edge of the data   End Home: last cell   Home: A1" + Chr( 10 ) + ;
      "Shift + arrows: select   Tab / Shift-Tab: a screen right / left" + Chr( 10 ) + ;
      "Alt-X: exit" )
RETURN
