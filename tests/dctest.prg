/* DigiCalc - tests of the core (console):  hbmk2 dctest.hbp && dctest */

#include "../digicalc.ch"

REQUEST HB_CODEPAGE_UTF8EX

STATIC s_nOk := 0, s_nBad := 0

PROCEDURE Main()

   LOCAL oS, cErr

   hb_cdpSelect( "UTF8EX" )

   ? "== addresses"
   Eq( DC_ColName( 1 ), "A" )
   Eq( DC_ColName( 26 ), "Z" )
   Eq( DC_ColName( 27 ), "AA" )
   Eq( DC_ColName( 16384 ), "XFD" )
   Eq( DC_ColNum( "XFD" ), 16384 )
   Eq( hb_ValToExp( DC_ParseRef( "$B$12" ) ), hb_ValToExp( { 2, 12, .T., .T. } ) )
   Eq( DC_ParseRef( "SUM" ), NIL )

   ? "== parse and text: Excel and 1-2-3"
   Eq( Txt( "SUM(A1:A5)*2" ), "SUM(A1:A5)*2" )
   Eq( Txt( "@SUM(A1..A5)*2" ), "SUM(A1:A5)*2" )
   Eq( Txt( "@sum(a1.a5)" ), "SUM(A1:A5)" )
   Eq( Txt( "+A1+B1" ), "+A1+B1" )
   Eq( Txt( "A1>0#AND#B1<5" ), "AND(A1>0,B1<5)" )
   Eq( Lot( "AND(A1>0,B1<5)" ), "A1>0#AND#B1<5" )
   Eq( Lot( "AVERAGE(B2:B9)" ), "@AVG(B2..B9)" )
   Eq( Txt( "@AVG(B2..B9)" ), "AVERAGE(B2:B9)" )
   Eq( Txt( "@PI*2" ), "PI()*2" )
   Eq( Txt( "@VLOOKUP(A1,B1..D9,2)" ), "VLOOKUP(A1,B1:D9,3,TRUE)" )
   Eq( Txt( "@INDEX(B1..D9,1,2)" ), "INDEX(B1:D9,3,2)" )
   Eq( Txt( "@PMT(1000,0.01,12)" ), "-PMT(0.01,12,1000)" )
   Eq( Txt( "$A$1*B$2" ), "$A$1*B$2" )
   Eq( Txt( '"a""b"&C1' ), '"a""b"&C1' )
   Eq( Txt( "IF(A1;1;2)" ), "IF(A1,1,2)" )
   Eq( Txt( "10%" ), "10%" )
   Eq( Txt( ".5+1E3" ), "0.5+1000" )
   Eq( ParseErr( "SUM(A1" ), "Missing ) after the arguments of SUM" )
   Eq( ParseErr( "1+" ), "Incomplete formula" )
   Eq( ParseErr( '"abc' ), "Missing closing quote" )

   ? "== values"
   oS := DCSheet():new()
   Set1( oS, "A1", "10" )
   Set1( oS, "A2", "20" )
   Set1( oS, "A3", "30" )
   Set1( oS, "B1", "=SUM(A1:A3)" )
   Eq( V( oS, "B1" ), 60 )
   Set1( oS, "B2", "@AVG(A1..A3)" )
   Eq( V( oS, "B2" ), 20 )
   Set1( oS, "B3", "+A1*2+A2" )
   Eq( V( oS, "B3" ), 40 )
   Set1( oS, "B4", "=2+3*4^2" )
   Eq( V( oS, "B4" ), 50 )
   Set1( oS, "B5", "=-2^2" )
   Eq( V( oS, "B5" ), 4 )
   Set1( oS, "B6", "=A1/0" )
   Eq( V( oS, "B6" ), "#DIV/0!" )
   Set1( oS, "B7", '="Total: "&B1' )
   Eq( V( oS, "B7" ), "Total: 60" )
   Set1( oS, "B8", '=IF(A1>5,"big","small")' )
   Eq( V( oS, "B8" ), "big" )
   Set1( oS, "B9", "+A1>5#AND#A2<5" )
   Eq( V( oS, "B9" ), .F. )
   Set1( oS, "C1", "'Label" )
   Eq( V( oS, "C1" ), "Label" )
   Eq( oS:cell( 3, 1 )[ CELL_ALIGN ], "L" )
   Set1( oS, "C2", '^Center' )
   Eq( oS:cell( 3, 2 )[ CELL_ALIGN ], "C" )
   Set1( oS, "C3", "plain text" )
   Eq( V( oS, "C3" ), "plain text" )
   Eq( oS:input( 2, 2, .T. ), "@AVG(A1..A3)" )
   Eq( oS:input( 2, 2, .F. ), "=AVERAGE(A1:A3)" )
   Eq( oS:input( 2, 3, .T. ), "+A1*2+A2" )
   Eq( oS:setInput( 4, 1, "=SUM(", @cErr ), .F. )

   ? "== recalculation"
   Set1( oS, "A1", "100" )
   Eq( V( oS, "B1" ), 150 )                  /* =SUM(A1:A3) */
   Eq( V( oS, "B3" ), 220 )                  /* +A1*2+A2 */
   Eq( V( oS, "B7" ), "Total: 150" )         /* through B1 */
   Set1( oS, "A3", "" )
   Eq( V( oS, "B1" ), 120 )
   Set1( oS, "D1", "=D2+1" )
   Set1( oS, "D2", "=D3+1" )
   Set1( oS, "D3", "5" )
   Eq( V( oS, "D1" ), 7 )
   Set1( oS, "D3", "=D1" )
   Eq( V( oS, "D1" ), "#CIRC!" )
   Set1( oS, "D3", "1" )
   Eq( V( oS, "D1" ), 3 )

   ? "== functions"
   oS := DCSheet():new()
   Set1( oS, "A1", "Apple" )   ; Set1( oS, "B1", "3" )
   Set1( oS, "A2", "Banana" )  ; Set1( oS, "B2", "5" )
   Set1( oS, "A3", "Cherry" )  ; Set1( oS, "B3", "7" )
   Eq( F( oS, 'VLOOKUP("Banana",A1:B3,2,FALSE)' ), 5 )
   Eq( F( oS, '@VLOOKUP("Cherry",A1..B3,1)' ), 7 )
   Eq( F( oS, "INDEX(A1:B3,3,1)" ), "Cherry" )
   Eq( F( oS, 'MATCH("Banana",A1:A3,0)' ), 2 )
   Eq( F( oS, 'COUNTIF(B1:B3,">4")' ), 2 )
   Eq( F( oS, 'SUMIF(A1:A3,"B*",B1:B3)' ), 5 )
   Eq( F( oS, "COUNT(A1:B3)" ), 3 )
   Eq( F( oS, "COUNTA(A1:B3)" ), 6 )
   Eq( F( oS, "MAX(B1:B3)-MIN(B1:B3)" ), 4 )
   Eq( F( oS, "ROUND(2.345,2)" ), 2.35 )
   Eq( F( oS, "ROUND(-2.5,0)" ), -3 )
   Eq( F( oS, "MOD(-7,3)" ), 2 )
   Eq( F( oS, "LEFT(A2,3)&UPPER(MID(A3,2,3))" ), "BanHER" )
   Eq( F( oS, 'LEN(TRIM("  a  b "))' ), 3 )
   Eq( F( oS, "@STRING(3.14159,2)" ), "3.14" )
   Eq( F( oS, 'TEXT(1234.5,"#,##0.00")' ), "1,234.50" )
   Eq( F( oS, "DATE(2026,11,25)" ), 46351 )
   Eq( F( oS, "YEAR(46351)*10000+MONTH(46351)*100+DAY(46351)" ), 20261125 )
   Eq( F( oS, "ROUND(PMT(0.01,12,1000),2)" ), -88.85 )
   Eq( F( oS, "ROUND(@PMT(1000,0.01,12),2)" ), 88.85 )
   Eq( F( oS, "ROUND(NPV(0.1,100,100),4)" ), 173.5537 )
   Eq( F( oS, "ISERROR(1/0)" ), .T. )
   Eq( F( oS, 'IFERROR(1/0,"x")' ), "x" )
   Eq( F( oS, 'CHOOSE(2,"a","b","c")' ), "b" )
   Eq( F( oS, "NOSUCH(1)" ), "#NAME?" )
   Eq( F( oS, "AND(B1>1,B2>1)" ), .T. )
   Eq( F( oS, 'SUM(B1:B3,10,"5")' ), 30 )

   ? "== moving references (Copy)"
   Eq( DC_AstText( DC_AstShift( DC_Parse( "A1+$B$1+C$2+$D3" ), 1, 2 ) ), "B3+$B$1+D$2+$D5" )
   Eq( DC_AstText( DC_AstShift( DC_Parse( "SUM(A1:A5)" ), 0, -1 ) ), "SUM(#REF!)" )

   ?
   ? "ok:", s_nOk, " failed:", s_nBad
   ?
   ErrorLevel( IIF( s_nBad == 0, 0, 1 ) )

   RETURN

STATIC FUNCTION Txt( c )
   LOCAL cErr, x := DC_Parse( c, @cErr )
RETURN IIF( x == NIL, "ERR: " + cErr, DC_AstText( x ) )

STATIC FUNCTION Lot( c )
   LOCAL cErr, x := DC_Parse( c, @cErr )
RETURN IIF( x == NIL, "ERR: " + cErr, DC_AstText( x, .T. ) )

STATIC FUNCTION ParseErr( c )
   LOCAL cErr := ""
   DC_Parse( c, @cErr )
RETURN cErr

STATIC PROCEDURE Set1( oS, cRef, cText )
   LOCAL a := DC_ParseRef( cRef ), cErr
   IF ! oS:setInput( a[ 1 ], a[ 2 ], cText, @cErr )
      ? "  setInput " + cRef + " <" + cText + ">: " + cErr
   ENDIF
   RETURN

STATIC FUNCTION V( oS, cRef )
   LOCAL a := DC_ParseRef( cRef ), x := oS:value( a[ 1 ], a[ 2 ] )
RETURN IIF( DC_IsErr( x ), x:cCode, x )

/* a formula in a scratch cell (Z99) */
STATIC FUNCTION F( oS, cFormula )
   Set1( oS, "Z99", IIF( Left( cFormula, 1 ) == "@", cFormula, "=" + cFormula ) )
RETURN V( oS, "Z99" )

STATIC PROCEDURE Eq( xGot, xWant )
   LOCAL lOk
   IF HB_ISNUMERIC( xGot ) .AND. HB_ISNUMERIC( xWant )
      lOk := Abs( xGot - xWant ) < 0.000000001
   ELSE
      lOk := ValType( xGot ) == ValType( xWant ) .AND. xGot == xWant
   ENDIF
   IF lOk
      s_nOk++
   ELSE
      s_nBad++
      ? "  FAIL line", ProcLine( 1 ), ": got", hb_ValToExp( xGot ), " want", hb_ValToExp( xWant )
   ENDIF
   RETURN
