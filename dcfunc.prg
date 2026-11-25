/** @file dcfunc.prg
 * DigiCalc - the functions of formulas (Excel names; the 1-2-3 names are
 * mapped in dcformula.prg: @AVG -> AVERAGE, @STRING -> TEXT ...).
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "digicalc.ch"

#define DC_DATE0   0d18991230              /* serial 0 of the Excel 1900 date system */

FUNCTION DC_CallFunc( cName, aArgs, oCtx )

   LOCAL nArgs := Len( aArgs ), a, x, y, z, n, i, aV

   SWITCH cName

   /*---------------------------------------------- aggregates */
   CASE "SUM"
      RETURN Agg( aArgs, oCtx, "SUM" )
   CASE "AVERAGE"
      RETURN Agg( aArgs, oCtx, "AVG" )
   CASE "MIN"
      RETURN Agg( aArgs, oCtx, "MIN" )
   CASE "MAX"
      RETURN Agg( aArgs, oCtx, "MAX" )
   CASE "COUNT"
      RETURN Agg( aArgs, oCtx, "COUNT" )
   CASE "PRODUCT"
      RETURN Agg( aArgs, oCtx, "PROD" )
   CASE "STDEVP"
      RETURN Agg( aArgs, oCtx, "STDP" )
   CASE "STDEV"
      RETURN Agg( aArgs, oCtx, "STD" )
   CASE "VARP"
      RETURN Agg( aArgs, oCtx, "VARP" )
   CASE "VAR"
      RETURN Agg( aArgs, oCtx, "VAR" )
   CASE "COUNTA"
      n := 0
      FOR EACH x IN Flat( aArgs, oCtx )
         IF x != NIL .AND. !( HB_ISSTRING( x ) .AND. Empty( x ) )
            n++
         ENDIF
      NEXT
      RETURN n
   CASE "COUNTBLANK"
      IF nArgs != 1 .OR. !( aArgs[ 1 ][ 1 ] == "A" )
         RETURN DC_Err( "#VALUE!" )
      ENDIF
      n := ( aArgs[ 1 ][ 3 ][ 1 ] - aArgs[ 1 ][ 2 ][ 1 ] + 1 ) * ( aArgs[ 1 ][ 3 ][ 2 ] - aArgs[ 1 ][ 2 ][ 2 ] + 1 )
      FOR EACH x IN oCtx:areaValues( aArgs[ 1 ][ 2 ], aArgs[ 1 ][ 3 ] )
         IF x != NIL .AND. !( HB_ISSTRING( x ) .AND. Empty( x ) )
            n--
         ENDIF
      NEXT
      RETURN n
   CASE "SUMIF"
   CASE "COUNTIF"
   CASE "AVERAGEIF"
      RETURN CondAgg( cName, aArgs, oCtx )

   /*---------------------------------------------- logic */
   CASE "IF"
      IF nArgs < 2 .OR. nArgs > 3
         RETURN DC_Err( "#VALUE!" )
      ENDIF
      x := DC_ToLog( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( x )
         RETURN x
      ENDIF
      IF x
         RETURN IIF( aArgs[ 2 ] == NIL, 0, Arg( aArgs, 2, oCtx ) )
      ENDIF
      RETURN IIF( nArgs < 3, .F., IIF( aArgs[ 3 ] == NIL, 0, Arg( aArgs, 3, oCtx ) ) )
   CASE "IFERROR"
      x := Arg( aArgs, 1, oCtx )
      RETURN IIF( DC_IsErr( x ), Arg( aArgs, 2, oCtx ), x )
   CASE "AND"
   CASE "OR"
      y := ( cName == "AND" )
      FOR EACH x IN Flat( aArgs, oCtx )
         IF DC_IsErr( x )
            RETURN x
         ENDIF
         IF x == NIL .OR. HB_ISSTRING( x )
            LOOP
         ENDIF
         x := DC_ToLog( x )
         y := IIF( cName == "AND", y .AND. x, y .OR. x )
      NEXT
      RETURN y
   CASE "NOT"
      x := DC_ToLog( Arg( aArgs, 1, oCtx ) )
      RETURN IIF( DC_IsErr( x ), x, ! x )
   CASE "TRUE"
      RETURN .T.
   CASE "FALSE"
      RETURN .F.
   CASE "ISERROR"
      RETURN DC_IsErr( Arg( aArgs, 1, oCtx ) )
   CASE "ISNA"
      x := Arg( aArgs, 1, oCtx )
      RETURN DC_IsErr( x ) .AND. x:cCode == "#N/A"
   CASE "ISNUMBER"
      RETURN HB_ISNUMERIC( Arg( aArgs, 1, oCtx ) )
   CASE "ISTEXT"
   CASE "ISSTRING"
      RETURN HB_ISSTRING( Arg( aArgs, 1, oCtx ) )
   CASE "ISBLANK"
      RETURN Arg( aArgs, 1, oCtx ) == NIL
   CASE "NA"
      RETURN DC_Err( "#N/A" )
   CASE "ERROR"
      RETURN DC_Err( "#VALUE!" )
   CASE "CHOOSE"
      n := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( n )
         RETURN n
      ENDIF
      n := Int( n )
      /* Excel CHOOSE(1,a,b) = a; 1-2-3 @CHOOSE(0,a,b) = a */
      RETURN IIF( n >= 1 .AND. n < nArgs, Arg( aArgs, n + 1, oCtx ), DC_Err( "#VALUE!" ) )

   /*---------------------------------------------- math */
   CASE "ABS"
      RETURN Num1( aArgs, oCtx, {| v | Abs( v ) } )
   CASE "INT"
      RETURN Num1( aArgs, oCtx, {| v | Floor( v ) } )
   CASE "SQRT"
      RETURN Num1( aArgs, oCtx, {| v | IIF( v < 0, DC_Err( "#NUM!" ), Sqrt( v ) ) } )
   CASE "EXP"
      RETURN Num1( aArgs, oCtx, {| v | Exp( v ) } )
   CASE "LN"
      RETURN Num1( aArgs, oCtx, {| v | IIF( v <= 0, DC_Err( "#NUM!" ), Log( v ) ) } )
   CASE "LOG10"
      RETURN Num1( aArgs, oCtx, {| v | IIF( v <= 0, DC_Err( "#NUM!" ), Log10( v ) ) } )
   CASE "LOG"
      IF nArgs == 1
         RETURN Num1( aArgs, oCtx, {| v | IIF( v <= 0, DC_Err( "#NUM!" ), Log10( v ) ) } )
      ENDIF
      RETURN Num2( aArgs, oCtx, {| v, b | IIF( v <= 0 .OR. b <= 0 .OR. b == 1, DC_Err( "#NUM!" ), Log( v ) / Log( b ) ) } )
   CASE "SIN"
      RETURN Num1( aArgs, oCtx, {| v | Sin( v ) } )
   CASE "COS"
      RETURN Num1( aArgs, oCtx, {| v | Cos( v ) } )
   CASE "TAN"
      RETURN Num1( aArgs, oCtx, {| v | Tan( v ) } )
   CASE "ATAN"
      RETURN Num1( aArgs, oCtx, {| v | ATan( v ) } )
   CASE "ASIN"
      RETURN Num1( aArgs, oCtx, {| v | IIF( Abs( v ) > 1, DC_Err( "#NUM!" ), ASin( v ) ) } )
   CASE "ACOS"
      RETURN Num1( aArgs, oCtx, {| v | IIF( Abs( v ) > 1, DC_Err( "#NUM!" ), ACos( v ) ) } )
   CASE "PI"
      RETURN 3.14159265358979324
   CASE "RAND"
      RETURN hb_Random()
   CASE "SIGN"
      RETURN Num1( aArgs, oCtx, {| v | IIF( v > 0, 1, IIF( v < 0, -1, 0 ) ) } )
   CASE "ROUND"
      RETURN Num2( aArgs, oCtx, {| v, d | RoundHalf( v, Int( d ) ) }, 0 )
   CASE "ROUNDUP"
      RETURN Num2( aArgs, oCtx, {| v, d | RoundDir( v, Int( d ), .T. ) }, 0 )
   CASE "ROUNDDOWN"
   CASE "TRUNC"
      RETURN Num2( aArgs, oCtx, {| v, d | RoundDir( v, Int( d ), .F. ) }, 0 )
   CASE "MOD"
      RETURN Num2( aArgs, oCtx, {| v, d | IIF( d == 0, DC_Err( "#DIV/0!" ), v - d * Floor( v / d ) ) } )
   CASE "POWER"
      RETURN Num2( aArgs, oCtx, {| v, p | DC_BinOp( "^", v, p ) } )
   CASE "ATAN2"
      RETURN Num2( aArgs, oCtx, {| x1, y1 | IIF( x1 == 0 .AND. y1 == 0, DC_Err( "#DIV/0!" ), ATn2( y1, x1 ) ) } )

   /*---------------------------------------------- text */
   CASE "LEN"
      RETURN Txt1( aArgs, oCtx, {| c | Len( c ) } )
   CASE "UPPER"
      RETURN Txt1( aArgs, oCtx, {| c | Upper( c ) } )
   CASE "LOWER"
      RETURN Txt1( aArgs, oCtx, {| c | Lower( c ) } )
   CASE "PROPER"
      RETURN Txt1( aArgs, oCtx, {| c | Proper( c ) } )
   CASE "TRIM"
      RETURN Txt1( aArgs, oCtx, {| c | TrimInner( c ) } )
   CASE "LEFT"
   CASE "RIGHT"
      x := Arg( aArgs, 1, oCtx )
      n := IIF( nArgs > 1, DC_ToNum( Arg( aArgs, 2, oCtx ) ), 1 )
      IF DC_IsErr( x ) .OR. DC_IsErr( n )
         RETURN IIF( DC_IsErr( x ), x, n )
      ENDIF
      x := DC_ToText( x )
      RETURN IIF( n < 0, DC_Err( "#VALUE!" ), IIF( cName == "LEFT", Left( x, n ), Right( x, n ) ) )
   CASE "MID"
      x := Arg( aArgs, 1, oCtx )
      y := DC_ToNum( Arg( aArgs, 2, oCtx ) )
      n := DC_ToNum( Arg( aArgs, 3, oCtx ) )
      FOR EACH z IN { x, y, n }
         IF DC_IsErr( z )
            RETURN z
         ENDIF
      NEXT
      /* Excel MID starts at 1; 1-2-3 @MID at 0 (an @MID becomes MID(,+1,) on entry? no: kept as Excel) */
      RETURN IIF( y < 1 .OR. n < 0, DC_Err( "#VALUE!" ), SubStr( DC_ToText( x ), y, n ) )
   CASE "CONCATENATE"
   CASE "CONCAT"
      x := ""
      FOR EACH y IN Flat( aArgs, oCtx, .T. )
         IF DC_IsErr( y )
            RETURN y
         ENDIF
         x += DC_ToText( y )
      NEXT
      RETURN x
   CASE "REPT"
      x := Arg( aArgs, 1, oCtx )
      n := DC_ToNum( Arg( aArgs, 2, oCtx ) )
      RETURN IIF( DC_IsErr( x ), x, IIF( DC_IsErr( n ), n, Replicate( DC_ToText( x ), Max( 0, Int( n ) ) ) ) )
   CASE "EXACT"
      x := Arg( aArgs, 1, oCtx )
      y := Arg( aArgs, 2, oCtx )
      RETURN DC_ToText( x ) == DC_ToText( y )
   CASE "FIND"
   CASE "SEARCH"
      x := DC_ToText( Arg( aArgs, 1, oCtx ) )
      y := DC_ToText( Arg( aArgs, 2, oCtx ) )
      n := IIF( nArgs > 2, DC_ToNum( Arg( aArgs, 3, oCtx ) ), 1 )
      IF DC_IsErr( n )
         RETURN n
      ENDIF
      IF cName == "SEARCH"
         x := Upper( x )
         y := Upper( y )
      ENDIF
      i := hb_At( x, y, Max( 1, n ) )
      RETURN IIF( i == 0, DC_Err( "#VALUE!" ), i )
   CASE "SUBSTITUTE"
      x := DC_ToText( Arg( aArgs, 1, oCtx ) )
      RETURN StrTran( x, DC_ToText( Arg( aArgs, 2, oCtx ) ), DC_ToText( Arg( aArgs, 3, oCtx ) ) )
   CASE "CHAR"
      RETURN Num1( aArgs, oCtx, {| v | hb_UChar( Int( v ) ) } )
   CASE "CODE"
      x := DC_ToText( Arg( aArgs, 1, oCtx ) )
      RETURN IIF( Empty( x ), DC_Err( "#VALUE!" ), hb_UCode( Left( x, 1 ) ) )
   CASE "VALUE"
      x := Arg( aArgs, 1, oCtx )
      RETURN IIF( HB_ISNUMERIC( x ), x, DC_ToNum( DC_ToText( x ) ) )
   CASE "N"                                /* 1-2-3 @N: a number or 0 */
      x := FirstCell( aArgs, oCtx )
      RETURN IIF( HB_ISNUMERIC( x ), x, 0 )
   CASE "S"                                /* 1-2-3 @S: a text or "" */
      x := FirstCell( aArgs, oCtx )
      RETURN IIF( HB_ISSTRING( x ), x, "" )
   CASE "TEXT"
      x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( x )
         RETURN x
      ENDIF
      y := IIF( nArgs > 1, Arg( aArgs, 2, oCtx ), 0 )
      /* 1-2-3 @STRING(x,2): decimals; Excel TEXT(x,"0.00"): a format */
      IF HB_ISNUMERIC( y )
         RETURN LTrim( Str( x, 30, Max( 0, Min( 15, Int( y ) ) ) ) )
      ENDIF
      RETURN DC_FormatPicture( x, DC_ToText( y ) )

   /*---------------------------------------------- lookup */
   CASE "VLOOKUP"
   CASE "HLOOKUP"
      RETURN Lookup( cName == "VLOOKUP", aArgs, oCtx )
   CASE "INDEX"
      IF nArgs < 2 .OR. !( aArgs[ 1 ][ 1 ] == "A" .OR. aArgs[ 1 ][ 1 ] == "R" )
         RETURN DC_Err( "#VALUE!" )
      ENDIF
      a := AreaOf( aArgs[ 1 ] )
      y := DC_ToNum( Arg( aArgs, 2, oCtx ) )
      x := IIF( nArgs > 2, DC_ToNum( Arg( aArgs, 3, oCtx ) ), 1 )
      IF DC_IsErr( x ) .OR. DC_IsErr( y )
         RETURN IIF( DC_IsErr( x ), x, y )
      ENDIF
      x := Int( x )
      y := Int( y )
      IF a[ 1 ][ 1 ] == a[ 2 ][ 1 ] .AND. nArgs == 2      /* one column: INDEX(r, n) */
         x := 1
      ELSEIF a[ 1 ][ 2 ] == a[ 2 ][ 2 ] .AND. nArgs == 2  /* one row */
         x := y
         y := 1
      ENDIF
      IF x < 1 .OR. y < 1 .OR. a[ 1 ][ 1 ] + x - 1 > a[ 2 ][ 1 ] .OR. a[ 1 ][ 2 ] + y - 1 > a[ 2 ][ 2 ]
         RETURN DC_Err( "#REF!" )
      ENDIF
      RETURN oCtx:evalCell( a[ 1 ][ 1 ] + x - 1, a[ 1 ][ 2 ] + y - 1 )
   CASE "MATCH"
      IF nArgs < 2 .OR. !( aArgs[ 2 ][ 1 ] == "A" )
         RETURN DC_Err( "#N/A" )
      ENDIF
      x := Arg( aArgs, 1, oCtx )
      n := IIF( nArgs > 2, DC_ToNum( Arg( aArgs, 3, oCtx ) ), 1 )
      aV := {}
      AEval( oCtx:areaMatrix( aArgs[ 2 ][ 2 ], aArgs[ 2 ][ 3 ] ), {| r | AEval( r, {| v | AAdd( aV, v ) } ) } )
      RETURN MatchIn( x, aV, n )
   CASE "ROWS"
   CASE "COLUMNS"
      IF nArgs != 1 .OR. !( aArgs[ 1 ][ 1 ] == "A" .OR. aArgs[ 1 ][ 1 ] == "R" )
         RETURN DC_Err( "#VALUE!" )
      ENDIF
      a := AreaOf( aArgs[ 1 ] )
      RETURN IIF( cName == "ROWS", a[ 2 ][ 2 ] - a[ 1 ][ 2 ] + 1, a[ 2 ][ 1 ] - a[ 1 ][ 1 ] + 1 )

   /*---------------------------------------------- dates (serial numbers, 1900 system) */
   CASE "TODAY"
      RETURN Date() - DC_DATE0
   CASE "NOW"
      RETURN ( Date() - DC_DATE0 ) + Seconds() / 86400
   CASE "DATE"
      x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      y := DC_ToNum( Arg( aArgs, 2, oCtx ) )
      z := DC_ToNum( Arg( aArgs, 3, oCtx ) )
      FOR EACH n IN { x, y, z }
         IF DC_IsErr( n )
            RETURN n
         ENDIF
      NEXT
      RETURN DC_DateSerial( Int( x ), Int( y ), Int( z ) )
   CASE "YEAR"
   CASE "MONTH"
   CASE "DAY"
   CASE "WEEKDAY"
      x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( x )
         RETURN x
      ENDIF
      y := DC_DATE0 + Int( x )
      SWITCH cName
      CASE "YEAR"  ; RETURN Year( y )
      CASE "MONTH" ; RETURN Month( y )
      CASE "DAY"   ; RETURN Day( y )
      ENDSWITCH
      RETURN DoW( y )                       /* 1 = Sunday, as Excel's default */
   CASE "HOUR"
   CASE "MINUTE"
   CASE "SECOND"
      x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( x )
         RETURN x
      ENDIF
      n := Round( ( x - Int( x ) ) * 86400, 0 )
      RETURN IIF( cName == "HOUR", Int( n / 3600 ), IIF( cName == "MINUTE", Int( n % 3600 / 60 ), n % 60 ) )

   /*---------------------------------------------- finance (Excel signs) */
   CASE "PMT"
   CASE "PV"
   CASE "FV"
   CASE "NPER"
      RETURN Finance( cName, aArgs, oCtx )
   CASE "NPV"
      x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
      IF DC_IsErr( x )
         RETURN x
      ENDIF
      y := 0
      n := 0
      FOR EACH z IN Flat( hb_ADel( AClone( aArgs ), 1, .T. ), oCtx )
         IF HB_ISNUMERIC( z )
            n++
            y += z / ( 1 + x ) ^ n
         ELSEIF DC_IsErr( z )
            RETURN z
         ENDIF
      NEXT
      RETURN y

   ENDSWITCH

RETURN DC_Err( "#NAME?" )

/*------------------------------------------------------------ helpers */

STATIC FUNCTION Arg( aArgs, n, oCtx )
RETURN IIF( n > Len( aArgs ) .OR. aArgs[ n ] == NIL, NIL, DC_Eval( aArgs[ n ], oCtx ) )

/* the cell of @N / @S: the first cell of an area too */
STATIC FUNCTION FirstCell( aArgs, oCtx )
   IF Len( aArgs ) >= 1 .AND. aArgs[ 1 ] != NIL .AND. aArgs[ 1 ][ 1 ] == "A"
      RETURN oCtx:evalCell( aArgs[ 1 ][ 2 ][ 1 ], aArgs[ 1 ][ 2 ][ 2 ] )
   ENDIF
RETURN Arg( aArgs, 1, oCtx )

STATIC FUNCTION AreaOf( x )
RETURN IIF( x[ 1 ] == "R", { x[ 2 ], x[ 2 ] }, { x[ 2 ], x[ 3 ] } )

/* every value of the arguments: areas and cells in full (lAll: texts and
   logicals of cells too, as CONCAT wants), literals as they are */
STATIC FUNCTION Flat( aArgs, oCtx, lAll )
   LOCAL aOut := {}, x
   HB_SYMBOL_UNUSED( lAll )
   FOR EACH x IN aArgs
      IF x == NIL
         LOOP
      ENDIF
      IF x[ 1 ] == "A"
         AEval( oCtx:areaValues( x[ 2 ], x[ 3 ] ), {| v | AAdd( aOut, v ) } )
      ELSE
         AAdd( aOut, DC_Eval( x, oCtx ) )
      ENDIF
   NEXT
RETURN aOut

/* SUM AVERAGE MIN MAX COUNT ...: numbers of cells and areas; texts and empty
   cells left out; a literal text that is a number counts (as Excel) */
STATIC FUNCTION Agg( aArgs, oCtx, cKind )

   LOCAL aNum := {}, x, v, nSum := 0, nMean, nVar := 0

   FOR EACH x IN aArgs
      IF x == NIL
         LOOP
      ENDIF
      IF x[ 1 ] == "A" .OR. x[ 1 ] == "R"
         FOR EACH v IN IIF( x[ 1 ] == "A", oCtx:areaValues( x[ 2 ], x[ 3 ] ), { DC_Eval( x, oCtx ) } )
            IF DC_IsErr( v )
               RETURN IIF( cKind == "COUNT", 0, v ) + 0 * 0
            ENDIF
            IF HB_ISNUMERIC( v )
               AAdd( aNum, v )
            ENDIF
         NEXT
      ELSE
         v := DC_Eval( x, oCtx )
         IF DC_IsErr( v )
            IF cKind == "COUNT"
               LOOP
            ENDIF
            RETURN v
         ENDIF
         IF cKind == "COUNT"
            IF HB_ISNUMERIC( v ) .OR. ( HB_ISSTRING( v ) .AND. DC_IsNumText( v ) )
               AAdd( aNum, 0 )
            ENDIF
            LOOP
         ENDIF
         v := DC_ToNum( v )
         IF DC_IsErr( v )
            RETURN v
         ENDIF
         AAdd( aNum, v )
      ENDIF
   NEXT

   SWITCH cKind
   CASE "COUNT"
      RETURN Len( aNum )
   CASE "SUM"
      AEval( aNum, {| n | nSum += n } )
      RETURN nSum
   CASE "PROD"
      nSum := 1
      AEval( aNum, {| n | nSum *= n } )
      RETURN IIF( Empty( aNum ), 0, nSum )
   CASE "MIN"
      RETURN IIF( Empty( aNum ), 0, ASort( aNum )[ 1 ] )
   CASE "MAX"
      RETURN IIF( Empty( aNum ), 0, ATail( ASort( aNum ) ) )
   ENDSWITCH
   IF Empty( aNum )
      RETURN DC_Err( "#DIV/0!" )
   ENDIF
   AEval( aNum, {| n | nSum += n } )
   nMean := nSum / Len( aNum )
   IF cKind == "AVG"
      RETURN nMean
   ENDIF
   AEval( aNum, {| n | nVar += ( n - nMean ) ^ 2 } )
   SWITCH cKind
   CASE "VARP" ; RETURN nVar / Len( aNum )
   CASE "STDP" ; RETURN Sqrt( nVar / Len( aNum ) )
   ENDSWITCH
   IF Len( aNum ) < 2
      RETURN DC_Err( "#DIV/0!" )
   ENDIF
RETURN IIF( cKind == "VAR", nVar / ( Len( aNum ) - 1 ), Sqrt( nVar / ( Len( aNum ) - 1 ) ) )

/* SUMIF(range, criteria [, sum_range])  COUNTIF(range, criteria)  AVERAGEIF */
STATIC FUNCTION CondAgg( cName, aArgs, oCtx )

   LOCAL aR, aS, xCrit, nRow, nCol, v, nSum := 0, nCnt := 0, aSumA

   IF Len( aArgs ) < 2 .OR. aArgs[ 1 ] == NIL .OR. !( aArgs[ 1 ][ 1 ] $ "A|R" )
      RETURN DC_Err( "#VALUE!" )
   ENDIF
   aR := AreaOf( aArgs[ 1 ] )
   xCrit := Arg( aArgs, 2, oCtx )
   aSumA := aR
   IF Len( aArgs ) > 2 .AND. aArgs[ 3 ] != NIL .AND. aArgs[ 3 ][ 1 ] $ "A|R"
      aS := AreaOf( aArgs[ 3 ] )
      aSumA := { aS[ 1 ], { aS[ 1 ][ 1 ] + aR[ 2 ][ 1 ] - aR[ 1 ][ 1 ], aS[ 1 ][ 2 ] + aR[ 2 ][ 2 ] - aR[ 1 ][ 2 ] } }
   ENDIF
   FOR nRow := aR[ 1 ][ 2 ] TO aR[ 2 ][ 2 ]
      FOR nCol := aR[ 1 ][ 1 ] TO aR[ 2 ][ 1 ]
         IF oCtx:cell( nCol, nRow ) == NIL .AND. !( HB_ISSTRING( xCrit ) .AND. Empty( xCrit ) )
            LOOP
         ENDIF
         IF CritMatch( oCtx:evalCell( nCol, nRow ), xCrit )
            nCnt++
            v := oCtx:evalCell( aSumA[ 1 ][ 1 ] + nCol - aR[ 1 ][ 1 ], aSumA[ 1 ][ 2 ] + nRow - aR[ 1 ][ 2 ] )
            IF HB_ISNUMERIC( v )
               nSum += v
            ENDIF
         ENDIF
      NEXT
   NEXT
   SWITCH cName
   CASE "COUNTIF" ; RETURN nCnt
   CASE "SUMIF"   ; RETURN nSum
   ENDSWITCH
RETURN IIF( nCnt == 0, DC_Err( "#DIV/0!" ), nSum / nCnt )

/* ">5", "<>x", "=abc", "ab*", 5, .T. */
STATIC FUNCTION CritMatch( v, xCrit )
   LOCAL cOp := "=", cVal, xVal, n
   IF ! HB_ISSTRING( xCrit )
      RETURN DC_Compare( v, xCrit ) == 0 .AND. ValType( v ) == ValType( xCrit )
   ENDIF
   cVal := xCrit
   FOR EACH n IN { "<>", "<=", ">=", "<", ">", "=" }
      IF Left( cVal, Len( n ) ) == n
         cOp := n
         cVal := SubStr( cVal, Len( n ) + 1 )
         EXIT
      ENDIF
   NEXT
   xVal := IIF( DC_IsNumText( cVal ), DC_Val( cVal ), cVal )
   IF DC_IsErr( v )
      RETURN .F.
   ENDIF
   IF cOp == "=" .OR. cOp == "<>"
      IF HB_ISSTRING( xVal )
         n := HB_ISSTRING( v ) .AND. hb_WildMatchI( xVal, v, .T. )
         IF v == NIL .AND. Empty( xVal )
            n := .T.
         ENDIF
      ELSE
         n := HB_ISNUMERIC( v ) .AND. v == xVal
      ENDIF
      RETURN IIF( cOp == "=", n, ! n )
   ENDIF
   IF ValType( v ) != ValType( xVal )
      RETURN .F.
   ENDIF
   n := DC_Compare( v, xVal )
   SWITCH cOp
   CASE "<"  ; RETURN n < 0
   CASE ">"  ; RETURN n > 0
   CASE "<=" ; RETURN n <= 0
   ENDSWITCH
RETURN n >= 0

STATIC FUNCTION Num1( aArgs, oCtx, bDo )
   LOCAL x
   IF Len( aArgs ) < 1
      RETURN DC_Err( "#VALUE!" )
   ENDIF
   x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
RETURN IIF( DC_IsErr( x ), x, Eval( bDo, x ) )

STATIC FUNCTION Num2( aArgs, oCtx, bDo, xDef2 )
   LOCAL x, y
   IF Len( aArgs ) < IIF( xDef2 == NIL, 2, 1 )
      RETURN DC_Err( "#VALUE!" )
   ENDIF
   x := DC_ToNum( Arg( aArgs, 1, oCtx ) )
   y := IIF( Len( aArgs ) < 2, xDef2, DC_ToNum( Arg( aArgs, 2, oCtx ) ) )
   IF DC_IsErr( x )
      RETURN x
   ENDIF
RETURN IIF( DC_IsErr( y ), y, Eval( bDo, x, y ) )

STATIC FUNCTION Txt1( aArgs, oCtx, bDo )
   LOCAL x
   IF Len( aArgs ) < 1
      RETURN DC_Err( "#VALUE!" )
   ENDIF
   x := Arg( aArgs, 1, oCtx )
RETURN IIF( DC_IsErr( x ), x, Eval( bDo, DC_ToText( x ) ) )

/* Excel ROUND: halves away from zero */
STATIC FUNCTION RoundHalf( v, d )
   LOCAL m := 10 ^ d
RETURN IIF( v < 0, -1, 1 ) * Int( Abs( v ) * m + 0.5 + 0.000000001 ) / m

STATIC FUNCTION RoundDir( v, d, lUp )
   LOCAL m := 10 ^ d, a := Abs( v ) * m
   a := IIF( lUp, IIF( a == Int( a ), a, Int( a ) + 1 ), Int( a ) )
RETURN IIF( v < 0, -1, 1 ) * a / m

STATIC FUNCTION ATn2( y, x )
   DO CASE
   CASE x > 0
      RETURN ATan( y / x )
   CASE x < 0 .AND. y >= 0
      RETURN ATan( y / x ) + 3.14159265358979324
   CASE x < 0
      RETURN ATan( y / x ) - 3.14159265358979324
   CASE y > 0
      RETURN 3.14159265358979324 / 2
   ENDCASE
RETURN -3.14159265358979324 / 2

STATIC FUNCTION Proper( c )
   LOCAL cOut := "", ch, lUp := .T.
   FOR EACH ch IN c
      cOut += IIF( lUp, Upper( ch ), Lower( ch ) )
      lUp := ! ( IsAlpha( ch ) .OR. IsDigit( ch ) )
   NEXT
RETURN cOut

STATIC FUNCTION TrimInner( c )
   c := AllTrim( c )
   DO WHILE "  " $ c
      c := StrTran( c, "  ", " " )
   ENDDO
RETURN c

/* VLOOKUP(x, area, col [, approx]) / HLOOKUP */
STATIC FUNCTION Lookup( lV, aArgs, oCtx )

   LOCAL x, a, n, lApprox := .T., aM, i, nHit, nLen, v

   IF Len( aArgs ) < 3 .OR. aArgs[ 2 ] == NIL .OR. !( aArgs[ 2 ][ 1 ] == "A" )
      RETURN DC_Err( "#N/A" )
   ENDIF
   x := Arg( aArgs, 1, oCtx )
   IF DC_IsErr( x )
      RETURN x
   ENDIF
   a := AreaOf( aArgs[ 2 ] )
   n := DC_ToNum( Arg( aArgs, 3, oCtx ) )
   IF DC_IsErr( n )
      RETURN n
   ENDIF
   n := Int( n )
   IF Len( aArgs ) > 3 .AND. aArgs[ 4 ] != NIL
      lApprox := DC_ToLog( Arg( aArgs, 4, oCtx ) )
      IF DC_IsErr( lApprox )
         RETURN lApprox
      ENDIF
   ENDIF
   nLen := IIF( lV, a[ 2 ][ 2 ] - a[ 1 ][ 2 ] + 1, a[ 2 ][ 1 ] - a[ 1 ][ 1 ] + 1 )
   IF n < 1 .OR. n > IIF( lV, a[ 2 ][ 1 ] - a[ 1 ][ 1 ] + 1, a[ 2 ][ 2 ] - a[ 1 ][ 2 ] + 1 )
      RETURN DC_Err( "#REF!" )
   ENDIF
   aM := {}
   FOR i := 1 TO nLen
      AAdd( aM, IIF( lV, oCtx:evalCell( a[ 1 ][ 1 ], a[ 1 ][ 2 ] + i - 1 ), oCtx:evalCell( a[ 1 ][ 1 ] + i - 1, a[ 1 ][ 2 ] ) ) )
   NEXT
   IF lApprox
      nHit := MatchIn( x, aM, 1 )
   ELSE
      nHit := MatchIn( x, aM, 0 )
   ENDIF
   IF DC_IsErr( nHit )
      RETURN nHit
   ENDIF
   v := IIF( lV, oCtx:evalCell( a[ 1 ][ 1 ] + n - 1, a[ 1 ][ 2 ] + nHit - 1 ), oCtx:evalCell( a[ 1 ][ 1 ] + nHit - 1, a[ 1 ][ 2 ] + n - 1 ) )
RETURN IIF( v == NIL, 0, v )

/* MATCH: 0 exact, 1 the largest <= x (sorted ascending), -1 the smallest >= x */
STATIC FUNCTION MatchIn( x, aV, nType )
   LOCAL i, nBest := 0
   FOR i := 1 TO Len( aV )
      IF aV[ i ] == NIL .OR. DC_IsErr( aV[ i ] )
         LOOP
      ENDIF
      IF nType == 0
         IF ValType( aV[ i ] ) == ValType( x ) .AND. DC_Compare( aV[ i ], x ) == 0
            RETURN i
         ENDIF
      ELSEIF ValType( aV[ i ] ) == ValType( x )
         IF nType > 0
            IF DC_Compare( aV[ i ], x ) <= 0
               nBest := i
            ELSE
               EXIT
            ENDIF
         ELSEIF DC_Compare( aV[ i ], x ) >= 0
            nBest := i
         ELSE
            EXIT
         ENDIF
      ENDIF
   NEXT
RETURN IIF( nBest == 0, DC_Err( "#N/A" ), nBest )

/* PMT(rate,nper,pv[,fv[,type]])  PV(rate,nper,pmt[,fv[,type]])  FV(rate,nper,pmt[,pv[,type]])
   NPER(rate,pmt,pv[,fv[,type]]) */
STATIC FUNCTION Finance( cName, aArgs, oCtx )

   LOCAL r, n, p, f := 0, t := 0, a := {}, x, v, q

   FOR EACH x IN aArgs               /* x is the array element: not assigned */
      v := IIF( x == NIL, 0, DC_ToNum( DC_Eval( x, oCtx ) ) )
      IF DC_IsErr( v )
         RETURN v
      ENDIF
      AAdd( a, v )
   NEXT
   IF Len( a ) < 3
      RETURN DC_Err( "#VALUE!" )
   ENDIF
   r := a[ 1 ]
   n := a[ 2 ]
   p := a[ 3 ]
   IF Len( a ) > 3
      f := a[ 4 ]
   ENDIF
   IF Len( a ) > 4
      t := IIF( a[ 5 ] != 0, 1, 0 )
   ENDIF
   SWITCH cName
   CASE "PMT"                         /* n = nper, p = pv */
      IF r == 0
         RETURN -( p + f ) / n
      ENDIF
      q := ( 1 + r ) ^ n
      RETURN -( p * q + f ) * r / ( ( 1 + r * t ) * ( q - 1 ) )
   CASE "PV"                          /* p = pmt */
      IF r == 0
         RETURN -( f + p * n )
      ENDIF
      q := ( 1 + r ) ^ n
      RETURN -( f + p * ( 1 + r * t ) * ( q - 1 ) / r ) / q
   CASE "FV"                          /* p = pmt, f = pv */
      IF r == 0
         RETURN -( f + p * n )
      ENDIF
      q := ( 1 + r ) ^ n
      RETURN -( f * q + p * ( 1 + r * t ) * ( q - 1 ) / r )
   CASE "NPER"                        /* n = pmt, p = pv */
      IF r == 0
         RETURN IIF( n == 0, DC_Err( "#NUM!" ), -( p + f ) / n )
      ENDIF
      q := ( n * ( 1 + r * t ) - f * r ) / ( n * ( 1 + r * t ) + p * r )
      RETURN IIF( q <= 0, DC_Err( "#NUM!" ), Log( q ) / Log( 1 + r ) )
   ENDSWITCH
RETURN DC_Err( "#VALUE!" )

/* Excel 1900 system; DATE(95,1,1) = 1995 as Excel (years 0..1899 + 1900) */
FUNCTION DC_DateSerial( nY, nM, nD )
   LOCAL d
   IF nY < 1900
      nY += 1900
   ENDIF
   nY += Int( ( nM - 1 ) / 12 ) - IIF( nM < 1 .AND. ( nM - 1 ) % 12 != 0, 1, 0 )
   nM := ( ( nM - 1 ) % 12 + 12 ) % 12 + 1
   d := hb_SToD( StrZero( nY, 4 ) + StrZero( nM, 2 ) + "01" )
   IF Empty( d )
      RETURN DC_Err( "#NUM!" )
   ENDIF
RETURN d + nD - 1 - DC_DATE0

FUNCTION DC_SerialDate( n )
RETURN DC_DATE0 + Int( n )

/* TEXT(x, "0.00") / "#,##0.00" / "0%" : the decimals and grouping of the picture */
FUNCTION DC_FormatPicture( x, cPic )
   LOCAL nDec := 0, lComma := "," $ cPic, lPct := "%" $ cPic, n
   n := At( ".", cPic )
   IF n > 0
      DO WHILE n + nDec + 1 <= Len( cPic ) .AND. SubStr( cPic, n + nDec + 1, 1 ) $ "0#"
         nDec++
      ENDDO
   ENDIF
   IF lPct
      x *= 100
   ENDIF
RETURN DC_FmtNumber( x, nDec, lComma ) + IIF( lPct, "%", "" )

/* 1234567.891, 2, .T. -> "1,234,567.89" */
FUNCTION DC_FmtNumber( n, nDec, lComma )
   LOCAL c := LTrim( Str( Abs( n ), 30, nDec ) ), cInt, cFrac := "", i, cOut := ""
   IF nDec > 0
      cInt := Left( c, At( ".", c ) - 1 )
      cFrac := SubStr( c, At( ".", c ) )
   ELSE
      cInt := c
   ENDIF
   IF lComma
      FOR i := Len( cInt ) TO 1 STEP -1
         cOut := SubStr( cInt, i, 1 ) + cOut
         IF ( Len( cInt ) - i + 1 ) % 3 == 0 .AND. i > 1
            cOut := "," + cOut
         ENDIF
      NEXT
      cInt := cOut
   ENDIF
RETURN IIF( n < 0 .AND. Val( c ) != 0, "-", "" ) + cInt + cFrac
