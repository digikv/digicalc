/** @file dcformula.prg
 * DigiCalc - formulas: addresses, the tokenizer, the parser (Excel and 1-2-3
 * syntax), the formula as text again, and the evaluation.
 *
 *   Excel   =SUM(A1:A5)*2      AND(A1>0,B1<5)     "a"&B1
 *   1-2-3   @SUM(A1..A5)*2     A1>0#AND#B1<5      +A1+B1     (A1)
 *
 * The parse tree (AST) is made of arrays:
 *   { "N", nNum }  { "S", cText }  { "B", lLog }  { "E", cErr }
 *   { "R", aRef }                       one cell; aRef = { nCol, nRow, lAbsCol, lAbsRow }
 *   { "A", aRef1, aRef2 }               an area (range)
 *   { "F", cName, aArgs }               a function (Excel name, upper case)
 *   { "O", cOp, xLeft, xRight }         + - * / ^ & = <> < > <= >= AND OR
 *   { "U", cOp, x }                     unary - + NOT
 *   { "P", x }                          x%
 *   { "()", x }                         parentheses as typed
 *   { "NM", cName }                     a name (a named range, later)
 *
 * Copyright (c) 2026 Srđan Dragojlović
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
 */

#include "hbclass.ch"
#include "digicalc.ch"

/*------------------------------------------------------------ errors */

CREATE CLASS DCErr
   VAR cCode INIT "#VALUE!"
   METHOD new( cCode )
ENDCLASS

METHOD new( cCode ) CLASS DCErr
   ::cCode := hb_defaultValue( cCode, "#VALUE!" )
RETURN Self

FUNCTION DC_Err( cCode )
RETURN DCErr():new( cCode )

FUNCTION DC_IsErr( x )
RETURN HB_ISOBJECT( x ) .AND. x:className() == "DCERR"

/*------------------------------------------------------------ addresses */

/* 1 -> A, 27 -> AA, 16384 -> XFD */
FUNCTION DC_ColName( n )
   LOCAL c := ""
   DO WHILE n > 0
      c := Chr( 65 + ( n - 1 ) % 26 ) + c
      n := Int( ( n - 1 ) / 26 )
   ENDDO
RETURN c

/* A -> 1, XFD -> 16384, 0 when not letters */
FUNCTION DC_ColNum( c )
   LOCAL n := 0, ch
   FOR EACH ch IN Upper( c )
      IF ch < "A" .OR. ch > "Z"
         RETURN 0
      ENDIF
      n := n * 26 + Asc( ch ) - 64
   NEXT
RETURN n

FUNCTION DC_CellName( nCol, nRow, lAbsCol, lAbsRow )
RETURN IIF( hb_defaultValue( lAbsCol, .F. ), "$", "" ) + DC_ColName( nCol ) + ;
       IIF( hb_defaultValue( lAbsRow, .F. ), "$", "" ) + hb_ntos( nRow )

/* "$B$12" -> { 2, 12, .T., .T. }, NIL when it is not a cell address */
FUNCTION DC_ParseRef( cRef )
   LOCAL i := 1, n := Len( cRef ), lAbsC := .F., lAbsR := .F., cCol := "", cRow := ""
   IF i <= n .AND. SubStr( cRef, i, 1 ) == "$"
      lAbsC := .T.
      i++
   ENDIF
   DO WHILE i <= n .AND. IsAlpha( SubStr( cRef, i, 1 ) )
      cCol += SubStr( cRef, i, 1 )
      i++
   ENDDO
   IF i <= n .AND. SubStr( cRef, i, 1 ) == "$"
      lAbsR := .T.
      i++
   ENDIF
   DO WHILE i <= n .AND. IsDigit( SubStr( cRef, i, 1 ) )
      cRow += SubStr( cRef, i, 1 )
      i++
   ENDDO
   IF i <= n .OR. Empty( cCol ) .OR. Empty( cRow ) .OR. Len( cCol ) > 3 .OR. ;
      DC_ColNum( cCol ) > DC_MAXCOL .OR. Val( cRow ) < 1 .OR. Val( cRow ) > DC_MAXROW
      RETURN NIL
   ENDIF
RETURN { DC_ColNum( cCol ), Int( Val( cRow ) ), lAbsC, lAbsR }

FUNCTION DC_Key( nCol, nRow )
RETURN nRow * DC_KEYMUL + nCol

FUNCTION DC_KeyCol( nKey )
RETURN nKey % DC_KEYMUL

FUNCTION DC_KeyRow( nKey )
RETURN Int( nKey / DC_KEYMUL )

/*------------------------------------------------------------ tokenizer */

/* tokens: { cType, xValue }
   NUM STR REF FUNC NAME OP LP RP SEP RANGE ERR */
FUNCTION DC_Tokens( cExpr, cErr )

   LOCAL aTok := {}, i := 1, n := Len( cExpr ), c, c2, cWord, j, cPrev, lAt := .F.

   cErr := ""
   DO WHILE i <= n
      c := SubStr( cExpr, i, 1 )
      cPrev := IIF( Len( aTok ) > 0, ATail( aTok )[ 1 ], "" )
      DO CASE
      CASE c == " " .OR. c == Chr( 9 )
         i++
      CASE IsDigit( c ) .OR. ( c == "." .AND. IsDigit( SubStr( cExpr, i + 1, 1 ) ) .AND. !( cPrev == "REF" ) )
         j := i
         DO WHILE j <= n .AND. ( IsDigit( SubStr( cExpr, j, 1 ) ) .OR. ;
               ( SubStr( cExpr, j, 1 ) == "." .AND. !( SubStr( cExpr, j + 1, 1 ) == "." ) ) )
            j++
         ENDDO
         IF j <= n .AND. Upper( SubStr( cExpr, j, 1 ) ) == "E" .AND. ;
            ( IsDigit( SubStr( cExpr, j + 1, 1 ) ) .OR. ( SubStr( cExpr, j + 1, 1 ) $ "+-" .AND. IsDigit( SubStr( cExpr, j + 2, 1 ) ) ) )
            j += 2
            DO WHILE j <= n .AND. IsDigit( SubStr( cExpr, j, 1 ) )
               j++
            ENDDO
         ENDIF
         AAdd( aTok, { "NUM", DC_Val( SubStr( cExpr, i, j - i ) ) } )
         i := j
      CASE c == '"'
         cWord := ""
         j := i + 1
         DO WHILE .T.
            IF j > n
               cErr := "Missing closing quote"
               RETURN NIL
            ENDIF
            IF SubStr( cExpr, j, 1 ) == '"'
               IF SubStr( cExpr, j + 1, 1 ) == '"'
                  cWord += '"'
                  j += 2
                  LOOP
               ENDIF
               EXIT
            ENDIF
            cWord += SubStr( cExpr, j, 1 )
            j++
         ENDDO
         AAdd( aTok, { "STR", cWord } )
         i := j + 1
      CASE c == "@"                      /* 1-2-3: @SUM - the function follows */
         lAt := .T.
         i++
         LOOP
      CASE c == "#"                      /* 1-2-3 #AND# #OR# #NOT#, or an error value */
         j := hb_At( "#", cExpr, i + 1 )
         cWord := Upper( SubStr( cExpr, i, IIF( j > 0, j - i + 1, 0 ) ) )
         IF cWord $ "#AND#|#OR#|#NOT#" .AND. j > 0
            AAdd( aTok, { "OP", StrTran( cWord, "#", "" ) } )
            i := j + 1
         ELSE
            cWord := DC_ErrLiteral( SubStr( cExpr, i ) )
            IF Empty( cWord )
               cErr := "Unknown #"
               RETURN NIL
            ENDIF
            AAdd( aTok, { "ERR", cWord } )
            i += Len( cWord )
         ENDIF
      CASE c == "$" .OR. IsAlpha( c ) .OR. c == "_"
         j := i
         DO WHILE j <= n .AND. ( IsAlpha( SubStr( cExpr, j, 1 ) ) .OR. IsDigit( SubStr( cExpr, j, 1 ) ) .OR. ;
               SubStr( cExpr, j, 1 ) $ "$_" )
            j++
         ENDDO
         cWord := SubStr( cExpr, i, j - i )
         i := j
         DO WHILE i <= n .AND. SubStr( cExpr, i, 1 ) == " "
            i++
         ENDDO
         DO CASE
         CASE i <= n .AND. SubStr( cExpr, i, 1 ) == "(" .AND. !( "$" $ cWord )
            AAdd( aTok, { "FUNC", DC_FuncAlias( Upper( cWord ) ), lAt } )
         CASE lAt                     /* 1-2-3 functions without arguments: @PI @NOW @TODAY @NA @ERR @RAND */
            AAdd( aTok, { "FUNC0", DC_FuncAlias( Upper( cWord ) ), .T. } )
         CASE ! HB_ISNIL( DC_ParseRef( cWord ) )
            AAdd( aTok, { "REF", DC_ParseRef( cWord ) } )
         CASE Upper( cWord ) == "TRUE" .OR. Upper( cWord ) == "FALSE"
            AAdd( aTok, { "BOOL", Upper( cWord ) == "TRUE" } )
         OTHERWISE
            AAdd( aTok, { "NAME", Upper( cWord ) } )
         ENDCASE
      CASE c == "." .AND. cPrev == "REF"   /* 1-2-3 range: A1..B5 or A1.B5 */
         i++
         IF SubStr( cExpr, i, 1 ) == "."
            i++
         ENDIF
         AAdd( aTok, { "RANGE", ":" } )
      CASE c == ":"
         AAdd( aTok, { "RANGE", ":" } )
         i++
      CASE c == "("
         AAdd( aTok, { "LP", c } )
         i++
      CASE c == ")"
         AAdd( aTok, { "RP", c } )
         i++
      CASE c == "," .OR. c == ";"
         AAdd( aTok, { "SEP", c } )
         i++
      CASE c $ "+-*/^&%"
         AAdd( aTok, { "OP", c } )
         i++
      CASE c $ "<>="
         c2 := SubStr( cExpr, i, 2 )
         IF c2 == "<>" .OR. c2 == "<=" .OR. c2 == ">="
            AAdd( aTok, { "OP", c2 } )
            i += 2
         ELSE
            AAdd( aTok, { "OP", c } )
            i++
         ENDIF
      OTHERWISE
         cErr := "Unexpected '" + c + "'"
         RETURN NIL
      ENDCASE
      lAt := .F.
   ENDDO

RETURN aTok

STATIC FUNCTION DC_ErrLiteral( c )
   LOCAL cE
   FOR EACH cE IN { "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#CIRC!" }
      IF Upper( Left( c, Len( cE ) ) ) == cE
         RETURN cE
      ENDIF
   NEXT
RETURN ""

/* 1-2-3 function names that Excel calls otherwise */
FUNCTION DC_FuncAlias( cName )
   SWITCH cName
   CASE "AVG"     ; RETURN "AVERAGE"
   CASE "STRING"  ; RETURN "TEXT"
   CASE "ISERR"   ; RETURN "ISERROR"
   CASE "ERR"     ; RETURN "ERROR"
   CASE "STD"     ; RETURN "STDEVP"
   CASE "VAR"     ; RETURN "VARP"
   CASE "CTERM"   ; RETURN "NPER"
   CASE "REPEAT"  ; RETURN "REPT"
   CASE "PROPER"  ; RETURN "PROPER"
   CASE "LENGTH"  ; RETURN "LEN"
   ENDSWITCH
RETURN cName

/* Excel name -> 1-2-3 name, for showing a formula the 1-2-3 way */
STATIC FUNCTION DC_FuncLotus( cName )
   SWITCH cName
   CASE "AVERAGE" ; RETURN "AVG"
   CASE "TEXT"    ; RETURN "STRING"
   CASE "ISERROR" ; RETURN "ISERR"
   CASE "ERROR"   ; RETURN "ERR"
   CASE "STDEVP"  ; RETURN "STD"
   CASE "VARP"    ; RETURN "VAR"
   CASE "REPT"    ; RETURN "REPEAT"
   ENDSWITCH
RETURN cName

/*------------------------------------------------------------ parser */

/* text of a formula (without = / with the 1-2-3 start) -> AST, NIL on error (cErr) */
FUNCTION DC_Parse( cExpr, cErr )

   LOCAL aTok := DC_Tokens( cExpr, @cErr ), aP, xAst

   IF HB_ISNIL( aTok )
      RETURN NIL
   ENDIF
   IF Empty( aTok )
      cErr := "Empty formula"
      RETURN NIL
   ENDIF
   aP := { aTok, 1, "" }               /* tokens, position, error */
   xAst := P_Or( aP )
   IF Empty( aP[ 3 ] ) .AND. aP[ 2 ] <= Len( aTok )
      aP[ 3 ] := "Unexpected " + P_Show( aTok[ aP[ 2 ] ] )
   ENDIF
   IF ! Empty( aP[ 3 ] )
      cErr := aP[ 3 ]
      RETURN NIL
   ENDIF
RETURN xAst

STATIC FUNCTION P_Show( aT )
RETURN IIF( aT[ 1 ] == "REF", DC_CellName( aT[ 2 ][ 1 ], aT[ 2 ][ 2 ] ), hb_CStr( aT[ 2 ] ) )

STATIC FUNCTION P_Peek( aP )
RETURN IIF( aP[ 2 ] <= Len( aP[ 1 ] ), aP[ 1 ][ aP[ 2 ] ], { "END", "" } )

STATIC FUNCTION P_IsOp( aP, xOps )
   LOCAL aT := P_Peek( aP )
RETURN aT[ 1 ] == "OP" .AND. AScan( xOps, aT[ 2 ] ) > 0

STATIC FUNCTION P_Or( aP )
   LOCAL x := P_And( aP )
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "OR" } )
      aP[ 2 ]++
      x := { "O", "OR", x, P_And( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_And( aP )
   LOCAL x := P_Not( aP )
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "AND" } )
      aP[ 2 ]++
      x := { "O", "AND", x, P_Not( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Not( aP )
   IF P_IsOp( aP, { "NOT" } )
      aP[ 2 ]++
      RETURN { "U", "NOT", P_Not( aP ) }
   ENDIF
RETURN P_Cmp( aP )

STATIC FUNCTION P_Cmp( aP )
   LOCAL x := P_Concat( aP ), cOp
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "=", "<>", "<", ">", "<=", ">=" } )
      cOp := P_Peek( aP )[ 2 ]
      aP[ 2 ]++
      x := { "O", cOp, x, P_Concat( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Concat( aP )
   LOCAL x := P_Add( aP )
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "&" } )
      aP[ 2 ]++
      x := { "O", "&", x, P_Add( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Add( aP )
   LOCAL x := P_Mul( aP ), cOp
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "+", "-" } )
      cOp := P_Peek( aP )[ 2 ]
      aP[ 2 ]++
      x := { "O", cOp, x, P_Mul( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Mul( aP )
   LOCAL x := P_Pow( aP ), cOp
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "*", "/" } )
      cOp := P_Peek( aP )[ 2 ]
      aP[ 2 ]++
      x := { "O", cOp, x, P_Pow( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Pow( aP )
   LOCAL x := P_Unary( aP )
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "^" } )
      aP[ 2 ]++
      x := { "O", "^", x, P_Unary( aP ) }
   ENDDO
RETURN x

STATIC FUNCTION P_Unary( aP )
   LOCAL cOp
   IF P_IsOp( aP, { "-", "+" } )
      cOp := P_Peek( aP )[ 2 ]
      aP[ 2 ]++
      RETURN { "U", cOp, P_Unary( aP ) }
   ENDIF
RETURN P_Postfix( aP )

STATIC FUNCTION P_Postfix( aP )
   LOCAL x := P_Primary( aP )
   DO WHILE Empty( aP[ 3 ] ) .AND. P_IsOp( aP, { "%" } )
      aP[ 2 ]++
      x := { "P", x }
   ENDDO
RETURN x

STATIC FUNCTION P_Primary( aP )

   LOCAL aT := P_Peek( aP ), x, aArgs, aT2

   IF ! Empty( aP[ 3 ] )
      RETURN NIL
   ENDIF
   aP[ 2 ]++
   SWITCH aT[ 1 ]
   CASE "NUM"
      RETURN { "N", aT[ 2 ] }
   CASE "STR"
      RETURN { "S", aT[ 2 ] }
   CASE "BOOL"
      RETURN { "B", aT[ 2 ] }
   CASE "ERR"
      RETURN { "E", aT[ 2 ] }
   CASE "NAME"
      RETURN { "NM", aT[ 2 ] }
   CASE "FUNC0"
      RETURN { "F", aT[ 2 ], {} }
   CASE "REF"
      IF P_Peek( aP )[ 1 ] == "RANGE"
         aP[ 2 ]++
         aT2 := P_Peek( aP )
         IF !( aT2[ 1 ] == "REF" )
            aP[ 3 ] := "A cell address expected after the range"
            RETURN NIL
         ENDIF
         aP[ 2 ]++
         RETURN DC_Area( aT[ 2 ], aT2[ 2 ] )
      ENDIF
      RETURN { "R", aT[ 2 ] }
   CASE "LP"
      x := P_Or( aP )
      IF !( P_Peek( aP )[ 1 ] == "RP" )
         IF Empty( aP[ 3 ] )
            aP[ 3 ] := "Missing )"
         ENDIF
         RETURN NIL
      ENDIF
      aP[ 2 ]++
      RETURN { "()", x }
   CASE "FUNC"
      aP[ 2 ]++                         /* the ( */
      aArgs := {}
      IF P_Peek( aP )[ 1 ] == "RP"
         aP[ 2 ]++
         RETURN { "F", aT[ 2 ], aArgs }
      ENDIF
      DO WHILE .T.
         IF P_Peek( aP )[ 1 ] == "SEP" .OR. P_Peek( aP )[ 1 ] == "RP"
            AAdd( aArgs, NIL )          /* an empty argument: IF(A1,,2) */
         ELSE
            AAdd( aArgs, P_Or( aP ) )
         ENDIF
         IF ! Empty( aP[ 3 ] )
            RETURN NIL
         ENDIF
         IF P_Peek( aP )[ 1 ] == "SEP"
            aP[ 2 ]++
            LOOP
         ENDIF
         IF P_Peek( aP )[ 1 ] == "RP"
            aP[ 2 ]++
            EXIT
         ENDIF
         aP[ 3 ] := "Missing ) after the arguments of " + aT[ 2 ]
         RETURN NIL
      ENDDO
      IF aT[ 3 ]
         RETURN DC_LotusFunc( aT[ 2 ], aArgs )
      ENDIF
      RETURN { "F", aT[ 2 ], aArgs }
   ENDSWITCH
   aP[ 3 ] := IIF( aT[ 1 ] == "END", "Incomplete formula", "Unexpected " + P_Show( aT ) )

RETURN NIL

/* 1-2-3 functions whose arguments differ from Excel's: the Excel form
   @VLOOKUP(x,r,off)  -> VLOOKUP(x,r,off+1)      offsets count from 0 in 1-2-3
   @HLOOKUP(x,r,off)  -> HLOOKUP(x,r,off+1)
   @INDEX(r,col,row)  -> INDEX(r,row+1,col+1)
   @PMT(p,i,n)        -> -PMT(i,n,p)             Excel: a payment is negative
   @PV(pmt,i,n)       -> -PV(i,n,pmt)
   @FV(pmt,i,n)       -> -FV(i,n,pmt)
   @MOD @ROUND @IF ... are the same */
STATIC FUNCTION DC_LotusFunc( cName, aA )
   LOCAL bPlus1 := {| x | IIF( ! HB_ISNIL( x ) .AND. x[ 1 ] == "N", { "N", x[ 2 ] + 1 }, { "O", "+", x, { "N", 1 } } ) }
   SWITCH cName
   CASE "VLOOKUP"
   CASE "HLOOKUP"
      IF Len( aA ) == 3
         RETURN { "F", cName, { aA[ 1 ], aA[ 2 ], Eval( bPlus1, aA[ 3 ] ), { "B", .T. } } }
      ENDIF
      EXIT
   CASE "INDEX"
      IF Len( aA ) == 3
         RETURN { "F", cName, { aA[ 1 ], Eval( bPlus1, aA[ 3 ] ), Eval( bPlus1, aA[ 2 ] ) } }
      ENDIF
      EXIT
   CASE "PMT"
   CASE "PV"
   CASE "FV"
      IF Len( aA ) == 3
         RETURN { "U", "-", { "F", cName, { aA[ 2 ], aA[ 3 ], aA[ 1 ] } } }
      ENDIF
      EXIT
   ENDSWITCH
RETURN { "F", cName, aA }

/* an area with the corners in order (B5:A1 -> A1:B5) */
FUNCTION DC_Area( aR1, aR2 )
   LOCAL a1 := AClone( aR1 ), a2 := AClone( aR2 ), x
   IF a1[ 1 ] > a2[ 1 ]
      x := a1[ 1 ] ; a1[ 1 ] := a2[ 1 ] ; a2[ 1 ] := x
      x := a1[ 3 ] ; a1[ 3 ] := a2[ 3 ] ; a2[ 3 ] := x
   ENDIF
   IF a1[ 2 ] > a2[ 2 ]
      x := a1[ 2 ] ; a1[ 2 ] := a2[ 2 ] ; a2[ 2 ] := x
      x := a1[ 4 ] ; a1[ 4 ] := a2[ 4 ] ; a2[ 4 ] := x
   ENDIF
RETURN { "A", a1, a2 }

/*------------------------------------------------------------ AST -> text */

/* lLotus: @SUM(A1..A5), #AND#; otherwise Excel (without the leading =) */
FUNCTION DC_AstText( x, lLotus )

   LOCAL c, a, cName

   lLotus := hb_defaultValue( lLotus, .F. )
   IF HB_ISNIL( x )
      RETURN ""
   ENDIF
   SWITCH x[ 1 ]
   CASE "N"
      RETURN DC_NumText( x[ 2 ] )
   CASE "S"
      RETURN '"' + StrTran( x[ 2 ], '"', '""' ) + '"'
   CASE "B"
      RETURN IIF( x[ 2 ], "TRUE", "FALSE" )
   CASE "E"
      RETURN x[ 2 ]
   CASE "NM"
      RETURN x[ 2 ]
   CASE "R"
      RETURN DC_RefText( x[ 2 ] )
   CASE "A"
      RETURN DC_RefText( x[ 2 ] ) + IIF( lLotus, "..", ":" ) + DC_RefText( x[ 3 ] )
   CASE "()"
      RETURN "(" + DC_AstText( x[ 2 ], lLotus ) + ")"
   CASE "P"
      RETURN DC_AstText( x[ 2 ], lLotus ) + "%"
   CASE "U"
      IF x[ 2 ] == "NOT"
         RETURN IIF( lLotus, "#NOT#", "NOT " ) + DC_AstText( x[ 3 ], lLotus )
      ENDIF
      RETURN x[ 2 ] + DC_AstText( x[ 3 ], lLotus )
   CASE "O"
      IF x[ 2 ] == "AND" .OR. x[ 2 ] == "OR"
         IF lLotus
            RETURN DC_AstText( x[ 3 ], lLotus ) + "#" + x[ 2 ] + "#" + DC_AstText( x[ 4 ], lLotus )
         ENDIF
         RETURN x[ 2 ] + "(" + DC_AstText( x[ 3 ], lLotus ) + "," + DC_AstText( x[ 4 ], lLotus ) + ")"
      ENDIF
      RETURN DC_AstText( x[ 3 ], lLotus ) + x[ 2 ] + DC_AstText( x[ 4 ], lLotus )
   CASE "F"
      /* 1-2-3 has no @AND/@OR/@NOT: the operators */
      IF lLotus .AND. ( x[ 2 ] == "AND" .OR. x[ 2 ] == "OR" ) .AND. Len( x[ 3 ] ) >= 2
         c := ""
         FOR EACH a IN x[ 3 ]
            c += IIF( a:__enumIndex() > 1, "#" + x[ 2 ] + "#", "" ) + DC_AstText( a, lLotus )
         NEXT
         RETURN c
      ENDIF
      IF lLotus .AND. x[ 2 ] == "NOT" .AND. Len( x[ 3 ] ) == 1
         RETURN "#NOT#" + DC_AstText( x[ 3 ][ 1 ], lLotus )
      ENDIF
      cName := IIF( lLotus, "@" + DC_FuncLotus( x[ 2 ] ), x[ 2 ] )
      c := ""
      FOR EACH a IN x[ 3 ]
         c += IIF( a:__enumIndex() > 1, ",", "" ) + DC_AstText( a, lLotus )
      NEXT
      RETURN cName + "(" + c + ")"
   ENDSWITCH
RETURN ""

FUNCTION DC_RefText( aRef )
RETURN DC_CellName( aRef[ 1 ], aRef[ 2 ], aRef[ 3 ], aRef[ 4 ] )

/* a number as it is written in a formula: no thousands, no trailing zeros */
FUNCTION DC_NumText( n )
   LOCAL c
   IF n == Int( n ) .AND. Abs( n ) < 1000000000000000
      RETURN hb_ntos( Int( n ) )
   ENDIF
   c := LTrim( Str( n, 30, 10 ) )
   IF "." $ c
      DO WHILE Right( c, 1 ) == "0"
         c := hb_StrShrink( c, 1 )
      ENDDO
      IF Right( c, 1 ) == "."
         c := hb_StrShrink( c, 1 )
      ENDIF
   ENDIF
RETURN c

/*------------------------------------------------------------ moving references */

/* a copy of the AST with the relative references moved by nDCol / nDRow
   (Copy); a reference that falls off the sheet becomes #REF! */
FUNCTION DC_AstShift( x, nDCol, nDRow )

   LOCAL y, i

   IF HB_ISNIL( x )
      RETURN NIL
   ENDIF
   SWITCH x[ 1 ]
   CASE "R"
      y := RefShift( x[ 2 ], nDCol, nDRow )
      RETURN IIF( HB_ISNIL( y ), { "E", "#REF!" }, { "R", y } )
   CASE "A"
      y := { RefShift( x[ 2 ], nDCol, nDRow ), RefShift( x[ 3 ], nDCol, nDRow ) }
      RETURN IIF( HB_ISNIL( y[ 1 ] ) .OR. HB_ISNIL( y[ 2 ] ), { "E", "#REF!" }, { "A", y[ 1 ], y[ 2 ] } )
   ENDSWITCH
   y := AClone( x )
   FOR i := 2 TO Len( y )
      IF HB_ISARRAY( y[ i ] )
         IF x[ 1 ] == "F" .AND. i == 3
            y[ i ] := AClone( x[ i ] )
            AEval( y[ i ], {| a, n | y[ 3 ][ n ] := DC_AstShift( a, nDCol, nDRow ) } )
         ELSE
            y[ i ] := DC_AstShift( x[ i ], nDCol, nDRow )
         ENDIF
      ENDIF
   NEXT
RETURN y

STATIC FUNCTION RefShift( aRef, nDCol, nDRow )
   LOCAL a := AClone( aRef )
   IF ! a[ 3 ]
      a[ 1 ] += nDCol
   ENDIF
   IF ! a[ 4 ]
      a[ 2 ] += nDRow
   ENDIF
   IF a[ 1 ] < 1 .OR. a[ 2 ] < 1 .OR. a[ 1 ] > DC_MAXCOL .OR. a[ 2 ] > DC_MAXROW
      RETURN NIL
   ENDIF
RETURN a

/* every reference in the AST: { { "R", aRef } | { "A", a1, a2 } | { "NM", c } ... } */
FUNCTION DC_AstRefs( x, aOut )
   LOCAL i
   aOut := IIF( HB_ISNIL( aOut ), {}, aOut )
   IF HB_ISNIL( x )
      RETURN aOut
   ENDIF
   SWITCH x[ 1 ]
   CASE "R"
   CASE "A"
   CASE "NM"
      AAdd( aOut, x )
      RETURN aOut
   CASE "F"
      AEval( x[ 3 ], {| a | DC_AstRefs( a, aOut ) } )
      RETURN aOut
   ENDSWITCH
   FOR i := 2 TO Len( x )
      IF HB_ISARRAY( x[ i ] )
         DC_AstRefs( x[ i ], aOut )
      ENDIF
   NEXT
RETURN aOut

/*------------------------------------------------------------ evaluation */

/* oCtx: the sheet; it answers :evalCell( nCol, nRow ) and :areaValues( a1, a2 ) */
FUNCTION DC_Eval( x, oCtx )

   LOCAL a, b

   IF HB_ISNIL( x )
      RETURN NIL
   ENDIF
   SWITCH x[ 1 ]
   CASE "N"
   CASE "S"
   CASE "B"
      RETURN x[ 2 ]
   CASE "E"
      RETURN DC_Err( x[ 2 ] )
   CASE "R"
      RETURN oCtx:evalCell( x[ 2 ][ 1 ], x[ 2 ][ 2 ] )
   CASE "A"
      /* an area where one value is wanted: the cell in the same row / column
         would be Excel's implicit intersection; keep it simple: the first cell */
      IF x[ 2 ][ 1 ] == x[ 3 ][ 1 ] .AND. x[ 2 ][ 2 ] == x[ 3 ][ 2 ]
         RETURN oCtx:evalCell( x[ 2 ][ 1 ], x[ 2 ][ 2 ] )
      ENDIF
      RETURN DC_Err( "#VALUE!" )
   CASE "NM"
      RETURN oCtx:evalName( x[ 2 ] )
   CASE "()"
      RETURN DC_Eval( x[ 2 ], oCtx )
   CASE "P"
      a := DC_ToNum( DC_Eval( x[ 2 ], oCtx ) )
      RETURN IIF( DC_IsErr( a ), a, a / 100 )
   CASE "U"
      a := DC_Eval( x[ 3 ], oCtx )
      IF x[ 2 ] == "NOT"
         a := DC_ToLog( a )
         RETURN IIF( DC_IsErr( a ), a, ! a )
      ENDIF
      a := DC_ToNum( a )
      RETURN IIF( DC_IsErr( a ), a, IIF( x[ 2 ] == "-", -a, a ) )
   CASE "O"
      a := DC_Eval( x[ 3 ], oCtx )
      b := DC_Eval( x[ 4 ], oCtx )
      RETURN DC_BinOp( x[ 2 ], a, b )
   CASE "F"
      RETURN DC_CallFunc( x[ 2 ], x[ 3 ], oCtx )
   ENDSWITCH
RETURN DC_Err( "#VALUE!" )

FUNCTION DC_BinOp( cOp, a, b )

   LOCAL n

   IF DC_IsErr( a )
      RETURN a
   ENDIF
   IF DC_IsErr( b )
      RETURN b
   ENDIF
   SWITCH cOp
   CASE "&"
      RETURN DC_ToText( a ) + DC_ToText( b )
   CASE "AND"
      a := DC_ToLog( a ) ; b := DC_ToLog( b )
      RETURN IIF( DC_IsErr( a ), a, IIF( DC_IsErr( b ), b, a .AND. b ) )
   CASE "OR"
      a := DC_ToLog( a ) ; b := DC_ToLog( b )
      RETURN IIF( DC_IsErr( a ), a, IIF( DC_IsErr( b ), b, a .OR. b ) )
   CASE "="
   CASE "<>"
   CASE "<"
   CASE ">"
   CASE "<="
   CASE ">="
      n := DC_Compare( a, b )
      SWITCH cOp
      CASE "="  ; RETURN n == 0
      CASE "<>" ; RETURN n != 0
      CASE "<"  ; RETURN n < 0
      CASE ">"  ; RETURN n > 0
      CASE "<=" ; RETURN n <= 0
      ENDSWITCH
      RETURN n >= 0
   ENDSWITCH

   a := DC_ToNum( a )
   b := DC_ToNum( b )
   IF DC_IsErr( a )
      RETURN a
   ENDIF
   IF DC_IsErr( b )
      RETURN b
   ENDIF
   SWITCH cOp
   CASE "+" ; RETURN a + b
   CASE "-" ; RETURN a - b
   CASE "*" ; RETURN a * b
   CASE "/"
      RETURN IIF( b == 0, DC_Err( "#DIV/0!" ), a / b )
   CASE "^"
      IF a == 0 .AND. b < 0
         RETURN DC_Err( "#DIV/0!" )
      ENDIF
      IF a < 0 .AND. b != Int( b )
         RETURN DC_Err( "#NUM!" )
      ENDIF
      RETURN a ^ b
   ENDSWITCH
RETURN DC_Err( "#VALUE!" )

/* Excel order: numbers < texts < logicals; texts without case; empty = 0 / "" */
FUNCTION DC_Compare( a, b )
   LOCAL nA := DC_TypeRank( a, b ), nB := DC_TypeRank( b, a )
   IF nA != nB
      RETURN IIF( nA < nB, -1, 1 )
   ENDIF
   IF HB_ISNIL( a )
      a := IIF( HB_ISSTRING( b ), "", IIF( HB_ISLOGICAL( b ), .F., 0 ) )
   ENDIF
   IF HB_ISNIL( b )
      b := IIF( HB_ISSTRING( a ), "", IIF( HB_ISLOGICAL( a ), .F., 0 ) )
   ENDIF
   IF HB_ISSTRING( a )
      a := Upper( a )
      b := Upper( b )
   ELSEIF HB_ISLOGICAL( a )
      a := IIF( a, 1, 0 )
      b := IIF( b, 1, 0 )
   ENDIF
RETURN IIF( a == b, 0, IIF( a < b, -1, 1 ) )

STATIC FUNCTION DC_TypeRank( x, xOther )
   IF HB_ISNIL( x )
      RETURN DC_TypeRank( xOther, 0 )
   ENDIF
RETURN IIF( HB_ISNUMERIC( x ), 1, IIF( HB_ISSTRING( x ), 2, 3 ) )

/* coercions: empty -> 0 / "" / .F.; a text that is a number -> the number */
FUNCTION DC_ToNum( x )
   DO CASE
   CASE HB_ISNIL( x )
      RETURN 0
   CASE HB_ISNUMERIC( x )
      RETURN x
   CASE HB_ISLOGICAL( x )
      RETURN IIF( x, 1, 0 )
   CASE HB_ISSTRING( x )
      IF Empty( x )
         RETURN 0
      ENDIF
      RETURN IIF( DC_IsNumText( x ), DC_Val( x ), DC_Err( "#VALUE!" ) )
   CASE DC_IsErr( x )
      RETURN x
   ENDCASE
RETURN DC_Err( "#VALUE!" )

FUNCTION DC_ToText( x )
   DO CASE
   CASE HB_ISNIL( x )
      RETURN ""
   CASE HB_ISSTRING( x )
      RETURN x
   CASE HB_ISNUMERIC( x )
      RETURN DC_NumText( x )
   CASE HB_ISLOGICAL( x )
      RETURN IIF( x, "TRUE", "FALSE" )
   CASE DC_IsErr( x )
      RETURN x:cCode
   ENDCASE
RETURN ""

FUNCTION DC_ToLog( x )
   DO CASE
   CASE HB_ISNIL( x )
      RETURN .F.
   CASE HB_ISLOGICAL( x )
      RETURN x
   CASE HB_ISNUMERIC( x )
      RETURN x != 0
   CASE HB_ISSTRING( x )
      IF Upper( AllTrim( x ) ) == "TRUE"
         RETURN .T.
      ELSEIF Upper( AllTrim( x ) ) == "FALSE"
         RETURN .F.
      ENDIF
   CASE DC_IsErr( x )
      RETURN x
   ENDCASE
RETURN DC_Err( "#VALUE!" )

/* Val() with an exponent: "1.5E3" -> 1500 */
FUNCTION DC_Val( c )
   LOCAL n
   c := Upper( AllTrim( c ) )
   n := At( "E", c )
   IF n > 0
      RETURN Val( Left( c, n - 1 ) ) * 10 ^ Val( SubStr( c, n + 1 ) )
   ENDIF
RETURN Val( c )

/* "12", "-1.5", "3e4", " 7 " */
FUNCTION DC_IsNumText( c )
   LOCAL i := 1, n, lDig := .F., lDot := .F., ch
   c := AllTrim( c )
   n := Len( c )
   IF n == 0
      RETURN .F.
   ENDIF
   IF Left( c, 1 ) $ "+-"
      i := 2
   ENDIF
   DO WHILE i <= n
      ch := SubStr( c, i, 1 )
      DO CASE
      CASE IsDigit( ch )
         lDig := .T.
      CASE ch == "." .AND. ! lDot
         lDot := .T.
      CASE Upper( ch ) == "E" .AND. lDig .AND. i < n
         i++
         IF SubStr( c, i, 1 ) $ "+-"
            i++
         ENDIF
         IF i > n
            RETURN .F.
         ENDIF
         DO WHILE i <= n
            IF ! IsDigit( SubStr( c, i, 1 ) )
               RETURN .F.
            ENDIF
            i++
         ENDDO
         RETURN .T.
      OTHERWISE
         RETURN .F.
      ENDCASE
      i++
   ENDDO
RETURN lDig
