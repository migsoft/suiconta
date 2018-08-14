FUNCTION TotReg()
   rTotReg:=RECNO()
   nTotReg:=0
   dbeval( {|| nTotReg++} )
   GO rTotReg
RETURN(nTotReg)


***DbEval( {|| nCount++}, {|| Saldo < 0} ) //contar los que el saldo es mayor de cero

/*
10-02-09
es preferible poner SET DELETE OFF
para recorrer la base solo una vez
pero hay que saltar los registros borrados
   SET DELETE OFF
   IF DELETED()=.T.
      SKIP
      LOOP
   ENDIF
   SET DELETE ON
*/
