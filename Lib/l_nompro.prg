FUNCTION NomPropio(NomPropio1)
   NomPropio2:=LOWER(NomPropio1)      //hay que ponerlo en minusculas primero
   NomPropio2:=TOKENUPPER(NomPropio2) //funciona 16-Abril-2008
RETURN(NomPropio2)


FUNCTION TituloqImp(TituloImp2)
   IF AT("IMPRIMIR: ",UPPER(TituloImp2))<>0
		TituloImp2:=SUBSTR(TituloImp2,11,LEN(TituloImp2))
	ENDIF
/* Lo combierte todo a mayusculas
   TituloImp2:=STRTRAN(UPPER(TituloImp2),"IMPRIMIR: ","")
   TituloImp2:=UPPER(LEFT(TituloImp2,1))+)
*/
RETURN(TituloImp2)


FUNCTION NomPropio2(NomPropio1)
   SiMayus:=.T.
   NomPropio2:=""
   NomPropioL1:="—¡…Õ”⁄¿»Ã“ŸƒÀœ÷‹¬ Œ‘€"
   NomPropioL2:="Ò·ÈÌÛ˙‡ËÏÚ˘‰ÎÔˆ¸‚ÍÓÙ˚"
   FOR NNP1=1 TO LEN(NomPropio1)
      NomPropioL:=SUBSTR(NomPropio1,NNP1,1)
      DO CASE
      CASE ISALPHA(NomPropioL)
         IF SiMayus=.T.
            NomPropio2:=NomPropio2+UPPER(NomPropioL)
         ELSE
            NomPropio2:=NomPropio2+LOWER(NomPropioL)
         ENDIF
         SiMayus:=.F.
      CASE ISALPHA(NomPropioL)
         IF SiMayus=.T.
            NomPropio2:=NomPropio2+UPPER(NomPropioL)
         ELSE
            NomPropio2:=NomPropio2+LOWER(NomPropioL)
         ENDIF
         SiMayus:=.F.
      CASE AT(NomPropioL,NomPropioL1)<>0
         IF SiMayus=.T.
            NomPropio2:=NomPropio2+NomPropioL
         ELSE
            NomPropio2:=NomPropio2+SUBSTR(NomPropioL2,AT(NomPropioL,NomPropioL1),1)
         ENDIF
         SiMayus:=.F.
      CASE AT(NomPropioL,NomPropioL2)<>0
         IF SiMayus=.T.
            NomPropio2:=NomPropio2+SUBSTR(NomPropioL1,AT(NomPropioL,NomPropioL2),1)
         ELSE
            NomPropio2:=NomPropio2+NomPropioL
         ENDIF
         SiMayus:=.F.
      OTHERWISE
         NomPropio2:=NomPropio2+NomPropioL
         SiMayus:=.T.
      ENDCASE
   NEXT
RETURN(NomPropio2)
