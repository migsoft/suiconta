
FUNCTION SERIE(SERIELET,SERIENUM,SERIELIN,NIVEL)
	NIVEL:=IF(NIVEL=NIL,0,NIVEL)
	IF NIVEL=2 .OR. NIVEL=-2
	   SERIELET:=CHR(ASC(SERIELET)+1)
	   SERIEFIN:=LEFT(LTRIM(SERIELET)+" ",1)+SPACE(10+3)
	ELSE
		SERIELIN:=IF(SERIELIN=NIL,"",STR(SERIELIN,3))
	   SERIEFIN:=LEFT(LTRIM(SERIELET)+" ",1)+STR(SERIENUM,10)+SERIELIN
	ENDIF
RETURN(SERIEFIN)

***SERIE USADA EN CLIPPER***
FUNCTION LETSER2(LetSerL,LetSerN,NIVEL)
LOCAL LetSerF,LetSerL2,LetSerN2
NIVEL:=IF(NIVEL=NIL,0,NIVEL)
DO CASE
CASE NIVEL=0
   IF LetSerN<0
      LetSerN2:="0"+STR(LetSerN,6) //NEGATIVOS
   ELSE
      LetSerN2:="1"+STR(LetSerN,6) //POSITIVOS
   ENDIF
   LetSerF:=LEFT(LetSerL,1)+LetSerN2
   RETURN(LetSerF)
CASE NIVEL=1 .OR. NIVEL=-1
   DO CASE
      CASE ASC(LetSerL)>=65 .AND. ASC(LetSerL)<=90
         LetSerF=1
      CASE ASC(LetSerL)=32 .OR. ASC(LetSerL)=48
         LetSerF=1
      OTHERWISE
         LetSerF=0
   ENDCASE
   RETURN(LetSerF)
CASE NIVEL=2 .OR. NIVEL=-2
   LetSerL2:=CHR(ASC(LetSerL)+1)
   LetSerF=LetSerL2+SPACE(1+6)
   RETURN(LetSerF)
ENDCASE
RETURN

FUNCTION MiSerie(MiSerieT,MiSerieN)
   MiSerieR:=MiSerieT%MiSerieN
   MiSerieR:=IF(MiSerieR=0,MiSerieN,MiSerieR)
RETURN(MiSerieR)
