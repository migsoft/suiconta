FUNCTION PCODCTA1(PCODCTA) //DE CARACTER A NUMERO
   PCODCTA:=LTRIM(PCODCTA)
   DO CASE
   CASE AT(".",PCODCTA)<>0
      PCODCTA:=VAL( PADR(SUBSTR(PCODCTA,1,AT(".",PCODCTA)-1), 8,"0") )+;
               VAL( SUBSTR(PCODCTA,AT(".",PCODCTA)+1,8) )
   CASE AT(",",PCODCTA)<>0
      PCODCTA:=VAL( PADR(SUBSTR(PCODCTA,1,AT(",",PCODCTA)-1), 8,"0") )+;
               VAL( SUBSTR(PCODCTA,AT(",",PCODCTA)+1,8) )
   OTHERWISE
      PCODCTA:=VAL(PADR(RTRIM(PCODCTA),8,"0"))
   ENDCASE
RETURN(PCODCTA)


FUNCTION PCODCTA2(PCODCTA) //DE NUMERO A CARACTER
   IF PCODCTA=0
      PCODCTA:=SPACE(8)
   ELSE
      PCODCTA:=PADR(LTRIM(STR(PCODCTA)),8," ")
   ENDIF
RETURN(PCODCTA)

FUNCTION PCODCTA3(PCODCTA) //EN CARACTER QUITAR EL PUNTO
   DO CASE
   CASE AT(".",PCODCTA)<>0
      NUM2:=VAL(PADR( LEFT(PCODCTA,AT(".",PCODCTA)-1) ,8,"0" ))
      NUM2:=NUM2+VAL( SUBSTR(PCODCTA,AT(".",PCODCTA)+1,8) )
      PCODCTA:=STR(NUM2,8)
   CASE AT(",",PCODCTA)<>0
      NUM2:=VAL(PADR( LEFT(PCODCTA,AT(",",PCODCTA)-1) ,8,"0" ))
      NUM2:=NUM2+VAL( SUBSTR(PCODCTA,AT(",",PCODCTA)+1,8) )
      PCODCTA:=STR(NUM2,8)
   ENDCASE
RETURN(PCODCTA)

FUNCTION PCODCTA4(PCODCTA) //DE NUMERO O CARACTER A CARACTER CON PUNTO "572.1"
   IF VALTYPE(PCODCTA)="C"
      PCODCTA:=VAL(PCODCTA3(PCODCTA))
   ENDIF
   IF PCODCTA=0
      PCODCTA:=""
   ELSE
      PCODCTA:=STR(PCODCTA,8)
      IF SUBSTR(PCODCTA,4,1)="0"
         PCODCTA:=LEFT(PCODCTA,3)+"."+LTRIM(STR(VAL(RIGHT(PCODCTA,4))))
      ELSE
         PCODCTA:=LEFT(PCODCTA,4)+"."+LTRIM(STR(VAL(RIGHT(PCODCTA,4))))
      ENDIF
   ENDIF
RETURN(PCODCTA)