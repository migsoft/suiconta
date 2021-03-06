PROCEDURE REMIN58()
LOCAL PAG:=0,LINR:=0,LINL:=0,TOT1:=0
AbrirDBF("empresa",,,,RUTAREMEMP)
GO RECNOEMPRESA
NORNIFEMP:=IMPNIF(CIF)
NORNOMEMP:=IF(FIELDPOS("NOMEMP")<>0,RTRIM(NOMEMP),RTRIM(EMP))
NORDIREMP:=RTRIM(DIRECCION)
NORPOBEMP:=RTRIM(CODPOS+"-"+POBLACION)

SELEC FIN
GO TOP
CODBAN2:=VAL(RIGHT(STR(CODBAN),4))
EURO2:=1

AbrirDBF("CUENTAS",,,,RUTAREMESA)
SEEK FIN->CODBAN
NORCTAEMP:=CTA_BAN_SUIZO(BANCTA,20)

SELEC FIN
GO TOP
DO WHILE .NOT. EOF()
   SET PRINT ON
   IF LINL=0
*** CABECERA DE PRESENTADOR
      ?? IF(EURO2=0,"01","51") //A1
      ?? "70" //A2
      ?? PADL(NORNIFEMP,9,"0") //B1-1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
      ?? DIA(FREM,-6) //B2
      ?? SPACE(6) //B3
      ?? PADR(NORNOMEMP,40," ") //C
      ?? SPACE(20) //D
      ?? PADL(SUBSTR(NORCTAEMP,1,4),4,"0") //E1
      ?? PADL(SUBSTR(NORCTAEMP,5,4),4,"0") //E2
      ?? SPACE(12)
      ?? SPACE(40)
      ?? SPACE(14)
*** CABECERA DE ORDENANTE
      ?  IF(EURO2=0,"03","53") //A1
      ?? "70" //A2
      ?? PADL(NORNIFEMP,9,"0") //B1-1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
      ?? DIA(FREM,-6) //B2
      ?? SPACE(6) //B3
      ?? PADR(NORNOMEMP,40," ") //C
      ?? PADL(SUBSTR(NORCTAEMP,1,4),4,"0")    //D1
      ?? PADL(SUBSTR(NORCTAEMP,5,4),4,"0")    //D2
      ?? PADL(SUBSTR(NORCTAEMP,9,2),2,"*")    //D3
      ?? PADL(SUBSTR(NORCTAEMP,11,10),10,"0") //D4
      ?? SPACE(8)  //E1
      ?? "06"      //E2
      ?? SPACE(10) //E3
      ?? SPACE(40) //F
      ?? SPACE( 2) //G1
      AbrirDBF("CUENTAS",,,,RUTAREMESA)
      SEEK CODBAN2
      ?? PADR(CODPOS,9," ") //G2
      ?? SPACE( 3) //G3
   ENDIF
*** INDIVIDUAL OBLIGATORIO
   SELEC FIN
   ?  IF(EURO2=0,"06","56") //A1
   ?? "70" //A2
   ?? PADL(NORNIFEMP,9,"0") //B1-1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
   ?? PADR(ALLTRIM(STR(CODCTA)),12," ") //B2
   ?? PADR(NOMCTA,40," ") //C
   BANCTA2:=CTA_BAN_SUIZO(BANCTA,20)
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,1,4)),4,"0")    //D1
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,5,4)),4,"0")    //D2
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,9,2)),2,"*")    //D3
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,11,10)),10,"0") //D4
   ?? PADL(LTRIM(FTONUMNOR(IMPORTE,EURO2)),10,"0") //E
   ?? PADR(ALLTRIM(NFRA),6," ")      //F1
   ?? PADR(ALLTRIM(STR(CODCTA)),10," ") //F2
   ?? PADR("FACTURA N� "+ALLTRIM(NFRA),40," ") //G
   ?? DIA(FVTO,-6) //H1
   ?? SPACE(2) //H2
*** REGISTRO DE DOMICILIO //PARA SIN CCC
   SELEC FIN
   ?  IF(EURO2=0,"06","56") //A1
   ?? "76" //A2
   ?? PADL(NORNIFEMP,9,"0") //B1-1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
   ?? PADR(ALLTRIM(STR(CODCTA)),12," ") //B2
   CODCTA2:=CODCTA
   AbrirDBF("CUENTAS",,,,RUTAREMESA)
   SEEK CODCTA2
   ?? PADR(DIRCTA,40," ") //C
   ?? PADR(POBCTA,35," ") //D1
   ?? PADR(CODPOS, 5," ") //D2
   SEEK CODBAN2
   ?? PADR(POBCTA,38," ") //E1
   ?? PADR(PROCTA, 2," ") //E2
   SELEC FIN
   ?? DIA(FREM,-6) //F1
   ?? SPACE(8) //F2
   SET PRINT OFF
   TOT1:=TOT1+IMPORTE
   LINL++
   SELEC FIN
   SKIP
ENDDO
*** TOTAL ORDENANTE
SET PRINT ON
?  IF(EURO2=0,"08","58") //A1
?? "70" //A2
?? PADL(NORNIFEMP,9,"0") //B1-1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
?? SPACE(12) //B2
?? SPACE(40) //C
?? SPACE(20) //D
?? PADL(LTRIM(FTONUMNOR(TOT1,EURO2)),10,"0") //E1
?? SPACE(6)  //E2
?? PADL(LTRIM(STR(LINL)),10,"0") //F1
?? PADL(LTRIM(STR((LINL*2)+2)),10,"0") //F2
?? SPACE(20) //F3
?? SPACE(18) //G
*** TOTAL GENERAL
?  IF(EURO2=0,"09","59") //A1
?? "70" //A2
?? PADL(NORNIFEMP,9,"0") //B1-1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //B1-2
?? SPACE(12) //B2
?? SPACE(40) //C
?? PADL(LTRIM(STR(1)),4,"0") //D1
?? SPACE(16) //D2
?? PADL(LTRIM(FTONUMNOR(TOT1,EURO2)),10,"0") //E1
?? SPACE(6)  //E2
?? PADL(LTRIM(STR(LINL)),10,"0") //F1
?? PADL(LTRIM(STR((LINL*2)+4)),10,"0") //F2
?? SPACE(20) //F3
?? SPACE(18) //G
?
SET PRINT OFF
SET PRINTER TO

