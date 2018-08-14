/*
10-09-2007
ASOCIACION ESPAÑOLA DE BANCA
Ordenes en fichero para emisión de transferencias y cheques
serie normas y procedimientos bancarios N34-1
Madrid - Octubre 2005
*/

PROCEDURE REMIN34()
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
   IF LINR=0
*** Registros de cabecera de Ordenante
      ?? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "001" //E
      ?? DIA(FREM,-6) //F1
      ?? DIA(FREM,-6) //F2
      ?? PADL(SUBSTR(NORCTAEMP,1,4),4,"0")    //F3
      ?? PADL(SUBSTR(NORCTAEMP,5,4),4,"0")    //F4
      ?? PADL(SUBSTR(NORCTAEMP,9,2),2,"*")    //F5
      ?? PADL(SUBSTR(NORCTAEMP,11,10),10,"0") //F6
      ?? "0" //F7
      ?? SPACE(8) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "002" //E
      ?? PADR(NORNOMEMP,36," ") //F
      ?? SPACE(5) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "003" //E
      ?? PADR(NORDIREMP,36," ") //F
      ?? SPACE(5) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "004" //E
      ?? PADR(NORPOBEMP,36," ") //F
      ?? SPACE(5) //G

*** Registro de Cabecera
      ? "04" //A
      ?? IF(EURO2=0,"06","56") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? SPACE(03) //E
      ?? SPACE(41) //F

   ENDIF

*** Registros de Detalle
   CODCTA2:=CODCTA
   AbrirDBF("CUENTAS",,,,RUTAREMESA)
   SEEK CODCTA2
   IF .NOT. EOF()
      NIFCLI2:=CIF
   ELSE
      NIFCLI2:="1"
   ENDIF
   SELEC FIN

   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(NIFCLI2,12,"0") //D
   ?? "010" //E
   ?? PADL(LTRIM(FTONUMNOR(IMPORTE,EURO2)),12,"0") //F1
   BANCTA2:=CTA_BAN_SUIZO(BANCTA,20)
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,1,4)),4,"0")      //F2
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,5,4)),4,"0")      //F3
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,9,2)),2,"*")      //F4
   ?? PADL(ALLTRIM(SUBSTR(BANCTA2,11,10)),10,"0")   //F5
   ?? "1" //F6
   DO CASE
   CASE TRANOMI=1
      ?? "1" //F7 NOMINA
   CASE TRANOMI=2
      ?? "8" //F7 PENSION
   OTHERWISE
      ?? "9" //F7 OTROS
   ENDCASE
   ?? "1"      //F8
   ?? SPACE(6) //G


   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(NIFCLI2,12,"0") //D
   ?? "011" //E
   ?? PADR(NOMCTA,36," ") //F
   ?? SPACE(5) //G

   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(NIFCLI2,12,"0") //D
   ?? "016" //E
   ?? PADR(NFRA,36," ") //F
   ?? SPACE(5) //G

   SET PRINT OFF
   TOT1:=TOT1+IMPORTE
   LINR++
   SELEC FIN
   SKIP
ENDDO

*** Registro de totales
SET PRINT ON
? "08" //A
??  IF(EURO2=0,"06","56") //B
?? PADL(NORNIFEMP,9,"0") //C1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
?? SPACE(12) //D
?? SPACE( 3) //E
?? PADL(LTRIM(FTONUMNOR(TOT1,EURO2)),12,"0") //F1
?? PADL(LTRIM(STR(LINR)),8,"0") //F2
?? PADL(LTRIM(STR((LINR*3)+2)),10,"0") //F3  lineas 0362 y 0962 no se cuentan
?? SPACE(6) //F4
?? SPACE(5) //G

*** Registro de total general
? "09" //A
??  IF(EURO2=0,"06","62") //B
?? PADL(NORNIFEMP,9,"0") //C1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
?? SPACE(12) //D
?? SPACE( 3) //E
?? PADL(LTRIM(FTONUMNOR(TOT1,EURO2)),12,"0") //F1
?? PADL(LTRIM(STR(LINR)),8,"0") //F2
?? PADL(LTRIM(STR((LINR*3)+7)),10,"0") //F3
?? SPACE(6) //F4
?? SPACE(5) //G

?
SET PRINT OFF
SET PRINTER TO

