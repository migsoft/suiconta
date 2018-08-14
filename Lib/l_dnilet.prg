FUNCTION DNI_LET(DNILET2,NIVEL)
LOCAL DNILET3:="",DNILETL1:=" ",DNILETL2:=" "
** NIVEL 1-NO SIGNOS
** NIVEL 2-DEVOLVER SI ES ERRONEA
NIVEL:=IF(NIVEL=NIL,0,NIVEL)
DNILET2:=UPPER(DNILET2)
DNIERROR:=.F.

FOR N=1 TO LEN(RTRIM(DNILET2))
   IF ISALPHA(SUBSTR(DNILET2,N,1))=.T. .OR. ;
      ISDIGIT(SUBSTR(DNILET2,N,1))=.T.
      DNILET3:=DNILET3+SUBSTR(DNILET2,N,1)
   ENDIF
NEXT

IF NIVEL=1
   DNILET3:=PADR(DNILET3,LEN(DNILET2)," ")
ELSE
   IF ISALPHA(LEFT(DNILET3,1))=.T. //CIF
      **Sumamos los dígitos pares (sin contar la letra)
      SumaA:=VAL(SUBSTR(DNILET3,3,1))+VAL(SUBSTR(DNILET3,5,1))+VAL(SUBSTR(DNILET3,7,1))
      **Para cada uno de los dígitos de la posiciones impares, multiplicarlo por 2 y sumar los dígitos del resultado.
      SumaB:=SumDigNum(VAL(SUBSTR(DNILET3,2,1))*2) + ;
             SumDigNum(VAL(SUBSTR(DNILET3,4,1))*2) + ;
             SumDigNum(VAL(SUBSTR(DNILET3,6,1))*2) + ;
             SumDigNum(VAL(SUBSTR(DNILET3,8,1))*2)
      SumaC:=SumaA+SumaB
      SumaD:=10-VAL(RIGHT(RTRIM(STR(SumaC)),1))
      DO CASE
      CASE LEFT(DNILET3,1)="K" .OR. LEFT(DNILET3,1)="P" .OR. LEFT(DNILET3,1)="Q" .OR. LEFT(DNILET3,1)="S"
         SumaL:=CHR(SumaD+64)
      CASE LEFT(DNILET3,1)="X"
         SumaL:=RIGHT(DNI_LET(SUBSTR(DNILET3,2,7)),1) //TOMAR LA ULTIMA LETRA DEL NIF
      OTHERWISE
         SumaL:=RIGHT(STR(SumaD),1)
      ENDCASE
      IF RIGHT(DNILET3,1)=SumaL
         DNIERROR:=.F.
      ELSE
         DNIERROR:=.T.
      ENDIF
   ELSE
      IF ISALPHA(RIGHT(RTRIM(DNILET3),1))=.T. //NIF
         DNILETL1:=RIGHT(RTRIM(DNILET3),1)
         ncDNILET3:=SUBSTR(DNILET3,1,LEN(DNILET3)-1)
         nnDNILET3:=VAL(ncDNILET3)
      ELSE
         ncDNILET3:=DNILET3
         nnDNILET3:=VAL(ncDNILET3)
      ENDIF
      IF nnDNILET3<>0
         DNILET4=nnDNILET3%23
         IF DNILET4<0 .OR. DNILET4>22
            DNILET4:=0
         ENDIF
         DNILETL2=SUBSTR("TRWAGMYFPDXBNJZSQVHLCKE",DNILET4+1,1)
         IF DNILETL1=" "
            DNILET3=ncDNILET3+DNILETL2
         ENDIF
         IF DNILETL2<>DNILETL1 .AND. DNILETL1<>" "
            DNILET4:=ncDNILET3+DNILETL2
            IF NIVEL=2 //2-NO PREGUNTAR
               DNIERROR:=.T.
            ELSE
               IF MSGYESNO("D.N.I. Nacional erroneo"+HB_OsNewLine()+DNILET2+" DNI erroneo "+HB_OsNewLine()+DNILET4+" DNI correcto"+HB_OsNewLine()+HB_OsNewLine()+ ;
                        "¿Desea modificar en DNI?","error")=.T.
                  DNILET3=DNILET4
               ENDIF
            ENDIF
         ENDIF
      ENDIF
   ENDIF
ENDIF

DO CASE
CASE NIVEL=0
   RETURN(DNILET3)
CASE NIVEL=1
   RETURN(DNILET3)
CASE NIVEL=2
   RETURN(DNIERROR)
ENDCASE
RETURN(DNILET3)



***Sumar los digitos de un numero
***125 = 1+2+5 = 8
FUNCTION SumDigNum(SumDigNum1)
SumDigNum2:=LTRIM(STR(SumDigNum1))
DO WHILE LEN(SumDigNum2)>1
   SumDigNum3:=0
   FOR N=1 TO LEN(SumDigNum2)
      SumDigNum3:=SumDigNum3+VAL(SUBSTR(SumDigNum2,N,1))
   NEXT
   SumDigNum2:=LTRIM(STR(SumDigNum3))
ENDDO
RETURN(VAL(SumDigNum2))
