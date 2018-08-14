#include "minigui.ch"
#include "winprint.ch"

FUNCTION CUESAL()

   DEFINE WINDOW W_CueSal1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 100 ;
      TITLE 'Actualizando ficheros' ;
      MODAL NOSIZE ;
      ON INIT CUESALp()

      @015,10 LABEL L_Apuntes VALUE 'Apuntes' AUTOSIZE TRANSPARENT
      @010,80 PROGRESSBAR P_Progres1 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

      @045,10 LABEL L_Cierre VALUE 'Cierre' AUTOSIZE TRANSPARENT
      @040,80 PROGRESSBAR P_Progres2 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

   END WINDOW
   VentanaCentrar("W_CueSal1","Ventana1")
   ACTIVATE WINDOW W_CueSal1


STATIC FUNCTION CUESALp()
   AbrirDBF("CUENTAS",,,"Exclusive",,1)
   REPLACE SALDO WITH 0 ALL

FOR NUMCS2=1 TO 2
   IF NUMCS2=1
      NOMFICAPU:="APUNTES"
   ELSE
      NOMFICAPU:="CIERRE"
   ENDIF
   AbrirDBF(NOMFICAPU)
   GO TOP
   NASI3:=0
   nProgres:=1
   DO WHILE .NOT. EOF()
      DO EVENTS
      IF NUMCS2=1
         W_CueSal1.P_Progres1.Value:=nProgres++*100/LASTREC()
      ELSE
         W_CueSal1.P_Progres2.Value:=nProgres++*100/LASTREC()
      ENDIF

      IF NASI=999998 //ASIENTO DE CIERRE
         SKIP
         LOOP
      ENDIF
      IF NASI=0 //PONER -1 PARA PODER BORRAR O CORREGIR
         NASI3:=1
      ENDIF
      NASI2:=NASI
      CODCTA2:=CODCTA
      SALDO2:=DEBE-HABER
      AbrirDBF("CUENTAS",,,"Exclusive",,1)
      SEEK CODCTA2
      IF EOF()
         APPEND BLANK
         REPLACE CODCTA WITH CODCTA2
         REPLACE NOMCTA WITH "ALTA ASIENTO: "+LTRIM(STR(NASI2))
      ENDIF
      REPLACE SALDO WITH SALDO+SALDO2
      AbrirDBF(NOMFICAPU)
      SKIP
   ENDDO
   IF NASI3=1
      AbrirDBF(NOMFICAPU)
      IF FLOCK()
         REPLACE NASI WITH -1 FOR NASI=0
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
   ENDIF
NEXT

AbrirDBF("CUENTAS",,,"Exclusive",,1)
USE

W_CueSal1.release

