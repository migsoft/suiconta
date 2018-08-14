#include "minigui.ch"

Function EmpSini()

   AbrirDBF("empresa",,,,RUTAPROGRAMA)
   GO RECNOEMPRESA
   IF NEMPSIG=0
      MSGSTOP("El siguiente ejercicio no esta aperturado","error")
      RETURN
   ENDIF
	SEEK NEMPSIG
	NOMEMPSIG:=EMP
	EJEEMPSIG:=EJERCICIO
   RUTASIG:=RUTAPROGRAMA+"\"+LTRIM(RTRIM(RUTA))
   FASI2:=CTOD("01-01-"+LTRIM(STR(EJERCICIO)))

   DEFINE WINDOW WinEmpSini ;
      AT 50,0    ;
      WIDTH 600  ;
      HEIGHT 350 ;
      TITLE "Traspasar saldos al siguiente ejercicio" ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("TURQUESA") ;
      ON RELEASE CloseTables()

      @ 05,010 LABEL L_Empa VALUE 'Empresa inicial:' AUTOSIZE TRANSPARENT
      @ 05,115 LABEL L_Emp1 VALUE RTRIM(NOMEMPRESC)+" "+STR(EJERANY,4) BOLD AUTOSIZE TRANSPARENT

      @ 25,010 LABEL L_Empb VALUE 'Empresa destino:' AUTOSIZE TRANSPARENT
      @ 25,115 LABEL L_Emp2 VALUE RTRIM(NOMEMPSIG)+" "+STR(EJERCICIO,4) BOLD AUTOSIZE TRANSPARENT


      @ 45,010 RICHEDITBOX Texto1 WIDTH 450 HEIGHT 70 VALUE ;
            'Se recomienda  haber  realizado los asientos de cierre del ejercicio,'+HB_OsNewLine()+ ;
            '(Asiento de regularizacion y Asiento de Cierre).'+HB_OsNewLine()+ ;
            'Tambien repasar los asientos descuadrados' BOLD BACKCOLOR MiColor("ROJOCLARO")


      @125,010 LABEL L_Label1 VALUE 'Traspasando saldos' AUTOSIZE TRANSPARENT
      @125,130 PROGRESSBAR P_progress1 RANGE 0 , 100 WIDTH 330 HEIGHT 18 SMOOTH

      @150,010 RICHEDITBOX Texto2 WIDTH 450 HEIGHT 70 VALUE '' BOLD BACKCOLOR MiColor("ROJOCLARO")



      @280,010 BUTTON Bt_Traspasar ;
         CAPTION 'Traspasar saldos' ;
         WIDTH 175 HEIGHT 25 ;
         ACTION EmpSini2() ;
         NOTABSTOP

      @280,210 BUTTONEX Bt_Salir ;
         CAPTION 'Salir' ;
         ICON icobus('salir') ;
         ACTION WinEmpSini.Release ;
         WIDTH 90 HEIGHT 25 ;
         NOTABSTOP

   END WINDOW
   VentanaCentrar("WinEmpSini","Ventana1","Alinear")

   CENTER WINDOW WinEmpSini
   ACTIVATE WINDOW WinEmpSini

Return Nil

STATIC FUNCTION EmpSini2()

   WinEmpSini.Bt_Traspasar.Enabled:=.F.
   WinEmpSini.Bt_Salir.Enabled:=.F.

   FIC_CIERRE(0,99999999,DMA2,DMA2,2,1)

   *** PASANDO SALDOS REALES ***
   SELEC 1
   IF .NOT. FILE(RUTASIG+"\CUENTAS.DBF") .OR. .NOT. FILE(RUTASIG+"\CUENTAS.CDX") .OR. ;
      .NOT. FILE(RUTASIG+"\APUNTES.DBF") .OR. .NOT. FILE(RUTASIG+"\APUNTES.CDX")
      MsgStop("El fichero de la siguiente empresa no existe o no esta indexado"+HB_OsNewLine()+ ;
              "No se han traspasado los saldos")
      RETURN
   ENDIF
   AbrirDBF("CUENTAS",,,"Exclusive",RUTASIG)
   AbrirDBF("FINSUIC",,,"Exclusive",gettempdir())
   AbrirDBF("CIERRE",,,"Exclusive")
   AbrirDBF("APUNTES",,,"Exclusive",RUTASIG)
   DELETE FOR NASI=1
   PACK

   SELEC CIERRE
   IF LASTREC()=0
      MsgStop("No se ha realizado el asiento de cierre")
      RETURN
   ENDIF
   SORT ON CODCTA FOR NASI=999998 TO FIN2
   AbrirDBF("FIN2","SIN_INDICE",,"Exclusive")
   INDEX ON CODCTA TO FIN2
   SALDOMAL:=0
   APU2:=1
   WinEmpSini.P_progress1.RangeMin:=0
   WinEmpSini.P_progress1.RangeMax:=LASTREC()
   DO WHILE .NOT. EOF()
      WinEmpSini.P_progress1.Value:=WinEmpSini.P_progress1.Value+1
      DO EVENTS
      YCOD=CODCTA
      YSAL=DEBE-HABER
      YSAL=YSAL*-1 //ASIENTO APERTURA CONTRARIO AL ASIENTO DE CIERRE

      ***COMPROBAR SALDO FINAL CON ASIENTO DE CIERRE***
      SELEC FINSUIC
      SEEK YCOD
      YNOM=NOMCTA
      IF STR(YSAL,15,MDA_DEC(EJERANY))<>STR(SALDO+DEBE-HABER,15,MDA_DEC(EJERANY))
         SALDOMAL:=1
         WinEmpSini.Texto2.Value:=WinEmpSini.Texto2.Value+STR(YCOD)+" "+MIL(YSAL,12,2)+" "+MIL(SALDO+DEBE-HABER,12,2)+HB_OsNewLine()
      ENDIF

      SELEC CUENTAS
      SEEK YCOD
      IF EOF()
         APPEND BLANK
         REPLACE CODCTA WITH YCOD
         REPLACE NOMCTA WITH YNOM
      ENDIF

      IF YSAL<>0
         REPLACE SALDO WITH SALDO+YSAL

         SELEC APUNTES
         APPEND BLANK
         REPLACE NASI WITH 1
         REPLACE APU WITH IF(APU2>=990,990,APU2++)
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FASI2
         REPLACE CODCTA WITH YCOD
         REPLACE NOMAPU WITH "Asiento de apertura"
         IF YSAL>0
            REPLACE DEBE WITH YSAL
         ELSE
            REPLACE HABER WITH YSAL*-1
         ENDIF
      ENDIF

      SELEC FIN2
      SKIP
   ENDDO
   CLOSE DATABASES

   IF SALDOMAL=1
      MsgStop("No coinciden los saldos finales con el asiento de cierre"+HB_OsNewLine()+ ;
              "Se recomienda volver a realizar el asiento de cierre")
   ENDIF
   ERASE FIN.DBF
   ERASE FIN.CDX
   ERASE FIN2.DBF
   ERASE FIN2.CDX

   MSGINFO("El traspaso de los saldos iniciales se ha realizado con exito")

   WinEmpSini.Bt_Salir.Enabled:=.T.

Return Nil