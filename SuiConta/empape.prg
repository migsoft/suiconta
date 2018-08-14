#include "minigui.ch"

Function EmpApe()

   IF EJERCERRADO=1
      MSGSTOP("El siguiente ejercicio esta aperturado","error")
      RETURN
   ENDIF

   NUMEMP2:=1
   DIR1:=DIRECTORY(RUTAPROGRAMA+"\SUIZO*.*","D")
   FOR N=1 TO LEN(DIR1)
      IF DIR1[N,5]="D"
         NUMEMP2=MAX(NUMEMP2,VAL(SUBSTR(DIR1[N,1],6,3)))
      ENDIF
   NEXT
   AbrirDBF("empresa",,,,RUTAPROGRAMA)
   GO TOP
   DO WHILE .NOT. EOF()
      NUMEMP2:=MAX(NUMEMP2,NEMP)
      SKIP
   ENDDO
   USE
   NUMEMP2++

   DEFINE WINDOW WinEmpApe ;
      AT 50,0    ;
      WIDTH 600  ;
      HEIGHT 350 ;
      TITLE "Aperturar ejercicio" ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("TURQUESA") ;
      ON RELEASE CloseTables()

      @ 05,010 LABEL L_Empresa1 VALUE 'Empresa:' AUTOSIZE TRANSPARENT
      @ 05,080 LABEL L_Empresa2 VALUE RTRIM(NOMEMPRESC)+" "+STR(EJERANY+1,4) BOLD AUTOSIZE TRANSPARENT

      @ 25,010 LABEL L_Directorio1 VALUE 'Direcctorio:' AUTOSIZE TRANSPARENT
      @ 25,080 LABEL L_Directorio2 VALUE RUTAPROGRAMA+"\SUIZO"+STRZERO(NUMEMP2,3) BOLD AUTOSIZE TRANSPARENT


      @ 45,010 RICHEDITBOX Texto1 WIDTH 450 HEIGHT 70 VALUE ;
            'Se recomienda  haber  realizado los asientos de cierre del ejercicio,'+HB_OsNewLine()+ ;
            '(Asiento de regularizacion y Asiento de Cierre).'+HB_OsNewLine()+ ;
            'Se pueden realizar despues de este proceso y traspasar'+HB_OsNewLine()+ ;
            'los saldos al nuevo ejercicio.' BOLD BACKCOLOR MiColor("ROJOCLARO")


      @125,010 PROGRESSBAR P_progress1 RANGE 0 , 100 WIDTH 100 HEIGHT 18 SMOOTH
      @125,115 LABEL L_Label1 VALUE 'Creando directorios' AUTOSIZE TRANSPARENT

      @150,010 PROGRESSBAR P_progress2 RANGE 0 , 100 WIDTH 100 HEIGHT 18 SMOOTH
      @150,115 LABEL L_Label2 VALUE 'Creando ficheros' AUTOSIZE TRANSPARENT

      @175,010 PROGRESSBAR P_progress3 RANGE 0 , 100 WIDTH 100 HEIGHT 18 SMOOTH
      @175,115 LABEL L_Label3 VALUE 'Actualizar subcuentas' AUTOSIZE TRANSPARENT

      @200,010 PROGRESSBAR P_progress4 RANGE 0 , 100 WIDTH 100 HEIGHT 18 SMOOTH
      @200,115 LABEL L_Label4 VALUE 'Actualizar codigos' AUTOSIZE TRANSPARENT

      @225,010 PROGRESSBAR P_progressReg RANGE 0 , 100 WIDTH 100 HEIGHT 18 SMOOTH
      @225,115 LABEL L_LabelReg VALUE 'Regenerando ficheros' AUTOSIZE TRANSPARENT


      @280,010 BUTTON Bt_Aperturar ;
         CAPTION 'Aperturar nuevo ejercicio' ;
         WIDTH 175 HEIGHT 25 ;
         ACTION EmpApe2() ;
         NOTABSTOP

      @280,210 BUTTONEX Bt_Salir ;
         CAPTION 'Salir' ;
         ICON icobus('salir') ;
         ACTION WinEmpApe.Release ;
         WIDTH 90 HEIGHT 25 ;
         NOTABSTOP

   END WINDOW
   
   VentanaCentrar("WinEmpApe","Ventana1","Alinear")
   CENTER WINDOW WinEmpApe
   ACTIVATE WINDOW WinEmpApe

STATIC FUNCTION EmpApe2()

   WinEmpApe.Bt_Aperturar.Enabled:=.F.
   WinEmpApe.Bt_Salir.Enabled:=.F.

***CREAR DIRECTORIO***
   WinEmpApe.P_progress1.Value:=20
   RUTAEMPRESA2:=WinEmpApe.L_Directorio2.Value
   CreateFolder(RUTAEMPRESA2)
   WinEmpApe.P_progress1.Value:=100

***CREAR FICHEROS***
   WinEmpApe.P_progress2.Value:=20
   SetCurrentFolder(RUTAEMPRESA2)
   set default to &RUTAEMPRESA2
   RUTAEMPRESA3:=RUTAEMPRESA
   RUTAEMPRESA:=RUTAEMPRESA2
   W_RegficConta("TODOS","SINVENTANA")
   RUTAEMPRESA:=RUTAEMPRESA3
   SetCurrentFolder(RUTAEMPRESA)
   set default to &RUTAEMPRESA
   WinEmpApe.P_progress2.Value:=100

***ACTUALIZAR FICHERO CUENTAS***
   WinEmpApe.P_progress3.Value:=20
   AbrirDBF("CUENTAS",,,"EXCLUSIVE")
   COPY TO &RUTAEMPRESA2\CUENTAS
   REPLACE DEBE WITH 0 , HABER WITH 0 , SALDO WITH 0 ALL
   USE
   WinEmpApe.P_progress3.Value:=100

***ACTUALIZAR FICHERO CODIGOS***
   WinEmpApe.P_progress4.Value:=20
   AbrirDBF("TOPCOD")
   COPY TO &RUTAEMPRESA2\TOPCOD
   WinEmpApe.P_progress4.Value:=100



***REGENERAR FICHEROS***
   WinEmpApe.P_progressReg.Value:=20
   SetCurrentFolder(RUTAEMPRESA2)
   set default to &RUTAEMPRESA2
   RUTAEMPRESA3:=RUTAEMPRESA
   RUTAEMPRESA:=RUTAEMPRESA2
   W_RegficConta("TODOS","SINVENTANA")
   RUTAEMPRESA:=RUTAEMPRESA3
   SetCurrentFolder(RUTAEMPRESA)
   set default to &RUTAEMPRESA
   WinEmpApe.P_progressReg.Value:=100


***COPIAR FICHEROS INI***
   FICINI:=DIRECTORY(RUTAEMPRESA+"\*.*")
   FOR N=1 TO LEN(FICINI)
      IF UPPER(RIGHT(FICINI[N,1],4))=".INI" .OR. ;
         UPPER(RIGHT(FICINI[N,1],4))=".MEM"
         DIREC2:=RUTAEMPRESA +"\"+FICINI[N,1]
         DIREA2:=RUTAEMPRESA2+"\"+FICINI[N,1]
         COPY FILE (DIREC2) TO (DIREA2)
      ENDIF
   NEXT


   AbrirDBF("empresa",,,,RUTAPROGRAMA)
   GO RECNOEMPRESA
   IF RLOCK()
      REPLACE NEMPSIG   WITH NUMEMP2
      REPLACE CIERREJER WITH 1
      DBCOMMIT()
      DBUNLOCK()
   ENDIF
   Duplicar_RegistroDBF()
   IF RLOCK()
      REPLACE NEMP      WITH NUMEMP2
      REPLACE EJERCICIO WITH EJERANY+1
      REPLACE NEMPANT   WITH NUMEMP
      REPLACE NEMPSIG   WITH 0
      REPLACE RUTA      WITH RIGHT(RUTAEMPRESA2,LEN(RUTAEMPRESA2)-RAT("\",RUTAEMPRESA2))
      REPLACE RUTACONTA WITH ''
      REPLACE CIERREJER WITH 0
      DBCOMMIT()
      DBUNLOCK()
   ENDIF

   MSGINFO("La apertura del ejercicio se ha realizado con exito"+HB_OsNewLine()+HB_OsNewLine()+ ;
           "Recuerde modificar la ruta al programa comercial")

   WinEmpApe.Bt_Salir.Enabled:=.T.

