#include "minigui.ch"
Function W_Regfic()

   CLOSE DATABASES

   DEFINE WINDOW Regfic ;
      AT 0,0     ;
      WIDTH 800  ;
      HEIGHT 500 ;
      TITLE "Regenerar ficheros "+RUTAEMPRESA ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("GRISCLARO") ;
      ON RELEASE CloseTables()
*      NOSYSMENU

      @ 5,200 LABEL L_Regenerar VALUE 'Regenerar ficheros' WIDTH 400 HEIGHT 25 FONT "ARIAL" SIZE 14 CENTERALIGN

      aLabel:={'Fichero de empresas',;
               'Fichero de cuentas contables',;
               'Fichero de Asientos Contables',;
               'Fichero de Asientos de Cierre',;
               'Fichero de Facturas Emitidas',;
               'Fichero de Cobros de Facturas',;
               'Fichero de Facturas Recibidas',;
               'Fichero de Pagos de Facturas',;
               'Fichero de Vto. Fras. Recibidas',;
               'Fichero de Codigos',;
               'Fichero de Plan General Contable',;
               'Fichero de Remesas'}

      LIN2:=30
      COL2:=10
      FOR N=1 TO LEN(aLabel)
         NOM1:="Si_"+LTRIM(STR(N))
         NOM2:="Progress_"+LTRIM(STR(N))
         NOM3:="Label_"+LTRIM(STR(N))
         @ LIN2  ,COL2     CHECKBOX &NOM1 CAPTION '' WIDTH 15 VALUE .F. BACKCOLOR MiColor("GRISCLARO")
         @ LIN2+5,COL2+020 PROGRESSBAR &NOM2 RANGE 0 , 100 WIDTH 100 HEIGHT 20 SMOOTH
         @ LIN2+5,COL2+125 LABEL &NOM3 VALUE aLabel[N] AUTOSIZE TRANSPARENT
         LIN2:=LIN2+30
      NEXT

      @ 430,010 BUTTON Bt_Marcar1 ;
      CAPTION 'Marcar' ;
      WIDTH 75 HEIGHT 25 ;
      ACTION W_Regfic3("MARCAR")

      @ 430,110 BUTTON Bt_Marcar2 ;
      CAPTION 'Desmarcar' ;
      WIDTH 75 HEIGHT 25 ;
      ACTION W_Regfic3("DESMARCAR")

      @ 430,410 BUTTON Button1 ;
      CAPTION 'Regenerar' ;
      WIDTH 75 HEIGHT 25 ;
      ACTION W_Regfic2()

      @ 430,510 BUTTON Button2 ;
      CAPTION 'Cancelar' ;
      WIDTH 75 HEIGHT 25 ;
      ACTION Regfic.Release

      @ 430,610 BUTTON Button3 ;
      CAPTION 'Salir' ;
      WIDTH 75 HEIGHT 25 ;
      ACTION Regfic.Release

      Regfic.Button3.Enabled := .F.

   END WINDOW
   VentanaCentrar("Regfic","Ventana1")
   ACTIVATE WINDOW Regfic

   Return Nil
*
STATIC FUNCTION W_Regfic2()

***Comprobar si se ha marcado algun fichero***
	SiHacer:=.F.
	FOR N=1 TO LEN(aLabel)
		SiHacer:=GetProperty("Regfic","Si_"+LTRIM(STR(N)),"Value")
		IF SiHacer=.T.
			EXIT
		ENDIF
	NEXT
	IF SiHacer=.F.
		MsgStop("No se ha seleccionado ningun fichero")
		RETURN
	ENDIF
***FIN Comprobar si se ha marcado algun fichero***

IF Regfic.Si_1.Value=.T.
   W_RegficConta("EMPRESA","Regfic")
ENDIF
IF Regfic.Si_2.Value=.T.
   W_RegficConta("CUENTAS","Regfic")
ENDIF
IF Regfic.Si_3.Value=.T.
   W_RegficConta("APUNTES","Regfic")
ENDIF
IF Regfic.Si_4.Value=.T.
   W_RegficConta("CIERRE","Regfic")
ENDIF
IF Regfic.Si_5.Value=.T.
   W_RegficConta("FAC92","Regfic")
ENDIF
IF Regfic.Si_6.Value=.T.
   W_RegficConta("COBROS","Regfic")
ENDIF
IF Regfic.Si_7.Value=.T.
   W_RegficConta("FACREB","Regfic")
ENDIF
IF Regfic.Si_8.Value=.T.
   W_RegficConta("PAGOS","Regfic")
ENDIF
IF Regfic.Si_9.Value=.T.
   W_RegficConta("FACVTO","Regfic")
ENDIF
IF Regfic.Si_10.Value=.T.
   W_RegficConta("TOPCOD","Regfic")
ENDIF
IF Regfic.Si_11.Value=.T.
   W_RegficConta("PGC","Regfic")
ENDIF
IF Regfic.Si_12.Value=.T.
   W_RegficConta("REMESA","Regfic")
ENDIF

***Ajustes de ficheros***
   CLOSE DATABASES
   W_RegficAjustes()
   CLOSE DATABASES
***fin Ajustes de ficheros***

   Regfic.Button1.Enabled := .F.
   Regfic.Button2.Enabled := .F.
   Regfic.Button3.Enabled := .T.

RETURN NIL


*
FUNCTION W_RegficConta(NomFic1,NomVen1)
CLOSE DATABASES
IF UPPER(NomFic1)=="EMPRESA" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( RUTAPROGRAMA+"\EMPRESA.DBF" )
      Crear(RUTAPROGRAMA+"\EMPRESA.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME(RUTAPROGRAMA+"\EMPRESA.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 RUTAPROGRAMA+"\EMPRESA.DBF")
      ELSE
         Crear(RUTAPROGRAMA+"\EMPRESA.DBF")
         Use &RUTAPROGRAMA\EMPRESA Alias EMPRESA
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE(RUTAPROGRAMA+"\EMPRESA.CDX")
   Use &RUTAPROGRAMA\EMPRESA Alias EMPRESA new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NEMP       TAG ORDEN1 to &RUTAPROGRAMA\EMPRESA.CDX EVAL PORCENT(1,1,2,NLAST2,NomVen1)
   Index on UPPER(EMP) TAG ORDEN2 to &RUTAPROGRAMA\EMPRESA.CDX EVAL PORCENT(1,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="CUENTAS" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "CUENTAS.DBF" )
      Crear("CUENTAS.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("CUENTAS.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\CUENTAS.DBF")
      ELSE
         Crear("CUENTAS.DBF")
         Use CUENTAS Alias CUENTAS
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('CUENTAS.CDX')
   Use CUENTAS Alias CUENTAS new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on CODCTA        TAG ORDEN1 to CUENTAS.CDX EVAL PORCENT(2,1,2,NLAST2,NomVen1)
   Index on UPPER(NOMCTA) TAG ORDEN2 to CUENTAS.CDX EVAL PORCENT(2,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
   IF FILE("APUNTES.DBF") .AND. FILE("APUNTES.CDX") .AND. ;
      FILE("CIERRE.DBF") .AND. FILE("CIERRE.CDX")
      CUESAL()
   ENDIF
ENDIF

IF UPPER(NomFic1)=="APUNTES" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "APUNTES.DBF" )
      Crear("APUNTES.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("APUNTES.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\APUNTES.DBF")
      ELSE
         Crear("APUNTES.DBF")
         Use APUNTES Alias APUNTES
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('APUNTES.CDX')
   Use APUNTES Alias APUNTES new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NASI                                  TAG ORDEN1 to APUNTES.CDX EVAL PORCENT(3,1,2,NLAST2,NomVen1)
   Index on STR(CODCTA,8)+DTOS(FECHA)+STR(NASI,6) TAG ORDEN2 to APUNTES.CDX EVAL PORCENT(3,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="CIERRE" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "CIERRE.DBF" )
      Crear("CIERRE.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("CIERRE.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\CIERRE.DBF")
      ELSE
         Crear("CIERRE.DBF")
         Use CIERRE Alias CIERRE
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('CIERRE.CDX')
   Use CIERRE Alias CIERRE new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NASI                                  TAG ORDEN1 to CIERRE.CDX EVAL PORCENT(4,1,2,NLAST2,NomVen1)
   Index on STR(CODCTA,8)+DTOS(FECHA)+STR(NASI,6) TAG ORDEN2 to CIERRE.CDX EVAL PORCENT(4,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="FAC92" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "FAC92.DBF" )
      Crear("FAC92.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("FAC92.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\FAC92.DBF")
      ELSE
         Crear("FAC92.DBF")
         Use FAC92 Alias FAC92
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('FAC92.CDX')
   Use FAC92 Alias FAC92 new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on SERIE(SERFAC,NFAC)        TAG ORDEN1 to FAC92.CDX EVAL PORCENT(5,1,2,NLAST2,NomVen1)
   Index on UPPER(CLIENTE)+DTOS(FFAC) TAG ORDEN2 to FAC92.CDX EVAL PORCENT(5,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="COBROS" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "COBROS.DBF" )
      Crear("COBROS.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("COBROS.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\COBROS.DBF")
      ELSE
         Crear("COBROS.DBF")
         Use COBROS Alias COBROS
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('COBROS.CDX')
   Use COBROS Alias COBROS new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on SERIE(SERFAC,NFAC) TAG ORDEN1 to COBROS.CDX EVAL PORCENT(6,1,1,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="FACREB" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "FACREB.DBF" )
      Crear("FACREB.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("FACREB.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\FACREB.DBF")
      ELSE
         Crear("FACREB.DBF")
         Use FACREB Alias FACREB
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('FACREB.CDX')
   Use FACREB Alias FACREB new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NREG                     TAG ORDEN1 to FACREB.CDX EVAL PORCENT(7,1,2,NLAST2,NomVen1)
   Index on UPPER(NOMCTA)+DTOS(FREG) TAG ORDEN2 to FACREB.CDX EVAL PORCENT(7,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="PAGOS" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "PAGOS.DBF" )
      Crear("PAGOS.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("PAGOS.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\PAGOS.DBF")
      ELSE
         Crear("PAGOS.DBF")
         Use PAGOS Alias PAGOS
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('PAGOS.CDX')
   Use PAGOS Alias PAGOS new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NREG TAG ORDEN1 to PAGOS.CDX EVAL PORCENT(8,1,1,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="FACVTO" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "FACVTO.DBF" )
      Crear("FACVTO.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("FACVTO.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\FACVTO.DBF")
      ELSE
         Crear("FACVTO.DBF")
         Use FACVTO Alias FACVTO
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('FACVTO.CDX')
   Use FACVTO Alias FACVTO new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on NREG TAG ORDEN1 to FACVTO.CDX EVAL PORCENT(9,1,2,NLAST2,NomVen1)
   Index on FVTO TAG ORDEN2 to FACVTO.CDX EVAL PORCENT(9,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="TOPCOD" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "TOPCOD.DBF" )
      Crear("TOPCOD.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("TOPCOD.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\TOPCOD.DBF")
      ELSE
         Crear("TOPCOD.DBF")
         Use TOPCOD Alias TOPCOD
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('TOPCOD.CDX')
   Use TOPCOD Alias TOPCOD new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on (GRUPO*1000)+COD TAG ORDEN1 to TOPCOD.CDX EVAL PORCENT(10,1,2,NLAST2,NomVen1)
   Index on UPPER(NOMCOD)    TAG ORDEN2 to TOPCOD.CDX EVAL PORCENT(10,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="PGC" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( RUTAPROGRAMA+"\PGC.DBF" )
      Crear(RUTAPROGRAMA+"\PGC.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME(RUTAPROGRAMA+"\PGC.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 RUTAPROGRAMA+"\PGC.DBF")
      ELSE
         Crear(RUTAPROGRAMA+"\PGC.DBF")
         Use &RUTAPROGRAMA\PGC Alias PGC
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE(RUTAPROGRAMA+"\PGC.CDX")
   Use &RUTAPROGRAMA\PGC Alias PGC new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on CODPGC         TAG ORDEN1 to &RUTAPROGRAMA\PGC.CDX EVAL PORCENT(11,1,2,NLAST2,NomVen1)
   Index on UPPER(NOMPGC1) TAG ORDEN2 to &RUTAPROGRAMA\PGC.CDX EVAL PORCENT(11,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF

IF UPPER(NomFic1)=="REMESA" .OR. UPPER(NomFic1)=="TODOS"
   If .NOT. FILE( "REMESA.DBF" )
      Crear("REMESA.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("REMESA.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\REMESA.DBF")
      ELSE
         Crear("REMESA.DBF")
         Use REMESA Alias REMESA
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('REMESA.CDX')
   Use REMESA Alias REMESA new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on (SERIE*100000)+NREM TAG ORDEN1 to REMESA.CDX EVAL PORCENT(12,1,2,NLAST2,NomVen1)
   CLOSE DATABASES

   If .NOT. FILE( "REMLIN.DBF" )
      Crear("REMLIN.DBF")
   else
      CLOSE DATABASES
      ERASE &RUTAPROGRAMA\FIN.DBF
      IF FRENAME("REMLIN.DBF" , RUTAPROGRAMA+"\FIN.DBF" )<>0 //-1 = error
         MSGSTOP("Error en la apertura del fichero: "+HB_OsNewLine()+;
                 GetCurrentFolder()+"\REMLIN.DBF")
      ELSE
         Crear("REMLIN.DBF")
         Use REMLIN Alias REMLIN
         APPEND FROM &RUTAPROGRAMA\FIN
         USE
      ENDIF
   EndIf
   FERASE('REMLIN.CDX')
   Use REMLIN Alias REMLIN new EXCLUSIVE
   *PACK
   NLAST2:=LASTREC()
   Index on STR(SERIE,2)+STR(NREM,5)+STR(LREM,3) TAG ORDEN1 to REMLIN.CDX EVAL PORCENT(12,2,2,NLAST2,NomVen1)
   CLOSE DATABASES
ENDIF



RETURN NIL


FUNCTION PORCENT(LIN2,FICN,FICT,NLAST2,NomVen1)
   LOCAL NUM2:=(RECNO()/NLAST2)*100
   NUM2=INT((NUM2/FICT)+((100/FICT)*(FICN-1)))
   IF UPPER(NomVen1)<>"SINVENTANA"
      CONTROL1:="Progress_"+LTRIM(STR(LIN2))
      IF IsWIndowDefined(&NomVen1)=.T.
         IF IsControlDefined(&CONTROL1,&NomVen1)=.T.
            SetProperty (NomVen1,CONTROL1,"Value", NUM2 )
         ENDIF
      ENDIF
   ENDIF
RETURN (.T.)


STATIC FUNCTION W_Regfic3(LLAMADA)
   nCONTROL1:=1
   DO WHILE .T.
      CONTROL1:="Si_"+LTRIM(STR(nCONTROL1++))
      IF IsControlDefined(&CONTROL1,Regfic)=.T.
         IF LLAMADA="MARCAR"
            SetProperty ("Regfic",CONTROL1,"Value", .T. )
         ELSE
            SetProperty ("Regfic",CONTROL1,"Value", .F. )
         ENDIF
      ELSE
         EXIT
      ENDIF
   ENDDO




STATIC FUNCTION W_RegficAjustes()
***para ajustes***
CLOSE DATABASES
IF DATE()>CTOD("01-05-2007") .AND. DATE()<CTOD("01-06-2007")
   IF MSGYESNO("¿desea actualizar codigos postales de las cuentas?")=.T.
      AbrirDBF("CUENTAS")
      IF CAMPO("CODPOS")
         IF FLOCK()=.T.
            REPLACE CODPOS WITH LEFT(PROCTA,5) FOR VAL(LEFT(PROCTA,5))>0
            REPLACE PROCTA WITH SUBSTR(PROCTA,7,50) FOR VAL(LEFT(PROCTA,5))>0
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
      ENDIF
   ENDIF
ENDIF

RETURN
