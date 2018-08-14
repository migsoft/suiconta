#include "minigui.ch"
Function CobRep(LLAMADA)

	IF UPPER(LLAMADA)="COBROS"
		SiCobros:=.T.
		SiPagos :=.F.
	ELSE
		SiCobros:=.F.
		SiPagos :=.T.
	ENDIF

   DEFINE WINDOW CobRep ;
      AT 0,0     ;
      WIDTH 720  ;
      HEIGHT 500 ;
      TITLE "Repasar cobros y pagos de facturas" ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("GRISCLARO") ;
      ON RELEASE CloseTables()
**      NOSYSMENU


      @ 10,10 CHECKBOX Cobros CAPTION 'Cobros de facturas emitidas' WIDTH 200 VALUE SiCobros BACKCOLOR MiColor("GRISCLARO")
      @ 10,210 PROGRESSBAR P_Cobros RANGE 0 , 100 WIDTH 400 HEIGHT 20 SMOOTH
      @ 15,620 LABEL L_FacCob VALUE '' AUTOSIZE TRANSPARENT

      @ 40,10 CHECKBOX Pagos CAPTION 'Pagos de facturas recibidas' WIDTH 200 VALUE SiPagos BACKCOLOR MiColor("GRISCLARO")
      @ 40,210 PROGRESSBAR P_Pagos RANGE 0 , 100 WIDTH 400 HEIGHT 20 SMOOTH
      @ 45,620 LABEL L_FacPag VALUE '' AUTOSIZE TRANSPARENT

      @ 70,010 GRID GR_error ;
      HEIGHT 350 ;
      WIDTH 700 ;
      HEADERS {'errores'} ;
      WIDTHS { 670 } ;
      ITEMS {}


      @440,510 BUTTON Button1 CAPTION 'Repasar' WIDTH 90 HEIGHT 25 ACTION W_CobRep2()

      @440,610 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 NOTABSTOP ;
         ACTION CobRep.Release


   END WINDOW
   VentanaCentrar("CobRep","Ventana1")
   ACTIVATE WINDOW CobRep

   Return Nil



STATIC FUNCTION W_CobRep2()

   PonerEspera("Actualizando ficheros")
IF CobRep.Cobros.Value=.T.
	W_CobRep_Cobros()
ENDIF
IF CobRep.Pagos.Value=.T.
	W_CobRep_Pagos()
ENDIF
   QuitarEspera()


RETURN NIL


*
FUNCTION W_CobRep_Cobros(VENTANA1,CONTROL1)
	DEFAULT VENTANA1 TO "CobRep"
	DEFAULT CONTROL1 TO "P_Cobros"
   IF IsWindowDefined(&VENTANA1)=.F.
		MsgStop("No existe la ventana "+VENTANA1,"error")
		RETURN
   ENDIF
   IF IsControlDefined(&CONTROL1,&VENTANA1)=.F.
		MsgStop("No existe el control "+CONTROL1,"error")
		RETURN
   ENDIF

	AbrirDBF("FAC92",,,"Exclusive")
	REPLACE PEND WITH TFAC ALL
	AbrirDBF("FAC92",,,,,1)

	AbrirDBF("COBROS",,,,,1)
   SetProperty(VENTANA1,CONTROL1,"RangeMax",LASTREC())
   SetProperty(VENTANA1,CONTROL1,"Value",0)

	SET DELETED OFF
	CONTADOR:=1
	GO TOP
	DO WHILE .NOT. EOF()
		DO EVENTS
	   SetProperty(VENTANA1,CONTROL1,"Value",CONTADOR++)
		CobRep.L_FacCob.Value:=LTRIM(STR(NFAC))+SERFAC
	   IF DELETED()=.T.
	      SKIP
	      LOOP
	   ENDIF
	   SFAC2:=SERFAC
	   NFAC2:=NFAC
	   IMPORTE2:=IMPORTE
		AbrirDBF("FAC92",,,,,1)
	   SEEK SERIE(SFAC2,NFAC2)
	   IF EOF()
	      APPEND BLANK
			IF RLOCK()
		      REPLACE SERFAC WITH SFAC2
		      REPLACE NFAC WITH NFAC2
		      REPLACE COD WITH 99999
		      REPLACE CLIENTE WITH PADC("Existen cobros",30,"*")
			   DBCOMMIT()
			   DBUNLOCK()
			ENDIF
		   IF IsControlDefined(&"GR_error",&VENTANA1)=.T.
			   SetProperty(VENTANA1,"GR_error","AddItem",LTRIM(STR(NFAC2))+SFAC2+" Factura inexistente")
		   ENDIF
	   ENDIF
		IF RLOCK()
		   REPLACE PEND WITH PEND-IMPORTE2
		   DBCOMMIT()
		   DBUNLOCK()
		ENDIF
		AbrirDBF("COBROS",,,,,1)
		SKIP
	ENDDO
	SET DELETED ON

CLOSE DATABASES




FUNCTION W_CobRep_Pagos(VENTANA1,CONTROL1)
	DEFAULT VENTANA1 TO "CobRep"
	DEFAULT CONTROL1 TO "P_Pagos"
   IF IsWindowDefined(&VENTANA1)=.F.
		MsgStop("No existe la ventana "+VENTANA1,"error")
		RETURN
   ENDIF
   IF IsControlDefined(&CONTROL1,&VENTANA1)=.F.
		MsgStop("No existe el control "+CONTROL1,"error")
		RETURN
   ENDIF

	AbrirDBF("FACREB",,,"Exclusive")
	REPLACE PEND WITH TFAC ALL
	AbrirDBF("FACREB",,,,,1)

	AbrirDBF("PAGOS",,,,,1)
   SetProperty(VENTANA1,CONTROL1,"RangeMax",LASTREC())
   SetProperty(VENTANA1,CONTROL1,"Value",0)

	SET DELETED OFF
	CONTADOR:=1
	GO TOP
	DO WHILE .NOT. EOF()
		DO EVENTS
	   SetProperty(VENTANA1,CONTROL1,"Value",CONTADOR++)
		CobRep.L_FacPag.Value:=LTRIM(STR(NREG))
	   IF DELETED()=.T.
	      SKIP
	      LOOP
	   ENDIF
	   NREG2:=NREG
	   IMPORTE2:=IMPORTE
		AbrirDBF("FACREB",,,,,1)
	   SEEK NREG2
	   IF EOF()
	      APPEND BLANK
			IF RLOCK()
		      REPLACE NREG WITH NREG2
		      REPLACE CODIGO WITH 99999
		      REPLACE NOMCTA WITH PADC("Existen pagos",30,"*")
			   DBCOMMIT()
			   DBUNLOCK()
			ENDIF
		   IF IsControlDefined(&"GR_error",&VENTANA1)=.T.
			   SetProperty(VENTANA1,"GR_error","AddItem",LTRIM(STR(NREG2))+" Factura inexistente")
		   ENDIF
	   ENDIF
		IF RLOCK()
		   REPLACE PEND WITH PEND-IMPORTE2
		   DBCOMMIT()
		   DBUNLOCK()
		ENDIF
		AbrirDBF("PAGOS",,,,,1)
		SKIP
	ENDDO
	SET DELETED ON

CLOSE DATABASES
