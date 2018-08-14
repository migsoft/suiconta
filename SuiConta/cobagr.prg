#include "minigui.ch"

PROCEDURE CobAgr()

   DEFINE WINDOW WinCobAgr ;
      AT 10,10 ;
      WIDTH 620 HEIGHT 340 ;
      TITLE 'Agrupar saldos pendiente facturas' ;
      MODAL  ;
      NOSIZE ;
      ON RELEASE CloseTables()

      @ 10,10 GRID BR_FacL1 ;
      HEIGHT 200 ;
      WIDTH 600 ;
      HEADERS {'Ruta','Empresa','Recno','Factura','Fecha','Codigo','Nombre','Importe','Pendiente'} ;
      WIDTHS { 0,0,0,75,80,0,200,80,80 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','DATE'}, ;
         {'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','NUMERIC'}} ;
      JUSTIFY {BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT, ;
			BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE 1

      DEFINE CONTEXT MENU CONTROL BR_FacL1 OF WinCobAgr
         ITEM "Añadir linea"        ACTION CobAgr_Nuevo()
         ITEM "Eliminar linea"      ACTION CobAgr_Eliminar()
      END MENU

      @230,010 BUTTONEX Bt_Nuevo CAPTION 'Añadir' ICON icobus('nuevo') ;
         ACTION CobAgr_Nuevo() WIDTH 90 HEIGHT 25 NOTABSTOP

      @230,110 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
         ACTION CobAgr_Eliminar() WIDTH 90 HEIGHT 25 NOTABSTOP

      @235,350 LABEL   L_Total VALUE 'Pendiente' AUTOSIZE TRANSPARENT
      @230,410 TEXTBOX T_Total WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


      @285,010 LABEL L_FecTras VALUE 'Fecha traspaso saldo' AUTOSIZE TRANSPARENT
      @280,150 DATEPICKER D_FecTras WIDTH 100 VALUE DATE()


      @280,380 BUTTONEX Bt_Guardar CAPTION 'Traspasar saldos' ICON icobus('guardar') ;
         ACTION CobAgr_Traspasar() WIDTH 140 HEIGHT 25 NOTABSTOP

      @280,530 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
         ACTION WinCobAgr.Release WIDTH 80 HEIGHT 25 NOTABSTOP


   END WINDOW
   VentanaCentrar("WinCobAgr","Ventana1")
   ACTIVATE WINDOW WinCobAgr


STATIC FUNCTION CobAgr_Nuevo()
	IF WinCobAgr.BR_FacL1.ItemCount>0
		aNfac:=WinCobAgr.BR_FacL1.Item(WinCobAgr.BR_FacL1.ItemCount)
		SFAC2:=SUBSTR(aNfac[4],AT("-",aNfac[4])+1,1)
		NFAC2:=VAL(LEFT(aNfac[4],AT("-",aNfac[4])-1))
		NUMEMP2:=aNfac[2]
	ELSE
		SFAC2:="A"
		NFAC2:=0
		NUMEMP2:=NUMEMP
	ENDIF
   aNFAC:=br_emibus(RUTAPROGRAMA,NUMEMP2,SFAC2,NFAC2)

	IF aNFAC[1]=0
		RETURN
	ENDIF

   AbrirDBF("FAC92",,,,aNFAC[2],1)
   GO aNFAC[1]
	IF .NOT. EOF()
		WinCobAgr.BR_FacL1.AddItem({aNFAC[2],aNFAC[3],aNFAC[1],LTRIM(STR(NFAC))+"-"+SERFAC,FFAC,COD,CLIENTE,TFAC,PEND})
		WinCobAgr.T_Total.Value:=WinCobAgr.T_Total.Value+PEND
	ENDIF


STATIC FUNCTION CobAgr_Eliminar()
	IF WinCobAgr.BR_FacL1.Value>=1
		NESTOY:=WinCobAgr.BR_FacL1.Value
		WinCobAgr.T_Total.Value:=WinCobAgr.T_Total.Value-WinCobAgr.BR_FacL1.Cell(WinCobAgr.BR_FacL1.Value,9)
		WinCobAgr.BR_FacL1.DeleteItem(WinCobAgr.BR_FacL1.Value)
		IF NESTOY>=WinCobAgr.BR_FacL1.ItemCount
			WinCobAgr.BR_FacL1.Value:=WinCobAgr.BR_FacL1.ItemCount
		ELSE
			WinCobAgr.BR_FacL1.Value:=NESTOY
		ENDIF
	ENDIF


STATIC FUNCTION CobAgr_Traspasar()
PonerEspera("Procesando los datos....")
	FOR N=1 TO WinCobAgr.BR_FacL1.ItemCount-1
		aNfac2:=WinCobAgr.BR_FacL1.Item(N)
		SFAC2:=SUBSTR(aNfac2[4],AT("-",aNfac2[4])+1,1)
		NFAC2:=VAL(LEFT(aNfac2[4],AT("-",aNfac2[4])-1))

		aNfac3:=WinCobAgr.BR_FacL1.Item(WinCobAgr.BR_FacL1.ItemCount)
		SFAC3:=SUBSTR(aNfac3[4],AT("-",aNfac3[4])+1,1)
		NFAC3:=VAL(LEFT(aNfac3[4],AT("-",aNfac3[4])-1))

	   AbrirDBF("fac92",,,,aNfac2[1],1)
      SEEK SERIE(SFAC2,NFAC2)
		IF .NOT. EOF()
		   IF RLOCK()
		      REPLACE PEND WITH PEND-aNfac2[9]
		      DBCOMMIT()
		      DBUNLOCK()
		   ENDIF
		ENDIF
	   AbrirDBF("COBROS",,,,aNfac2[1],1)
      APPEND BLANK
	   IF RLOCK()
	      REPLACE SERFAC WITH SFAC2
	      REPLACE NFAC WITH NFAC2
	      REPLACE FFAC WITH aNfac2[5]
	      REPLACE FCOB WITH WinCobAgr.D_FecTras.Value
	      REPLACE IMPORTE WITH aNfac2[9]
	      REPLACE DESCRIP WITH "Saldo a factura "+LTRIM(STR(NFAC3))+"-"+SFAC3
	      REPLACE NEMP WITH aNfac2[2]
	      DBCOMMIT()
	      DBUNLOCK()
	   ENDIF

	   AbrirDBF("fac92",,,,aNfac3[1],1)
      SEEK SERIE(SFAC3,NFAC3)
		IF .NOT. EOF()
		   IF RLOCK()
		      REPLACE PEND WITH PEND+aNfac2[9]
		      DBCOMMIT()
		      DBUNLOCK()
		   ENDIF
		ENDIF
	   AbrirDBF("COBROS",,,,aNfac3[1],1)
      APPEND BLANK
	   IF RLOCK()
	      REPLACE SERFAC WITH SFAC3
	      REPLACE NFAC WITH NFAC3
	      REPLACE FFAC WITH aNfac3[5]
	      REPLACE FCOB WITH WinCobAgr.D_FecTras.Value
	      REPLACE IMPORTE WITH aNfac2[9]*-1
	      REPLACE DESCRIP WITH "Saldo de factura "+LTRIM(STR(NFAC2))+"-"+SFAC2
	      REPLACE NEMP WITH aNfac3[2]
	      DBCOMMIT()
	      DBUNLOCK()
	   ENDIF
	NEXT
QuitarEspera()
MsgInfo("Saldos agrupados")
