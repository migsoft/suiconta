#include "minigui.ch"

PROCEDURE PagAgr()

   DEFINE WINDOW WinPagAgr ;
      AT 10,10 ;
      WIDTH 620 HEIGHT 340 ;
      TITLE 'Agrupar saldos pendiente facturas' ;
      MODAL  ;
      NOSIZE ;
      ON RELEASE CloseTables()

      @ 10,10 GRID BR_FacL1 ;
      HEIGHT 200 ;
      WIDTH 600 ;
      HEADERS {'Ruta','Empresa','Recno','Registro','Fecha','Codigo','Nombre','Factura','Importe','Pendiente'} ;
      WIDTHS { 0,0,0,75,80,0,200,60,80,80 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','DATE'}, ;
         {'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','NUMERIC'}} ;
      JUSTIFY {BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT, ;
			BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE 1

      DEFINE CONTEXT MENU CONTROL BR_FacL1 OF WinPagAgr
         ITEM "Añadir linea"        ACTION PagAgr_Nuevo()
         ITEM "Eliminar linea"      ACTION PagAgr_Eliminar()
      END MENU

      @230,010 BUTTONEX Bt_Nuevo CAPTION 'Añadir' ICON icobus('nuevo') ;
         ACTION PagAgr_Nuevo() WIDTH 90 HEIGHT 25 NOTABSTOP

      @230,110 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
         ACTION PagAgr_Eliminar() WIDTH 90 HEIGHT 25 NOTABSTOP

      @235,350 LABEL   L_Total VALUE 'Pendiente' AUTOSIZE TRANSPARENT
      @230,410 TEXTBOX T_Total WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


      @285,010 LABEL L_FecTras VALUE 'Fecha traspaso saldo' AUTOSIZE TRANSPARENT
      @280,150 DATEPICKER D_FecTras WIDTH 100 VALUE DATE()


      @280,380 BUTTONEX Bt_Guardar CAPTION 'Traspasar saldos' ICON icobus('guardar') ;
         ACTION PagAgr_Traspasar() WIDTH 140 HEIGHT 25 NOTABSTOP

      @280,530 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
         ACTION WinPagAgr.Release WIDTH 80 HEIGHT 25 NOTABSTOP


   END WINDOW
   VentanaCentrar("WinPagAgr","Ventana1")
   ACTIVATE WINDOW WinPagAgr


STATIC FUNCTION PagAgr_Nuevo()
	IF WinPagAgr.BR_FacL1.ItemCount>0
		aNfac:=WinPagAgr.BR_FacL1.Item(WinPagAgr.BR_FacL1.ItemCount)
		NREG2:=aNfac[4]
		NUMEMP2:=aNfac[2]
	ELSE
		NREG2:=0
		NUMEMP2:=NUMEMP
	ENDIF
   aNFAC:=br_recbus(RUTAPROGRAMA,NUMEMP2,NREG2)

	IF aNFAC[1]=0
		RETURN
	ENDIF

   AbrirDBF("FACREB",,,,aNFAC[2],1)
   GO aNFAC[1]
	IF .NOT. EOF()
		WinPagAgr.BR_FacL1.AddItem({aNFAC[2],aNFAC[3],aNFAC[1],NREG,FREG,CODIGO,NOMCTA,REF,TFAC,PEND})
		WinPagAgr.T_Total.Value:=WinPagAgr.T_Total.Value+PEND
	ENDIF


STATIC FUNCTION PagAgr_Eliminar()
	IF WinPagAgr.BR_FacL1.Value>=1
		NESTOY:=WinPagAgr.BR_FacL1.Value
		WinPagAgr.T_Total.Value:=WinPagAgr.T_Total.Value-WinPagAgr.BR_FacL1.Cell(WinPagAgr.BR_FacL1.Value,10)
		WinPagAgr.BR_FacL1.DeleteItem(WinPagAgr.BR_FacL1.Value)
		IF NESTOY>=WinPagAgr.BR_FacL1.ItemCount
			WinPagAgr.BR_FacL1.Value:=WinPagAgr.BR_FacL1.ItemCount
		ELSE
			WinPagAgr.BR_FacL1.Value:=NESTOY
		ENDIF
	ENDIF


STATIC FUNCTION PagAgr_Traspasar()
PonerEspera("Procesando los datos....")
	FOR N=1 TO WinPagAgr.BR_FacL1.ItemCount-1
		aNfac2:=WinPagAgr.BR_FacL1.Item(N)
		NREG2:=aNfac2[4]

		aNfac3:=WinPagAgr.BR_FacL1.Item(WinPagAgr.BR_FacL1.ItemCount)
		NREG3:=aNfac3[4]

	   AbrirDBF("FACREB",,,,aNfac2[1],1)
      SEEK NREG2
		IF .NOT. EOF()
		   IF RLOCK()
		      REPLACE PEND WITH PEND-aNfac2[10]
		      DBCOMMIT()
		      DBUNLOCK()
		   ENDIF
		ENDIF
	   AbrirDBF("PAGOS",,,,aNfac2[1],1)
      APPEND BLANK
	   IF RLOCK()
	      REPLACE NREG WITH NREG2
	      REPLACE FREG WITH aNfac2[5]
	      REPLACE FPAG WITH WinPagAgr.D_FecTras.Value
	      REPLACE IMPORTE WITH aNfac2[10]
	      REPLACE DESCRIP WITH "Saldo a factura "+aNfac3[8]
	      REPLACE NEMP WITH aNfac2[2]
	      DBCOMMIT()
	      DBUNLOCK()
	   ENDIF

	   AbrirDBF("FACREB",,,,aNfac3[1],1)
      SEEK NREG3
		IF .NOT. EOF()
		   IF RLOCK()
		      REPLACE PEND WITH PEND+aNfac2[9]
		      DBCOMMIT()
		      DBUNLOCK()
		   ENDIF
		ENDIF
	   AbrirDBF("PAGOS",,,,aNfac3[1],1)
      APPEND BLANK
	   IF RLOCK()
	      REPLACE NREG WITH NREG3
	      REPLACE FREG WITH aNfac3[5]
	      REPLACE FPAG WITH WinPagAgr.D_FecTras.Value
	      REPLACE IMPORTE WITH aNfac2[9]*-1
	      REPLACE DESCRIP WITH "Saldo de factura "+aNfac2[8]
	      REPLACE NEMP WITH aNfac3[2]
	      DBCOMMIT()
	      DBUNLOCK()
	   ENDIF
	NEXT
QuitarEspera()
MsgInfo("Saldos agrupados")
