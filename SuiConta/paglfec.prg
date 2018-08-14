#include "minigui.ch"
#include "winprint.ch"

procedure Paglfec()
   TituloImp:="Listado de pagos de facturas recibidas"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 380 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
      @ 15,215 LABEL L_Fec1b VALUE 'Año = ejercicios anteriores' AUTOSIZE TRANSPARENT

      @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DATE()
      @ 45,215 LABEL L_Fec2b VALUE 'Año = ejercicios posteriores' AUTOSIZE TRANSPARENT

      @ 70,10 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Cuenta' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1","T_NomCta1")  ;
         WIDTH 90 HEIGHT 25 NOTABSTOP
      @ 70,110 TEXTBOX T_CodCta1 WIDTH 80 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @ 70,200 TEXTBOX T_NomCta1 WIDTH 190 VALUE '' READONLY

      @105,010 LABEL L_NomCob VALUE 'Descripcion' AUTOSIZE TRANSPARENT
      @100,110 TEXTBOX T_NomCob WIDTH 285 VALUE ""



      @135,10 LABEL L_OrdLis VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
      @130,110 COMBOBOX C_OrdLis WIDTH 150 ;
              ITEMS {'Fecha de pago','Fecha vencimiento','Nombre cuenta'} ;
              VALUE 1

LINW:=220
COLW:=0
draw rectangle in window W_Imp1 at LINW,COLW+10 to LINW+2,COLW+390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @LINW+15,COLW+10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @LINW+10,COLW+100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @LINW+45,COLW+220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @LINW+40,COLW+280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @LINW+40,COLW+10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @LINW+70,COLW+10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @LINW+100,COLW+10 BUTTONEX B_Imprimir CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION Paglfeci()

      @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Paglfeci()
   local oprint
	local aCobrosT:={},aCobros1:={}

   FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
      RUTA2:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
      RUTA2:=RUTA2[1]
      IF .NOT. FILE(RUTA2+"\FACREB.DBF") .OR. .NOT. FILE(RUTA2+"\PAGOS.DBF")
         LOOP
      ENDIF
		AbrirDBF("PAGOS",,,,RUTA2)
		GO TOP
		DO WHILE .NOT. EOF()
			IF FPAG>=W_Imp1.D_Fec1.value .AND. FPAG<=W_Imp1.D_Fec2.value
				IF AT(UPPER(W_Imp1.T_NomCob.value),UPPER(DESCRIP))<>0 .OR. LEN(RTRIM(W_Imp1.T_NomCob.value))=0
					aCobros1:={}
					AADD(aCobros1,NEMP)
					AADD(aCobros1,NREG)
					AADD(aCobros1,"")
					AADD(aCobros1,FPAG)
					AADD(aCobros1,DESCRIP)
					AADD(aCobros1,IMPORTE)
					AADD(aCobros1,FVTO)
					AADD(aCobros1,BANCO)
					NREG2:=NREG
					AbrirDBF("FACREB",,,,RUTA2,1)
					SEEK NREG2
					IF CODIGO=VAL(W_Imp1.T_CodCta1.value) .OR. VAL(W_Imp1.T_CodCta1.value)=0
						AADD(aCobros1,FREG)
						AADD(aCobros1,CODIGO)
						AADD(aCobros1,NOMCTA)
						AADD(aCobros1,REF)
						AADD(aCobrosT,aCobros1)
					ENDIF
					AbrirDBF("PAGOS",,,,RUTA2)
				ENDIF
			ENDIF
			SKIP
		ENDDO
   NEXT

   DO CASE
   CASE W_Imp1.C_OrdLis.value=1
	   ASORT(aCobrosT,,, { |x, y| DTOS(x[4])+UPPER(x[11]) < DTOS(y[4])+UPPER(y[11]) })
   CASE W_Imp1.C_OrdLis.value=2
	   ASORT(aCobrosT,,, { |x, y| DTOS(x[7])+UPPER(x[11]) < DTOS(y[7])+UPPER(y[11]) })
   CASE W_Imp1.C_OrdLis.value=3
	   ASORT(aCobrosT,,, { |x, y| UPPER(x[11]) < UPPER(y[11]) })
   ENDCASE

   IF LEN(aCobrosT)=0
      MsgExclamation("No hay datos en las fechas introducidas","Informacion")
      RETURN
   ENDIF

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .T. , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloImp)
   oprint:setpreviewsize(1)  // tamaño del preview
   oprint:beginpage()

PAG:=0
LIN:=0
TOT1:=0
TOT2:=0
FOR N=1 TO LEN(aCobrosT)
   IF LIN>=190 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(LIN,200,"Suma","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,240,MIL(TOT1,12,2),"times new roman",10,.F.,,"R",)
         LIN:=LIN+5
         oprint:printdata(LIN,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,280,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,150,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,150,TituloImp,"times new roman",18,.F.,,"C",)

      oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
      oprint:printdata(45,20,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

		IF VAL(W_Imp1.T_CodCta1.Value)<>0
	      oprint:printdata(40,80,'Cuenta: '+W_Imp1.T_CodCta1.Value,"times new roman",12,.F.,,"L",)
	      oprint:printdata(45,80,W_Imp1.T_NomCta1.Value,"times new roman",12,.F.,,"L",)
		ENDIF

		IF LEN(RTRIM(W_Imp1.T_NomCob.value))<>0
	      oprint:printdata(40,170,'Descripcion:',"times new roman",12,.F.,,"L",)
	      oprint:printdata(45,170,W_Imp1.T_NomCob.value,"times new roman",12,.F.,,"L",)
		ENDIF

      LIN:=55
      oprint:printdata(LIN, 35,"Registro","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 37,"Fecha","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN, 70,"Codigo","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 72,"Nombre","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,125,"Factura","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,155,"Fecha cobro","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,175,"Descripcion","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,240,"Importe","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,242,"Vencimiento","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,260,"Banco","times new roman",10,.F.,,"L",)
      oprint:printline(LIN+4,20,LIN+4,280,,0.5)

      LIN:=LIN+5
   ENDIF
/*
   ?? " EMPRESA NFRA  FECHA-FRA.  COD.  CLIENTE   FEC.COB. DESCRIPCION-COBRO  IMPORTE F.VTO. BANCO"
   ?? " EMPRESA NFRA SERIE FEC.COB. DESCRIPCION-COBRO IMPORTE F.VTO. BANCO FECHA-FRA.  COD.  CLIENTE"
*/

   oprint:printdata(LIN, 35,LTRIM(STR(aCobrosT[N,2])),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 37,DIA(aCobrosT[N,9],10),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN, 70,aCobrosT[N,10],"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 72,aCobrosT[N,11],"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,125,aCobrosT[N,12],"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,155,DIA(aCobrosT[N,4],10),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,175,aCobrosT[N,5],"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,240,MIL(aCobrosT[N,6],14,2),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,242,DIA(aCobrosT[N,7],10),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,260,aCobrosT[N,8],"times new roman",10,.F.,,"L",)

   TOT1:=TOT1+aCobrosT[N,6]

   LIN:=LIN+5

NEXT

   oprint:printdata(LIN,200,"Total","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,240,MIL(TOT1,12,2),"times new roman",10,.F.,,"R",)
/*
   SELEC FIN
   FIN->( DBCLOSEAREA() )
*/
   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil





