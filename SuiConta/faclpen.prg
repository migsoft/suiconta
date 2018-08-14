#include "minigui.ch"
#include "winprint.ch"

procedure FaclPen()
   TituloImp:="Listado de facturas emitidas pendiente"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 330 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
      @ 15,215 LABEL L_Fec1b VALUE 'A�o = ejercicios anteriores' AUTOSIZE TRANSPARENT

      @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2
      @ 45,215 LABEL L_Fec2b VALUE 'A�o = ejercicios posteriores' AUTOSIZE TRANSPARENT

      @ 75,10 LABEL L_OrdLis ;
              VALUE 'Ordenar por:' ;
              WIDTH 90 HEIGHT 25
      @ 70,110 COMBOBOX C_OrdLis ;
              WIDTH 150 ;
              ITEMS {'Fecha de factura','Codigo cuenta','Nombre cuenta'} ;
              VALUE 1

draw rectangle in window W_Imp1 at 170,010 to 172,390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @185,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @180,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @215,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @210,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @210, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @240, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @270, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION FaclPeni()

      @270,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION FaclPeni()
   local oprint

   IF FILE("FIN.DBF")
      AbrirDBF("Fin","SIN_INDICE")
      FIN->( DBCLOSEAREA() )
      ERASE FIN.DBF
      ERASE FIN.CDX
   ENDIF

   AbrirDBF("FAC92")
   COPY STRUC TO FIN
   FAC92->( DBCLOSEAREA() )
   AbrirDBF("Fin","SIN_INDICE")

   FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
      RUTA2:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
      RUTA2:=RUTA2[1]+"\FAC92.DBF"
      IF .NOT. FILE(RUTA2)
         LOOP
      ENDIF
      APPEND FROM &RUTA2 FOR PEND<>0 .AND. ;
      FFAC>=W_Imp1.D_Fec1.value .AND. FFAC<=W_Imp1.D_Fec2.value
   NEXT

   DO CASE
   CASE W_Imp1.C_OrdLis.value=1
      INDEX ON DTOS(FFAC)+SERFAC+STR(NFAC,10) TO FIN
   CASE W_Imp1.C_OrdLis.value=2
      INDEX ON STR(COD,10)+SERFAC+STR(NFAC,10) TO FIN
   CASE W_Imp1.C_OrdLis.value=3
      INDEX ON CLIENTE+SERFAC+STR(NFAC,10) TO FIN
   ENDCASE

   GO TOP
   IF LASTREC()=0
      MsgExclamation("No hay datos en las fechas introducidas","Informacion")
      FIN->( DBCLOSEAREA() )
      RETURN
   ENDIF

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloImp)
   oprint:setpreviewsize(2)  // tama�o del preview
   oprint:beginpage()

PAG:=0
LIN:=0
TOT1:=0
TOT2:=0
DO WHILE .NOT. EOF()
   IF LIN>=260 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(LIN,100,"Suma","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,140,MIL(TOT1,12,2),"times new roman",10,.F.,,"R",)
         oprint:printdata(LIN,160,MIL(TOT2,12,2),"times new roman",10,.F.,,"R",)
         LIN:=LIN+5
         oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

      oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
      oprint:printdata(45,20,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

      oprint:printdata(40,150,'Ordenado por:',"times new roman",12,.F.,,"L",)
      oprint:printdata(45,150,W_Imp1.C_OrdLis.DisplayValue,"times new roman",12,.F.,,"L",)

      LIN:=55
      oprint:printdata(LIN, 30,"Factura","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 32,"Fecha","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN, 60,"Codigo","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 62,"Nombre","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,140,"Total fac.","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,160,"Pendiente","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,162,"Forma pago","times new roman",10,.F.,,"L",)
      oprint:printline(LIN+4,20,LIN+4,180,,0.5)

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN, 30,MIL(NFAC,10,0)+"-"+SERFAC,"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 32,DIA(FFAC,10),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN, 60,COD,"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 62,CLIENTE,"times new roman",10,.F.,,"L",)

   oprint:printdata(LIN,140,MIL(TFAC,12,2),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,160,MIL(PEND,12,2),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,162,VENCI(1,FPAGO),"times new roman",10,.F.,,"L",)

   LIN:=LIN+5
   TOT1:=TOT1+TFAC
   TOT2:=TOT2+PEND
   SKIP

ENDDO

   oprint:printdata(LIN,100,"Total","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,140,MIL(TOT1,12,2),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,160,MIL(TOT2,12,2),"times new roman",10,.F.,,"R",)

   SELEC FIN
   FIN->( DBCLOSEAREA() )

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil




