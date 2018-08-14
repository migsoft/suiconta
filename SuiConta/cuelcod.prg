#include "minigui.ch"
#include "winprint.ch"

procedure Cuelcodigo()
   TituloImp:="Listado de cuentas"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 330 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @015,10 LABEL L_CodCta1 VALUE 'Desde el codigo' AUTOSIZE TRANSPARENT
      @010,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @010,225 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @045,10 LABEL L_CodCta2 VALUE 'Hasta el codigo' AUTOSIZE TRANSPARENT
      @040,110 TEXTBOX T_CodCta2 WIDTH 100 VALUE "99999999" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta2.Value)
      @040,225 BUTTONEX Bt_BuscarCue2 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta2.Value),"W_Imp1","T_CodCta2") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @ 75,10 LABEL L_OrdLis VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
      @ 70,110 COMBOBOX C_OrdLis WIDTH 150 ITEMS {'Codigo','Descripcion'} VALUE 1

      @100,10 CHECKBOX Si_Dir CAPTION 'Incluir direcciones' WIDTH 150 VALUE .F.

      @130,10 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH


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
               ACTION Cuelcodigoi()

      @270,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Cuelcodigoi()
   local oprint

   AbrirDBF("CUENTAS")
   IF W_Imp1.C_OrdLis.value=1
      DBSETORDER(1)
   ELSE
      DBSETORDER(2)
   ENDIF

   IF W_Imp1.Si_Dir.value=.T.
      TOTLIN:=210
      TOTCOL:=297
      PAG_HORIZ:=.T.
   ELSE
      TOTLIN:=297
      TOTCOL:=210
      PAG_HORIZ:=.F.
   ENDIF

   GO TOP
   IF LASTREC()=0
      MsgExclamation("No hay datos en los parametros introducidos","Informacion")
      RETURN
   ENDIF

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , PAG_HORIZ , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloImp)
   oprint:setpreviewsize(1)  // tamaño del preview
   oprint:beginpage()

W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
PAG:=0
LIN:=0
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS
   IF CODCTA<VAL(W_Imp1.T_CodCta1.value) .OR. CODCTA>VAL(W_Imp1.T_CodCta2.value)
      SKIP
      LOOP
   ENDIF

   IF LIN>=TOTLIN-30 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(LIN+5,TOTCOL/2,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,TOTCOL-20,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,TOTCOL/2,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,TOTCOL/2,TituloImp,"times new roman",18,.F.,,"C",)

      oprint:printdata(40,20,'Desde: '+W_Imp1.T_CodCta1.value,"times new roman",12,.F.,,"L",)
      oprint:printdata(45,20,'Hasta: '+W_Imp1.T_CodCta2.value,"times new roman",12,.F.,,"L",)

      oprint:printdata(45,100,'Ordenado por:',"times new roman",12,.F.,,"L",)
      oprint:printdata(45,130,W_Imp1.C_OrdLis.DisplayValue,"times new roman",12,.F.,,"L",)

      LIN:=55
      oprint:printdata(LIN, 35,"Codigo","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 40,"Descripcion","times new roman",10,.F.,,"L",)
      IF W_Imp1.Si_Dir.value=.T.
         oprint:printdata(LIN,130,"Direccion","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,175,"Cod.Pos.","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,190,"Poblacion","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,235,"Telefono","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,255,"Fax","times new roman",10,.F.,,"L",)
         oprint:printline(LIN+4,20,LIN+4,275,,0.5)
      ELSE
         oprint:printline(LIN+4,20,LIN+4,180,,0.5)
      ENDIF

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN, 35,CODCTA,"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 40,NOMCTA,"times new roman",10,.F.,,"L",)

   IF W_Imp1.Si_Dir.value=.T.
      oprint:printdata(LIN,130,DIRCTA,"times new roman",10,.F.,,"L",)
      IF CAMPO("CODPOS")
         oprint:printdata(LIN,175,CODPOS,"times new roman",10,.F.,,"L",)
      ENDIF
      oprint:printdata(LIN,190,POBCTA,"times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,235,TEL1,"times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,255,FAX1,"times new roman",10,.F.,,"L",)
   ENDIF

   LIN:=LIN+5
   SKIP

ENDDO

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil





