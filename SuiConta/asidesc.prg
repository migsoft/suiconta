#include "minigui.ch"
#include "winprint.ch"

procedure AsiDesc()
   TituloImp:="Listado asientos descuadrados"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 800 HEIGHT 330 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @010,10 CHECKBOX Si_Apuntes CAPTION 'Fichero de apuntes' WIDTH 150 VALUE .T.
      @040,10 PROGRESSBAR P_Progres1 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

      @070,10 CHECKBOX Si_Cierre  CAPTION 'Fichero de cierre'  WIDTH 150 VALUE .T.
      @100,10 PROGRESSBAR P_Progres2 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

      @010,410 GRID GR_Apu1 ;
      HEIGHT 250 ;
      WIDTH 380 ;
      HEADERS {'Apunte','Fecha','Descuadre','Importe','Fichero'} ;
      WIDTHS {50,70,130,100,50 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','NUMERIC'  },{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
                      {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','CHARACTER'}} ;
      JUSTIFY {BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT, ;
               BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT} ;
      VALUE 1 ;
      ON DBLCLICK br_suizoasi(RUTAEMPRESA,W_Imp1.GR_Apu1.Cell(W_Imp1.GR_Apu1.Value,1),NUMEMP)

      @270,450 LABEL L_Total VALUE 'Total' AUTOSIZE TRANSPARENT


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

      @270,010 BUTTONEX B_Bus CAPTION 'Buscar' ICON icobus('buscar') WIDTH 90 HEIGHT 25 ;
               ACTION ApuDescb()

      @270,110 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION ApuDesci()

      @270,210 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      @270,310 BUTTONEX B_Sal CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION ApuDescb()
   W_Imp1.GR_Apu1.DeleteAllItems
FOR N1=1 TO 2
   IF N1=1
      IF W_Imp1.Si_Apuntes.Value=.T.
         NOMFICAPU:="APUNTES"
      ELSE
         LOOP
      ENDIF
   ELSE
      IF W_Imp1.Si_Cierre.Value=.T.
         NOMFICAPU:="CIERRE"
      ELSE
         LOOP
      ENDIF
   ENDIF
   AbrirDBF(NOMFICAPU)
   nProgres:=1
   GO TOP
   DO WHILE .NOT. EOF()
      DO EVENTS
      NASI2:=NASI
      FASI2:=FECHA
      DEB2:=0
      HAB2:=0
      APU2:=0
      DO WHILE NASI2=NASI .AND. .NOT. EOF()
         IF N1=1
            W_Imp1.P_Progres1.Value:=(nProgres++*100)/LASTREC()
         ELSE
            W_Imp1.P_Progres2.Value:=(nProgres++*100)/LASTREC()
         ENDIF
         DEB2:=DEB2+DEBE
         HAB2:=HAB2+HABER
         IF YEAR(FECHA)<>EJERANY .AND. APU2=0 //FECHA FUERA DEL EJERCICIO
            APU2:=1
            W_Imp1.GR_Apu1.AddItem({NASI2,DIA(FASI2,8),"Fecha fuera del ejercicio",0,IF(N1=1,"Apuntes","Cierre") })
         ENDIF
         CODCTA2=CODCTA
         IF LEN(LTRIM(STR(CODCTA2)))<>8 //CUENTA DISTINTA A 8 DIGITOS
            W_Imp1.GR_Apu1.AddItem({NASI2,DIA(FASI2,8),LTRIM(STR(CODCTA2))+" cuenta erronea",0,IF(N1=1,"Apuntes","Cierre") })
         ELSE
            AbrirDBF("CUENTAS")
            SEEK CODCTA2
            IF EOF()   //CUENTA INEXISTENTE
               W_Imp1.GR_Apu1.AddItem({NASI2,DIA(FASI2,8),LTRIM(STR(CODCTA2))+" cuenta inexistente",0,IF(N1=1,"Apuntes","Cierre") })
            ENDIF
         ENDIF
         AbrirDBF(NOMFICAPU)
         SKIP
      ENDDO
      IF STRZERO(DEB2-HAB2,15,4)<>STRZERO( 0 ,15,4) //ASIENTO DESCUADRADO
         W_Imp1.GR_Apu1.AddItem({NASI2,DIA(FASI2,8),"Importe descuadrado",DEB2-HAB2,IF(N1=1,"Apuntes","Cierre") })
      ENDIF
   ENDDO
NEXT
W_Imp1.L_Total.Value:='Total asientos descuadrados '+LTRIM(MIL(W_Imp1.GR_Apu1.ItemCount,12,0))





STATIC FUNCTION ApuDesci()
   local oprint

   GO TOP
   IF W_Imp1.GR_Apu1.ItemCount=0
      MsgExclamation("No hay asientos descuadrados","Informacion")
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
   oprint:setpreviewsize(2)  // tamaño del preview
   oprint:beginpage()

PAG:=0
LIN:=0
TOT1:=0
FOR N1=1 TO W_Imp1.GR_Apu1.ItemCount
   IF LIN>=260 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(LIN,110,"Suma","times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,150,MIL(TOT1,15,2),"times new roman",10,.F.,,"R",)
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

      LIN:=40
      oprint:printdata(LIN, 35,"Asiento","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN, 40,"Fecha","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN, 65,"Descuadre","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,150,"Importe","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,152,"Fichero","times new roman",10,.F.,,"L",)
      oprint:printline(LIN+4,20,LIN+4,170,,0.5)

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN, 35,MIL(W_Imp1.GR_Apu1.Cell(N1,1),10,0),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN, 40,DIA(CTOD(W_Imp1.GR_Apu1.Cell(N1,2)),10),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN, 65,W_Imp1.GR_Apu1.Cell(N1,3),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,150,MIL(W_Imp1.GR_Apu1.Cell(N1,4),15,2),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,152,W_Imp1.GR_Apu1.Cell(N1,5),"times new roman",10,.F.,,"L",)

   TOT1:=TOT1+W_Imp1.GR_Apu1.Cell(N1,4)

   LIN:=LIN+5

NEXT

   oprint:printdata(LIN,020,"Asientos descuadrados:","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,055,LTRIM(MIL(W_Imp1.GR_Apu1.ItemCount,15,0)),"times new roman",10,.F.,,"L",)

   oprint:printdata(LIN,110,"Total","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN,150,MIL(TOT1,15,2),"times new roman",10,.F.,,"R",)

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil





