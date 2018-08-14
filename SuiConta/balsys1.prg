#include "minigui.ch"
#include "winprint.ch"

procedure balsys1()
   TituloImp:="Balance de sumas y saldos cuenta"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 270 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_CodCta1 VALUE 'Codigo cuenta' AUTOSIZE TRANSPARENT
      @ 10,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON CHANGE balsys1_Cuenta() ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @ 10,225 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP
      @ 40,10 TEXTBOX T_NomCta WIDTH 300 VALUE ''

      @ 70,10 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

draw rectangle in window W_Imp1 at 110,010 to 112,390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @125,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @120,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @155,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @150,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @150, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @180, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @210, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION balsys1i("IMPRESORA")

      @210,110 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 90 HEIGHT 25 ;
               ACTION balsys1i("CALC")

      @210,210 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION balsys1_Cuenta()
   AbrirDBF("CUENTAS",,,,,1)
   SEEK VAL(W_Imp1.T_CodCta1.Value)
   IF .NOT. EOF()
      W_Imp1.T_NomCta.Value:=RTRIM(NOMCTA)
   ELSE
      W_Imp1.T_NomCta.Value:=""
   ENDIF




STATIC FUNCTION balsys1i(LLAMADA)
   local oprint

aSUMCTA1:={}
aSUMCTA2:={}
AADD(aSUMCTA1,{DIAINIMES(DMA),0,0})

AbrirDBF("APUNTES")
W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
GO TOP
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS
   IF CODCTA=VAL(W_Imp1.T_CodCta1.Value)
      IF NASI=1
         aSUMCTA1[1,2]=aSUMCTA1[1,2]+DEBE
         aSUMCTA1[1,3]=aSUMCTA1[1,3]+HABER
      ELSE
         FASI2:=DIAINIMES(FECHA)
         NUM2:=ASCAN(aSUMCTA2,{ |AVAL| AVAL[1]==FASI2 })
         IF NUM2=0
            AADD(aSUMCTA2,{FASI2,DEBE,HABER})
         ELSE
            aSUMCTA2[NUM2,2]=aSUMCTA2[NUM2,2]+DEBE
            aSUMCTA2[NUM2,3]=aSUMCTA2[NUM2,3]+HABER
         ENDIF
      ENDIF
   ENDIF
   SKIP
ENDDO

ASORT(aSUMCTA2,,,{|X,Y| DTOS(X[1])<DTOS(Y[1])})

IF LLAMADA="IMPRESORA"
   balsys1i_Imp()
ELSE
   balsys1i_Calc()
ENDIF




STATIC FUNCTION balsys1i_Imp()
   local oprint

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

      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

      oprint:printdata(40,30,'Cuenta: '+W_Imp1.T_CodCta1.value+" "+W_Imp1.T_NomCta.Value,"times new roman",12,.F.,,"L",)


   LIN:=55

   oprint:printdata(LIN, 90,"Debe" ,"times new roman",12,.F.,,"R",)
   oprint:printdata(LIN,125,"Haber","times new roman",12,.F.,,"R",)
   oprint:printdata(LIN,160,"Saldo","times new roman",12,.F.,,"R",)
   oprint:printline(LIN+4,20,LIN+4,160,,0.5)
   LIN:=LIN+5

   oprint:printdata(LIN, 20,"Saldo Inicial","times new roman",12,.F.,,"L",)
   oprint:printdata(LIN, 90,MIL(aSUMCTA1[1,2],13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   oprint:printdata(LIN,125,MIL(aSUMCTA1[1,3],13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   SALDO2:=0
   TOTDEB2:=0
   TOTHAB2:=0
   TOTDEB2:=aSUMCTA1[1,2]
   TOTHAB2:=aSUMCTA1[1,3]
   SALDO2:=TOTDEB2-TOTHAB2
   oprint:printdata(LIN,160,MIL(SALDO2,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)

   LIN:=LIN+5

IF LEN(aSUMCTA2)=0
   FEC1:=DMA
   FEC2:=DMA2
ELSE
   FEC1:=aSUMCTA2[1,1]
   FEC2:=MAX(aSUMCTA2[LEN(aSUMCTA2),1] , DMA2 )
ENDIF
DO WHILE FEC1<FEC2
   NUM2:=ASCAN(aSUMCTA2,{ |AVAL| AVAL[1]==DIAINIMES(FEC1) })
   oprint:printdata(LIN, 20,RTRIM(MES(MONTH(FEC1)))+"-"+STR(YEAR(FEC1),4),"times new roman",12,.F.,,"L",)
   IF NUM2=0
      oprint:printdata(LIN, 90,MIL( 0 ,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
      oprint:printdata(LIN,125,MIL( 0 ,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN, 90,MIL(aSUMCTA2[NUM2,2],13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
      oprint:printdata(LIN,125,MIL(aSUMCTA2[NUM2,3],13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
      TOTDEB2:=TOTDEB2+aSUMCTA2[NUM2,2]
      TOTHAB2:=TOTHAB2+aSUMCTA2[NUM2,3]
      SALDO2:=TOTDEB2-TOTHAB2
   ENDIF
   oprint:printdata(LIN,160,MIL(SALDO2,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   FEC1:=DIAINIMES(FEC1+35)
   LIN:=LIN+5
ENDDO
   oprint:printline(LIN-1,20,LIN-1,160,,0.5)
   oprint:printdata(LIN, 20,"Total","times new roman",12,.F.,,"L",)
   oprint:printdata(LIN, 90,MIL(TOTDEB2,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   oprint:printdata(LIN,125,MIL(TOTHAB2,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)
   oprint:printdata(LIN,160,MIL(SALDO2,13,MDA_DEC(EJERANY)),"times new roman",12,.F.,,"R",)


   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil



STATIC FUNCTION balsys1i_Calc()
   local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

   PonerEspera("Procesando datos del listado")

   // inicializa
   oServiceManager := TOleAuto():New("com.sun.star.ServiceManager")
   oDesktop := oServiceManager:createInstance("com.sun.star.frame.Desktop")
   IF oDesktop = NIL
      QuitarEspera()
      MsgStop('OpenOffice Calc no esta disponible','error')
      RETURN Nil
   ENDIF
   oDocument := oDesktop:loadComponentFromURL("private:factory/scalc","_blank", 0, {})

   // tomar hoja
   oSchedule := oDocument:GetSheets()

   // tomar primera hoja por nombre oSheet := oSchedule:GetByName("Hoja1")
   // o por indice
   oSheet := oSchedule:GetByIndex(0)

   // escribir en las celdas
   LIN:=6
   oSheet:getCellByPosition(0,LIN):SetString( "Mes" ) // columna,linea
   oSheet:getCellByPosition(1,LIN):SetString( "Debe" ) // columna,linea
   oSheet:getCellByPosition(2,LIN):SetString( "Haber" ) // columna,linea
   oSheet:getCellByPosition(3,LIN):SetString( "Saldo" ) // columna,linea
   oSheet:getCellRangeByPosition(1,LIN,3,LIN):HoriJustify:=3
   oSheet:getCellRangeByPosition(0,LIN,3,LIN):CharWeight:=150 //NEGRITA
   aMiColor:=MiColor("AMARILLOPALIDO")
   oSheet:getCellRangeByPosition(0,LIN,3,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
   LIN++
   oSheet:getCellByPosition(0,LIN):SetString( "Saldo Inicial" ) // columna,linea
   oSheet:getCellByPosition(1,LIN):SetValue( aSUMCTA1[1,2] ) // columna,linea
   oSheet:getCellByPosition(2,LIN):SetValue( aSUMCTA1[1,3] ) // columna,linea
   oSheet:getCellByPosition(3,LIN):SetFormula( "=+B"+LTRIM(STR(LIN+1))+"-C"+LTRIM(STR(LIN+1)) )
   oSheet:getCellRangeByPosition(1,LIN,3,LIN):NumberFormat:=4 //#.##0,00
   LIN++
   LIN1:=LIN
   FOR N1=1 TO LEN(aSUMCTA2)
      oSheet:getCellByPosition(0,LIN):SetString( MES(MONTH(aSUMCTA2[N1,1])) ) // columna,linea
      oSheet:getCellByPosition(1,LIN):SetValue( aSUMCTA2[N1,2] ) // columna,linea
      oSheet:getCellByPosition(2,LIN):SetValue( aSUMCTA2[N1,3] ) // columna,linea
   oSheet:getCellByPosition(3,LIN):SetFormula( "=+B"+LTRIM(STR(LIN+1))+"-C"+LTRIM(STR(LIN+1))+"+D"+LTRIM(STR(LIN)) )
      oSheet:getCellRangeByPosition(1,LIN,3,LIN):NumberFormat:=4 //#.##0,00
      LIN++
   NEXT

   oSheet:getCellByPosition(0,LIN):SetString("TOTAL") 
   FOR N1=1 TO 2
      oSheet:getCellByPosition(N1,LIN):SetFormula("=SUM("+LetraExcel(N1+1)+LTRIM(STR(LIN1))+":"+LetraExcel(N1+1)+LTRIM(STR(LIN))+")")
   NEXT
   oSheet:getCellByPosition(3,LIN):SetFormula( "=+B"+LTRIM(STR(LIN+1))+"-C"+LTRIM(STR(LIN+1)) )
   oSheet:getCellRangeByPosition(1,LIN,3,LIN):NumberFormat:=4 //#.##0,00
   oSheet:getCellRangeByPosition(0,LIN,3,LIN):CharWeight:=150 //NEGRITA
   aMiColor:=MiColor("VERDEPALIDO")
   oSheet:getCellRangeByPosition(0,LIN,3,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

   oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)

   oSheet:getCellByPosition(0,0):SetString(Nombre_Programa())
   oSheet:getCellByPosition(0,1):SetString(DIA(DATE(),10))
   oSheet:getCellByPosition(0,2):SetString(NOMEMPRESA)
   oSheet:getCellByPosition(0,3):SetString(TituloImp)
   oSheet:getCellByPosition(0,3):CharHeight:=16
   oSheet:getCellByPosition(0,4):SetString('Cuenta:')
   oSheet:getCellByPosition(1,4):SetString(W_Imp1.T_CodCta1.value+" "+W_Imp1.T_NomCta.Value)
   oSheet:getCellByPosition(1,4):HoriJustify:=1

   QuitarEspera()

   W_Imp1.release

return nil 




