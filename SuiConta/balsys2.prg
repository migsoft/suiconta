#include "minigui.ch"
#include "winprint.ch"

procedure balsys2()
   TituloImp:="Balance de sumas y saldos"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 450 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_CodCta1 VALUE 'Desde el codigo' AUTOSIZE TRANSPARENT
      @ 10,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @ 10,225 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @ 45,10 LABEL L_CodCta2 VALUE 'Hasta el codigo' AUTOSIZE TRANSPARENT
      @ 40,110 TEXTBOX T_CodCta2 WIDTH 100 VALUE "99999999" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta2.Value)
      @ 40,225 BUTTONEX Bt_BuscarCue2 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta2.Value),"W_Imp1","T_CodCta2") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @ 75,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 70,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1

      @105,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @100,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

      @135,10 LABEL L_Nivel VALUE 'A nivel' AUTOSIZE TRANSPARENT
      @130,110 COMBOBOX C_Nivel WIDTH 50 ITEMS {"1","2","3","4","5","6","7","8"} VALUE 8

      @165,10 LABEL L_Incluir VALUE 'Incluir cuentas' AUTOSIZE TRANSPARENT
      @160,110 COMBOBOX C_Incluir WIDTH 125 ITEMS {"Con saldo","Con saldo cero","Sin movimientos"} VALUE 2

      @190,10 CHECKBOX C_Cierre CAPTION 'Incluir asientos de cierre' ;
            WIDTH 160 VALUE .F. TRANSPARENT NOTABSTOP

      @225,10 LABEL L_Hoja VALUE 'Hoja inicial' AUTOSIZE TRANSPARENT
      @220,110 TEXTBOX T_Hoja WIDTH 100 VALUE 1 NUMERIC RIGHTALIGN

      @250,10 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

draw rectangle in window W_Imp1 at 290,010 to 292,390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @305,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @300,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @335,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @330,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @330, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @360, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @390, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION balsys2i("IMPRESORA")

      @390,110 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 90 HEIGHT 25 ;
               ACTION balsys2i("HOJACALC")


      @390,210 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION balsys2i(LLAMADA)
   local oprint

   W_Imp1.T_CodCta1.Value:=PADR(LTRIM(W_Imp1.T_CodCta1.Value),W_Imp1.C_Nivel.Value,"0")
   W_Imp1.T_CodCta2.Value:=PADR(LTRIM(W_Imp1.T_CodCta2.Value),W_Imp1.C_Nivel.Value,"0")
   DIVI2:=VAL("1"+REPLICATE("0",8-W_Imp1.C_Nivel.Value))
   FIC_CIERRE(VAL(W_Imp1.T_CodCta1.Value),VAL(W_Imp1.T_CodCta2.Value), ;
              W_Imp1.D_Fec1.Value,W_Imp1.D_Fec2.Value, ;
              IF(W_Imp1.C_Cierre.Value=.F.,1,2),DIVI2)

   AbrirDBF1("FINSUIC",,,,gettempdir(),1)
   IF W_Imp1.C_Incluir.Value=1
      IF FLOCK()
         DELETE FOR SALDO+DEBE-HABER=0
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
   ENDIF
   IF W_Imp1.C_Incluir.Value=2 .OR. W_Imp1.C_Incluir.Value=1
      IF FLOCK()
         DELETE FOR SALDO=0 .AND. DEBE=0 .AND. HABER=0
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
   ENDIF



   GO TOP
   IF LASTREC()=0
      MsgExclamation("No hay datos en los parametros introducidos","Informacion")
      RETURN
   ENDIF


   IF LLAMADA="HOJACALC"
      balsys2i_Calc()
   ELSE
      balsys2iF()
   ENDIF


STATIC FUNCTION balsys2iF()
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

W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
PAG:=W_Imp1.T_Hoja.value-1
LIN:=55
   SALT:=0
   DEBT:=0
   HABT:=0
SELEC("FINSUIC")
GO TOP
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS

   IF LIN>=270 .OR. PAG=W_Imp1.T_Hoja.value-1
      IF PAG<>W_Imp1.T_Hoja.value-1
         oprint:printdata(LIN,090,"Sumas","times new roman",10,.F.,,"R",)
         oprint:printdata(LIN,120,MIL(SALT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
         oprint:printdata(LIN,145,MIL(DEBT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
         oprint:printdata(LIN,170,MIL(HABT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
         oprint:printdata(LIN,195,MIL(SALT+DEBT-HABT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
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

      oprint:printdata(40,30,'Desde codigo: '+W_Imp1.T_CodCta1.value,"times new roman",12,.F.,,"L",)
      oprint:printdata(45,30,'Hasta codigo: '+W_Imp1.T_CodCta2.value,"times new roman",12,.F.,,"L",)

      oprint:printdata(40,100,'Desde: '+DIA(W_Imp1.D_Fec1.Value,10),"times new roman",12,.F.,,"L",)
      oprint:printdata(45,100,'Hasta: '+DIA(W_Imp1.D_Fec2.Value,10),"times new roman",12,.F.,,"L",)

      LIN:=55
      oprint:printdata(LIN, 15,"Codigo","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN, 32,"Descripcion","times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,120,"Saldo inicial","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,145,"Debe","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,170,"Haber","times new roman",10,.F.,,"R",)
      oprint:printdata(LIN,195,"Saldo","times new roman",10,.F.,,"R",)
      oprint:printline(LIN+4,15,LIN+4,195,,0.5)

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN, 15,LTRIM(STR(CODCTA)),"times new roman",10,.F.,,"L",)
   oprint:printdata(LIN, 32,NOMCTA,"times new roman",8,.F.,,"L",)
   oprint:printdata(LIN,120,MIL(SALDO,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,145,MIL(DEBE ,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,170,MIL(HABER,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,195,MIL(SALDO+DEBE-HABER,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)

   SALT:=SALT+SALDO
   DEBT:=DEBT+DEBE
   HABT:=HABT+HABER

   LIN:=LIN+5
   SKIP

ENDDO

   oprint:printdata(LIN,090,"Total","times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,120,MIL(SALT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,145,MIL(DEBT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,170,MIL(HABT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
   oprint:printdata(LIN,195,MIL(SALT+DEBT-HABT,14,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil



STATIC FUNCTION balsys2i_Calc()
   nOption:=InputWindow("Elija el tipo de salida",{"Programa"},{1},{{"Hoja Calculo","Hoja Excel"}},300,300,.T.,{"Aceptar","Cancelar"})
   DO CASE
   CASE nOption[1] == Nil
      RETURN
   CASE nOption[1]=1
      balsys2i_Calc2()
   CASE nOption[1]=2
      balsys2i_Excel()
   ENDCASE



STATIC FUNCTION balsys2i_Calc2()
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

   oSheet:getCellByPosition(0,0):SetString(Nombre_Programa())
   oSheet:getCellByPosition(0,1):SetString(DIA(DATE(),10))
   oSheet:getCellByPosition(0,2):SetString(NOMEMPRESA)
   oSheet:getCellByPosition(0,3):SetString(TituloImp)
   oSheet:getCellByPosition(0,3):CharHeight:=16
   oSheet:getCellByPosition(0,4):SetString('Desde codigo:')
   oSheet:getCellByPosition(1,4):SetValue(W_Imp1.T_CodCta1.value)
   oSheet:getCellByPosition(1,4):HoriJustify:=1
   oSheet:getCellByPosition(0,5):SetString('Hasta codigo:')
   oSheet:getCellByPosition(1,5):SetValue(W_Imp1.T_CodCta2.value)
   oSheet:getCellByPosition(1,5):HoriJustify:=1
   oSheet:getCellByPosition(0,6):SetString('Desde fecha:')
   oSheet:getCellByPosition(1,6):SetString(DIA(W_Imp1.D_Fec1.value,10))
   oSheet:getCellByPosition(0,7):SetString('Hasta fecha:')
   oSheet:getCellByPosition(1,7):SetString(DIA(W_Imp1.D_Fec2.value,10))

   LIN:=9

   oSheet:getCellByPosition(0,LIN):SetString("Codigo")
   oSheet:getCellByPosition(1,LIN):SetString("Descripcion")
   oSheet:getCellByPosition(2,LIN):SetString("Saldo inicial")
   oSheet:getCellByPosition(3,LIN):SetString("Debe")
   oSheet:getCellByPosition(4,LIN):SetString("Haber")
   oSheet:getCellByPosition(5,LIN):SetString("Saldo")
   oSheet:getCellRangeByPosition(2,LIN,5,LIN):HoriJustify:=3
   oSheet:getCellRangeByPosition(0,LIN,5,LIN):CharWeight:=150 //NEGRITA
   aMiColor:=MiColor("AMARILLOPALIDO")
   oSheet:getCellRangeByPosition(0,LIN,5,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

   LIN++
   LIN1:=LIN+1

W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
PAG:=W_Imp1.T_Hoja.value
*LIN:=0
   SALT:=0
   DEBT:=0
   HABT:=0
SELEC("FINSUIC")
GO TOP
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS

   oSheet:getCellByPosition(0,LIN):SetString(LTRIM(STR(CODCTA)))
   oSheet:getCellByPosition(1,LIN):SetString(RTRIM(NOMCTA))
   oSheet:getCellByPosition(2,LIN):SetValue(SALDO)
   oSheet:getCellByPosition(3,LIN):SetValue(DEBE)
   oSheet:getCellByPosition(4,LIN):SetValue(HABER)
   oSheet:getCellByPosition(5,LIN):SetFormula("=+C"+LTRIM(STR(LIN+1))+"+D"+LTRIM(STR(LIN+1))+"-E"+LTRIM(STR(LIN+1)))
   oSheet:getCellRangeByPosition(2,LIN,5,LIN):NumberFormat:=4 //#.##0,00

   LIN++
   SKIP

ENDDO

   oSheet:getCellByPosition(1,LIN):SetString("Total")
   oSheet:getCellByPosition(2,LIN):SetFormula("=SUM("+LetraExcel(3)+LTRIM(STR(LIN1))+":"+LetraExcel(3)+LTRIM(STR(LIN))+")")
   oSheet:getCellByPosition(3,LIN):SetFormula("=SUM("+LetraExcel(4)+LTRIM(STR(LIN1))+":"+LetraExcel(4)+LTRIM(STR(LIN))+")")
   oSheet:getCellByPosition(4,LIN):SetFormula("=SUM("+LetraExcel(5)+LTRIM(STR(LIN1))+":"+LetraExcel(5)+LTRIM(STR(LIN))+")")
   oSheet:getCellByPosition(5,LIN):SetFormula("=SUM("+LetraExcel(6)+LTRIM(STR(LIN1))+":"+LetraExcel(6)+LTRIM(STR(LIN))+")")
   oSheet:getCellRangeByPosition(2,LIN,5,LIN):NumberFormat:=4 //#.##0,00
   oSheet:getCellRangeByPosition(1,LIN,5,LIN):CharWeight:=150 //NEGRITA
   aMiColor:=MiColor("VERDEPALIDO")
   oSheet:getCellRangeByPosition(1,LIN,5,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

   oColumns:=oSheet:getColumns()
   oColumns:getByIndex(1):setPropertyValue("OptimalWidth", .T.)
*   oColumn:Width:=5*100

   QuitarEspera()

W_Imp1.release



STATIC FUNCTION balsys2i_Excel()
   LOCAL oExcel, oHoja

   PonerEspera("Procesando datos del listado")

   oExcel := TOleAuto():New( "Excel.Application" )
   IF Ole2TxtError() != 'S_OK'
      QuitarEspera()
      MsgStop('Excel no esta disponible','error')
      RETURN Nil
   ENDIF
   oExcel:WorkBooks:Add()
   oExcel:Sheets("Hoja1"):Name := "Listado"
   oHoja := oExcel:Get( "ActiveSheet" )
   oHoja:Cells:Font:Name := "Arial"
   oHoja:Cells:Font:Size := 10

   LIN:=10

oHoja:Cells( LIN, 1 ):Value := "Codigo"
oHoja:Cells( LIN, 1 ):HorizontalAlignment:= -4152  //Derecha
oHoja:Cells( LIN, 2 ):Value := "Descripcion"
oHoja:Cells( LIN, 3 ):Value := "Saldo inicial"
oHoja:Cells( LIN, 4 ):Value := "Debe"
oHoja:Cells( LIN, 5 ):Value := "Haber"
oHoja:Cells( LIN, 6 ):Value := "Saldo"
oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Font:Bold := .T.
oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Interior:ColorIndex := 36 //sombrear celdas  amarillo claro
oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Borders(4):LineStyle:= 1  //linea inferior
oHoja:Range(LetraExcel(3)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):HorizontalAlignment := -4152  //Derecha

   LIN++
   LIN1:=LIN

W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
PAG:=W_Imp1.T_Hoja.value
*LIN:=0
   SALT:=0
   DEBT:=0
   HABT:=0
SELEC("FINSUIC")
GO TOP
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS

   oHoja:Cells( LIN, 1 ):Value := LTRIM(STR(CODCTA))
   oHoja:Cells( LIN, 2 ):Value := RTRIM(NOMCTA)
   oHoja:Cells( LIN, 3 ):Value := SALDO
   oHoja:Cells( LIN, 4 ):Value := DEBE
   oHoja:Cells( LIN, 5 ):Value := HABER
   oHoja:Cells( LIN, 6 ):Value := "=+C"+LTRIM(STR(LIN))+"+D"+LTRIM(STR(LIN))+"-E"+LTRIM(STR(LIN))
   oHoja:Range(LetraExcel(3)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Set( "NumberFormat", "#.##0,00" )

   LIN++
   SKIP

ENDDO

   oHoja:Cells( LIN, 2 ):Value := "Total"
   oHoja:Cells( LIN, 2 ):HorizontalAlignment:= -4152  //Derecha
   oHoja:Cells( LIN, 3 ):Value := "=SUMA("+LetraExcel(3)+LTRIM(STR(LIN1))+":"+LetraExcel(3)+LTRIM(STR(LIN-1))+")"
   oHoja:Cells( LIN, 4 ):Value := "=SUMA("+LetraExcel(4)+LTRIM(STR(LIN1))+":"+LetraExcel(4)+LTRIM(STR(LIN-1))+")"
   oHoja:Cells( LIN, 5 ):Value := "=SUMA("+LetraExcel(5)+LTRIM(STR(LIN1))+":"+LetraExcel(5)+LTRIM(STR(LIN-1))+")"
   oHoja:Cells( LIN, 6 ):Value := "=SUMA("+LetraExcel(6)+LTRIM(STR(LIN1))+":"+LetraExcel(6)+LTRIM(STR(LIN-1))+")"
   oHoja:Range(LetraExcel(3)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Set( "NumberFormat", "#.##0,00" )
   oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Interior:ColorIndex := 35 //sombrear celdas verde claro
   oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Borders(8):LineStyle:= 1  //linea inferior


   oHoja:Cells( 1, 1 ):Value := Nombre_Programa()
   oHoja:Cells( 2, 1 ):Value := DIA(DATE(),10)
   oHoja:Cells( 5, 1 ):Value := 'Desde codigo:'
   oHoja:Cells( 5, 2 ):Value := W_Imp1.T_CodCta1.value
   oHoja:Cells( 6, 1 ):Value := 'Hasta codigo:'
   oHoja:Cells( 6, 2 ):Value := W_Imp1.T_CodCta2.value
   oHoja:Cells( 7, 1 ):Value := 'Desde fecha:'
   oHoja:Cells( 7, 2 ):Value := DIA(W_Imp1.D_Fec1.value,10)
   oHoja:Cells( 8, 1 ):Value := 'Hasta fecha:'
   oHoja:Cells( 8, 2 ):Value := DIA(W_Imp1.D_Fec2.value,10)
   oHoja:Range("A1:B8"):HorizontalAlignment:= -4131  //Izquierda

   FOR nCol:=1 TO FCOUNT()
      oHoja:Columns( nCol ):AutoFit()
   NEXT

   oHoja:Cells( 3, 1 ):Value := NOMEMPRESA
   oHoja:Cells( 4, 1 ):Value := TituloImp
   oHoja:Cells( 4, 1 ):Font:Size := 16

   *Guardar como
   *oHoja:SaveAs( TituloImp )

   oHoja:Cells( 1, 1 ):Select()
   oExcel:Visible := .T.

   IF AT("XHARBOUR",UPPER(VERSION()))=1
      oHoja:end()
      oExcel:end()
   ENDIF

*   SELEC FIN
*   FIN->( DBCLOSEAREA() )

   QuitarEspera()

   W_Imp1.release

Return Nil




