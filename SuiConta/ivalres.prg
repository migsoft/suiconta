#include "minigui.ch"
#include "winprint.ch"

procedure Ivalres()
    LOCAL Color_Iva1 := { |x,nItem| IF( x[7]="Emi" , RGB(200,255,200), RGB(200,200,255) ) }

    TituloImp:="Resumen de IVA"

    DEFINE WINDOW W_IvaR ;
        AT 10,10 ;
        WIDTH 800 HEIGHT 600 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 10,10 RADIOGROUP R_Trimestre OPTIONS {'1º Trim.','2º Trim.','3º Trim.','4º Trim.'} ;
            VALUE 1 WIDTH 75 HORIZONTAL ON CHANGE Ivalres_Trimestre(W_IvaR.R_Trimestre.Value)

        @ 45,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 75,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 70,110 DATEPICKER D_Fec2 WIDTH 100 VALUE IF(DMA2<DATE(),DMA2,DATE())

        @ 45,390 LABEL L_FacEmi VALUE 'Facturas emitidas' AUTOSIZE TRANSPARENT
        @ 40,500 PROGRESSBAR P_FacEmi RANGE 0 , 100 WIDTH 270 HEIGHT 20 SMOOTH
        @ 75,390 LABEL L_FacRec VALUE 'Facturas recibidas' AUTOSIZE TRANSPARENT
        @ 70,500 PROGRESSBAR P_FacRec RANGE 0 , 100 WIDTH 270 HEIGHT 20 SMOOTH

        @110,010 GRID GR_Lis ;
            WIDTH 530 HEIGHT 380 ;
            HEADERS {'Resumen'} ;
            WIDTHS {510} ;
            ITEMS {} ;
            FONT "Courier" SIZE 10 ;
            COLUMNCONTROLS {{'TEXTBOX','CHARACTER'}} ;
            JUSTIFY {BROWSE_JTFY_LEFT}

        @110,550 GRID GR_Iva ;
            WIDTH 220 HEIGHT 380 ;
            HEADERS {'TIPIVA2','REGIMEN','IVAT','BIMPT','CUOTAT','IVA','FACTURA'} ;
            WIDTHS {0,0,50,100,100,50,100 } ;
            ITEMS {} ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'}, ;
            {'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT, ;
            BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT} ;
            DYNAMICBACKCOLOR { Color_Iva1,Color_Iva1,Color_Iva1,Color_Iva1,Color_Iva1,Color_Iva1,Color_Iva1}

        Menu_Grid("W_IvaR","GR_Iva","MENU",,{"COPCELPP","COPREGPP","COPTABPP"})


        @530, 10 BUTTON B_Calcular CAPTION 'Calcular' WIDTH 95 HEIGHT 25 ;
            ACTION Ivalres_Calcular()

        @530,110 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 95 HEIGHT 25 ;
            ACTION Ivalres_imp()

        @530,210 BUTTONEX B_Writer CAPTION 'Hoja Texto' ICON icobus('writer') WIDTH 95 HEIGHT 25 ;
            ACTION Ivalres_Writer()

        @530,310 BUTTONEX B_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 95 HEIGHT 25 ;
            ACTION Ivalres_Calc()

        @530,410 BUTTON B_Conta CAPTION 'Contabilizar' WIDTH 95 HEIGHT 25 ;
            ACTION Ivalres_Conta()

        @530,510 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 95 HEIGHT 25 ;
            ACTION W_IvaR.release

        Ivalres_Trimestre(0)


    END WINDOW
    VentanaCentrar("W_IvaR","Ventana1")
    CENTER WINDOW W_IvaR
    ACTIVATE WINDOW W_IvaR

Return Nil

STATIC FUNCTION Ivalres_Trimestre(NumTrim)
    IF NumTrim=0
        DO CASE
        CASE MONTH(DATE())<=4 .AND. YEAR(DATE())=EJERANY
            W_IvaR.R_Trimestre.Value:=1
        CASE MONTH(DATE())<=7 .AND. YEAR(DATE())=EJERANY
            W_IvaR.R_Trimestre.Value:=2
        CASE MONTH(DATE())<=10 .AND. YEAR(DATE())=EJERANY
            W_IvaR.R_Trimestre.Value:=3
        OTHERWISE
            W_IvaR.R_Trimestre.Value:=4
        ENDCASE
        NumTrim:=W_IvaR.R_Trimestre.Value
    ENDIF

    DO CASE
    CASE NumTrim=1
        W_IvaR.D_Fec1.value:=CTOD("01-01-"+STR(EJERANY,4))
        W_IvaR.D_Fec2.value:=DIAMESMAS(W_IvaR.D_Fec1.value,3)-1
    CASE NumTrim=2
        W_IvaR.D_Fec1.value:=CTOD("01-04-"+STR(EJERANY,4))
        W_IvaR.D_Fec2.value:=DIAMESMAS(W_IvaR.D_Fec1.value,3)-1
    CASE NumTrim=3
        W_IvaR.D_Fec1.value:=CTOD("01-07-"+STR(EJERANY,4))
        W_IvaR.D_Fec2.value:=DIAMESMAS(W_IvaR.D_Fec1.value,3)-1
    CASE NumTrim=4
        W_IvaR.D_Fec1.value:=CTOD("01-10-"+STR(EJERANY,4))
        W_IvaR.D_Fec2.value:=DIAMESMAS(W_IvaR.D_Fec1.value,3)-1
    ENDCASE
Return Nil

STATIC FUNCTION Ivalres_Calcular()
    MIVAE:={}
    AbrirDBF("FAC92")
    W_IvaR.P_FacEmi.RangeMax:=LASTREC()
    W_IvaR.P_FacEmi.Value:=0
    GO TOP
    DO WHILE .NOT. EOF()
        W_IvaR.P_FacEmi.Value:=W_IvaR.P_FacEmi.Value+1
        DO EVENTS
        IF FFAC>=W_IvaR.D_Fec1.value .AND. FFAC<=W_IvaR.D_Fec2.value
            TIPIVA2:=(REGIMEN*1000)+IVA
            NUM2:=ASCAN(MIVAE,{|AVAL| AVAL[1]=TIPIVA2})
            IF NUM2=0
                AADD(MIVAE,{TIPIVA2,REGIMEN,IVA,BIMP,IMPIVA,"IVA","Emitidas"})
            ELSE
                MIVAE[NUM2,4]=MIVAE[NUM2,4]+BIMP
                MIVAE[NUM2,5]=MIVAE[NUM2,5]+IMPIVA
            ENDIF
            IF CAMPO("BIMPT2")=.T.
                IF IMPIVAT2<>0
                    TIPIVA2:=(REGIMEN*1000)+IVAT2
                    NUM2:=ASCAN(MIVAE,{|AVAL| AVAL[1]=TIPIVA2})
                    IF NUM2=0
                        AADD(MIVAE,{TIPIVA2,REGIMEN,IVAT2,BIMPT2,IMPIVAT2,"IVA","Emitidas"})
                    ELSE
                        MIVAE[NUM2,4]=MIVAE[NUM2,4]+BIMPT2
                        MIVAE[NUM2,5]=MIVAE[NUM2,5]+IMPIVAT2
                    ENDIF
                ENDIF
                IF IMPIVAT3<>0
                    TIPIVA2:=(REGIMEN*1000)+IVAT3
                    NUM2:=ASCAN(MIVAE,{|AVAL| AVAL[1]=TIPIVA2})
                    IF NUM2=0
                        AADD(MIVAE,{TIPIVA2,REGIMEN,IVAT3,BIMPT3,IMPIVAT3,"IVA","Emitidas"})
                    ELSE
                        MIVAE[NUM2,4]=MIVAE[NUM2,4]+BIMPT3
                        MIVAE[NUM2,5]=MIVAE[NUM2,5]+IMPIVAT3
                    ENDIF
                ENDIF
            ENDIF
            IF REQ<>0
                TIPIVA2:=(REGIMEN*1000)+REQ+800
                NUM2:=ASCAN(MIVAE,{|AVAL| AVAL[1]=TIPIVA2})
                IF NUM2=0
                    AADD(MIVAE,{TIPIVA2,REGIMEN,REQ,BIMP,IMPREQ,"REQ","Emitidas"})
                ELSE
                    MIVAE[NUM2,4]=MIVAE[NUM2,4]+BIMP
                    MIVAE[NUM2,5]=MIVAE[NUM2,5]+IMPREQ
                ENDIF
            ENDIF
        ENDIF
        SKIP
    ENDDO
    MIVAR:={}
    AbrirDBF("FACREB")
    W_IvaR.P_FacRec.RangeMax:=LASTREC()
    W_IvaR.P_FacRec.Value:=0
    GO TOP
    DO WHILE .NOT. EOF()
        W_IvaR.P_FacRec.Value:=W_IvaR.P_FacRec.Value+1
        DO EVENTS
        IF FREG>=W_IvaR.D_Fec1.value .AND. FREG<=W_IvaR.D_Fec2.value
            IF REQ<>0
                TIPIVA2:=(REGIMEN*1000)+REQ+800
                NUM2:=ASCAN(MIVAR,{|AVAL| AVAL[1]=TIPIVA2})
                IF NUM2=0
                    AADD(MIVAR,{TIPIVA2,REGIMEN,REQ,BIMP+BIMPT2+BIMPT3,IMPREQ,"REQ","Recibidas"})
                ELSE
                    MIVAR[NUM2,4]=MIVAR[NUM2,4]+BIMP+BIMPT2+BIMPT3
                    MIVAR[NUM2,5]=MIVAR[NUM2,5]+IMPREQ
                ENDIF
            ENDIF
            TIPIVA2:=(REGIMEN*1000)+IVA
            NUM2:=ASCAN(MIVAR,{|AVAL| AVAL[1]=TIPIVA2})
            IF NUM2=0
                AADD(MIVAR,{TIPIVA2,REGIMEN,IVA,BIMP,CUOTA,"IVA","Recibidas"})
            ELSE
                MIVAR[NUM2,4]=MIVAR[NUM2,4]+BIMP
                MIVAR[NUM2,5]=MIVAR[NUM2,5]+CUOTA
            ENDIF
            IF BIMPT2<>0
                TIPIVA2:=(REGIMEN*1000)+IVAT2
                NUM2:=ASCAN(MIVAR,{|AVAL| AVAL[1]=TIPIVA2})
                IF NUM2=0
                    AADD(MIVAR,{TIPIVA2,REGIMEN,IVAT2,BIMPT2,CUOTAT2,"IVA","Recibidas"})
                ELSE
                    MIVAR[NUM2,4]=MIVAR[NUM2,4]+BIMPT2
                    MIVAR[NUM2,5]=MIVAR[NUM2,5]+CUOTAT2
                ENDIF
            ENDIF
            IF BIMPT3<>0
                TIPIVA2:=(REGIMEN*1000)+IVAT3
                NUM2:=ASCAN(MIVAR,{|AVAL| AVAL[1]=TIPIVA2})
                IF NUM2=0
                    AADD(MIVAR,{TIPIVA2,REGIMEN,IVAT3,BIMPT3,CUOTAT3,"IVA","Recibidas"})
                ELSE
                    MIVAR[NUM2,4]=MIVAR[NUM2,4]+BIMPT3
                    MIVAR[NUM2,5]=MIVAR[NUM2,5]+CUOTAT3
                ENDIF
            ENDIF

            IF REGIMEN=2                            //IVA INTRACOMUNITARIO FACTURAS RECIBIDAS -> FACTURAS EMITIDAS
                TIPIVA2:=(REGIMEN*1000)+IVA
                NUM2:=ASCAN(MIVAE,{|AVAL| AVAL[1]=TIPIVA2})
                IF NUM2=0
                    AADD(MIVAE,{TIPIVA2,REGIMEN,IVA,BIMP,CUOTA,"IVA","Emitidas"})
                ELSE
                    MIVAE[NUM2,4]=MIVAE[NUM2,4]+BIMP
                    MIVAE[NUM2,5]=MIVAE[NUM2,5]+CUOTA
                ENDIF
            ENDIF
        ENDIF
        SKIP
    ENDDO
    ASORT(MIVAE,,,{|X,Y| X[1]<Y[1]})
    ASORT(MIVAR,,,{|X,Y| X[1]<Y[1]})

***GRID***
    W_IvaR.GR_Iva.DeleteAllItems
    FOR N=1 TO LEN(MIVAE)
        W_IvaR.GR_Iva.AddItem(MIVAE[N])
    NEXT
    FOR N=1 TO LEN(MIVAR)
        W_IvaR.GR_Iva.AddItem(MIVAR[N])
    NEXT
    W_IvaR.GR_Iva.Refresh
***GRID***


    IF W_IvaR.GR_Iva.ItemCount=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF




    W_IvaR.GR_Lis.DeleteAllItems

    TEXT02:=SPACE(3)+"Resumen de I.V.A.  del "+DTOC(W_IvaR.D_Fec1.value)+" al "+DTOC(W_IvaR.D_Fec2.value)
    W_IvaR.GR_Lis.AddItem({TEXT02})
    W_IvaR.GR_Lis.AddItem({" "})
    TEXT02:="Tipo Fac. Regimen  Tipo Iva     Base Imp.         Cuota"
    W_IvaR.GR_Lis.AddItem({TEXT02})
    TEXT02:=PADR("-",55,"-")
    W_IvaR.GR_Lis.AddItem({TEXT02})


    TOT1:=0
    TOT2:=0
    LIN:=0
    TIPFAC3:=""
    REGIMEN3:=-1

    FOR N=1 TO W_IvaR.GR_Iva.ItemCount

        TIPREG2:=W_IvaR.GR_Iva.Cell(N,1)
        REGIMEN2:=W_IvaR.GR_Iva.Cell(N,2)
        TIPIVA2:=W_IvaR.GR_Iva.Cell(N,3)
        BIMP2:=W_IvaR.GR_Iva.Cell(N,4)
        IMPIVA2:=W_IvaR.GR_Iva.Cell(N,5)
        SIREQ2:=W_IvaR.GR_Iva.Cell(N,6)
        TIPFAC2:=W_IvaR.GR_Iva.Cell(N,7)

        IF TIPFAC3<>TIPFAC2
            TIPFAC3:=TIPFAC2
            IF AT("EMITIDAS",UPPER(TIPFAC3))<>0
                TEXT02:="Facturas Emitidas"
                W_IvaR.GR_Lis.AddItem({TEXT02})
            ELSE
                TEXT02:=SPACE(10)+"Total "+MDA_NOM2(EJERANY)+" Facturas emitidas  "+MIL(TOT1,12,MDA_DEC(EJERANY))
                W_IvaR.GR_Lis.AddItem({TEXT02})
                W_IvaR.GR_Lis.AddItem({""})
                TEXT02:="Facturas Recibidas"
                W_IvaR.GR_Lis.AddItem({TEXT02})
                REGIMEN3:=-1
            ENDIF
        ENDIF
        IF REGIMEN3<>REGIMEN2
            REGIMEN3:=REGIMEN2
            TEXT02:=SPACE(8)
            IF REGIMEN2=1
                IF TIPREG2>2000                     //FACTURAS INTRACOMUINITARIAS
                    TEXT02:=TEXT02+MIL(REGIMEN2,3,0)+"-"+REGIVAREC(REGIMEN2)
                ELSE
                    TEXT02:=TEXT02+MIL(REGIMEN2,3,0)+"-"+REGIVAEMI(REGIMEN2)
                ENDIF
            ELSE
                TEXT02:=TEXT02+MIL(REGIMEN2,3,0)+"-"+REGIVAREC(REGIMEN2)
            ENDIF
            W_IvaR.GR_Lis.AddItem({TEXT02})
        ENDIF
        TEXT02:=SPACE(13)
        IF SIREQ2="REQ"
            TEXT02:=TEXT02+"Recargo "
        ELSE
            TEXT02:=TEXT02+SPACE(8)
        ENDIF
        TEXT02:=TEXT02+MIL(TIPIVA2,5,2)+"%"+SPACE(2)+MIL(BIMP2,12,MDA_DEC(EJERANY))
        TEXT02:=TEXT02+SPACE(2)+MIL(IMPIVA2,12,MDA_DEC(EJERANY))
        W_IvaR.GR_Lis.AddItem({TEXT02})

        IF AT("EMITIDAS",UPPER(TIPFAC3))<>0
            TOT1=TOT1+IMPIVA2
        ELSE
            TOT2=TOT2+IMPIVA2
        ENDIF

    NEXT
    TEXT02:=SPACE(10)+"Total "+MDA_NOM2(EJERANY)+" Facturas Recibidas "+MIL(TOT2,12,MDA_DEC(EJERANY))
    W_IvaR.GR_Lis.AddItem({TEXT02})
    W_IvaR.GR_Lis.AddItem({""})


    TEXT02:=SPACE(10)+"Total "+MDA_NOM2(EJERANY)+" Facturas Emitidas  "+MIL(TOT1,12,MDA_DEC(EJERANY))
    W_IvaR.GR_Lis.AddItem({TEXT02})
    TEXT02:=SPACE(10)+"Total "+MDA_NOM2(EJERANY)+" Facturas Recibidas "+MIL(TOT2,12,MDA_DEC(EJERANY))
    W_IvaR.GR_Lis.AddItem({TEXT02})
    TEXT02:=SPACE(10)+PADR("-",45,"-")
    W_IvaR.GR_Lis.AddItem({TEXT02})
    TEXT02:=SPACE(10)+"Diferencia"+SPACE(23)+MIL(TOT1-TOT2,12,MDA_DEC(EJERANY))
    W_IvaR.GR_Lis.AddItem({TEXT02})


Return Nil



STATIC FUNCTION Ivalres_Writer()
    nOption:=InputWindow("Elija el tipo de salida",{"Programa"},{1},{{"Hoja Texto","Hoja Word"}},,,.T.,{"Aceptar","Cancelar"})
    DO CASE
    CASE nOption[1] == Nil
        RETURN
    CASE nOption[1]=1
        Ivalres_Writer2()
    CASE nOption[1]=2
        Ivalres_Word()
    ENDCASE
Return Nil

STATIC FUNCTION Ivalres_Writer2()
    local oServiceManager,oDesktop,oDocument,oCursor

    PonerEspera("Procesando datos del listado")

    // inicializa
    IF ( oServiceManager := win_oleCreateObject("com.sun.star.ServiceManager") ) != NIL
        IF ( oDesktop := oServiceManager:createInstance("com.sun.star.frame.Desktop") ) != NIL
            QuitarEspera()
            MsgStop('OpenOffice Writer no esta disponible','error')
            RETURN Nil
        ENDIF
        oDocument := oDesktop:loadComponentFromURL("private:factory/swriter","_blank", 0, {})

        oCursor := oDocument:Text:CreateTextCursor()
        oCursor:CharFontName:="Courier"
        oCursor:CharHeight:=10
        oCursor:CharWeight:=150

        oDocument:Text:InsertString(oCursor, NOMPROGRAMA+SPACE(53)+"Hoja: 1"+CHR(13) , .F.)
        IF AT("\SUIZOWIN.ET",UPPER(RUTAPROGRAMA))<>0
            oDocument:Text:InsertString(oCursor, DIA(DATE(),10)+SPACE(11)+NOMEMPRESA+SPACE(20)+"Juan Bravo López"+CHR(13) , .F.)
        ELSE
            oDocument:Text:InsertString(oCursor, DIA(DATE(),10)+SPACE(11)+NOMEMPRESA+CHR(13) , .F.)
        ENDIF
        oCursor:CharHeight:=16
        oDocument:Text:InsertString(oCursor, SPACE(13)+TituloImp+CHR(13)+CHR(13) , .F.)
        oCursor:CharHeight:=10

        FOR N=1 TO W_IvaR.GR_Lis.ItemCount
            oDocument:Text:InsertString(oCursor, W_IvaR.GR_Lis.Cell(N,1)+CHR(13) , .F.)
        NEXT

    ELSE
        MsgInfo( "Error: OpenOffice not available. ", win_oleErrorText() )
    ENDIF
    QuitarEspera()

Return Nil

STATIC FUNCTION Ivalres_Word()
    LOCAL oWord, oTexto

    PonerEspera("Procesando datos del listado")

    oWord:=TOleAuto():New( "Word.Application" )
    IF Ole2TxtError() != 'S_OK'
        QuitarEspera()
        MsgStop('Word no esta disponible','error')
        RETURN Nil
    ENDIF
    oWord:Documents:Add()
    oTexto := oWord:Selection()

    Texto2:=NOMPROGRAMA+SPACE(43)+"Hoja: 1"+CRLF
    IF AT("\SUIZOWIN.ET",UPPER(RUTAPROGRAMA))<>0
        Texto2:=Texto2+DIA(DATE(),10)+SPACE(11)+NOMEMPRESA+SPACE(10)+"Juan Bravo López"+CRLF
    ELSE
        Texto2:=Texto2+DIA(DATE(),10)+SPACE(11)+NOMEMPRESA+CRLF
    ENDIF
    Texto2:=Texto2+SPACE(21)+TituloImp+CRLF+CRLF

    FOR N=1 TO W_IvaR.GR_Lis.ItemCount
        Texto2:=Texto2+W_IvaR.GR_Lis.Cell(N,1)+CRLF
    NEXT

    oTexto:Text := Texto2
    oTexto:Font:Name := "Courier"
    oTexto:Font:Size := 10
    oTexto:Font:Bold := .T.

    oWord:Visible := .T.
    oWord:Set( "WindowState", 1 )                   // Maximizado

    IF AT("XHARBOUR",UPPER(VERSION()))=1
        oTexto:End()
        oWord:End()
    ENDIF

    QuitarEspera()

Return Nil

STATIC FUNCTION Ivalres_Calc()
    local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

    PonerEspera("Procesando datos del listado")

    // inicializa
    IF ( oServiceManager := win_oleCreateObject("com.sun.star.ServiceManager") ) != NIL
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

        LIN:=5

        oSheet:getCellByPosition(0,LIN):SetString("Tipo Fac.")
        oSheet:getCellByPosition(1,LIN):SetString("Regimen")
        oSheet:getCellByPosition(2,LIN):SetString("Tipo Iva")
        oSheet:getCellByPosition(3,LIN):SetString("Base Imp.")
        oSheet:getCellByPosition(4,LIN):SetString("Cuota")
        oSheet:getCellRangeByPosition(2,LIN,4,LIN):HoriJustify:=3
        oSheet:getCellRangeByPosition(0,LIN,4,LIN):CharWeight:=150  //NEGRITA
        aMiColor:=MiColor("AMARILLOPALIDO")
        oSheet:getCellRangeByPosition(0,LIN,4,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

        LIN++
        LIN1:=LIN+1


        FOR N=1 TO W_IvaR.GR_Iva.ItemCount
            DO EVENTS
            IF AT("EMITIDAS",UPPER(W_IvaR.GR_Iva.Cell(N,7)))<>0
                oSheet:getCellByPosition(0,LIN):SetString(W_IvaR.GR_Iva.Cell(N,7))
                oSheet:getCellByPosition(1,LIN):SetString(REGIVAEMI(W_IvaR.GR_Iva.Cell(N,2)))
                oSheet:getCellByPosition(2,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,1))
                oSheet:getCellByPosition(3,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,4))
                oSheet:getCellByPosition(4,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,5))
                oSheet:getCellRangeByPosition(2,LIN,4,LIN):NumberFormat:=4  //#.##0,00
                LIN++
            ENDIF
        NEXT

        oSheet:getCellByPosition(3,LIN):SetString("Total")
        oSheet:getCellByPosition(4,LIN):SetFormula("=SUM("+LetraExcel(5)+LTRIM(STR(LIN1))+":"+LetraExcel(5)+LTRIM(STR(LIN))+")")
        oSheet:getCellRangeByPosition(4,LIN,4,LIN):NumberFormat:=4  //#.##0,00
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CharWeight:=150  //NEGRITA
        aMiColor:=MiColor("VERDEPALIDO")
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
        LIN2A:=LIN+1

        LIN:=LIN+2
        LIN1:=LIN+1

        FOR N=1 TO W_IvaR.GR_Iva.ItemCount
            DO EVENTS
            IF AT("EMITIDAS",UPPER(W_IvaR.GR_Iva.Cell(N,7)))=0
                oSheet:getCellByPosition(0,LIN):SetString(W_IvaR.GR_Iva.Cell(N,7))
                oSheet:getCellByPosition(1,LIN):SetString(REGIVAREC(W_IvaR.GR_Iva.Cell(N,2)))
                oSheet:getCellByPosition(2,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,1))
                oSheet:getCellByPosition(3,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,4))
                oSheet:getCellByPosition(4,LIN):SetValue(W_IvaR.GR_Iva.Cell(N,5))
                oSheet:getCellRangeByPosition(2,LIN,4,LIN):NumberFormat:=4  //#.##0,00
                LIN++
            ENDIF
        NEXT
        oSheet:getCellByPosition(3,LIN):SetString("Total")
        oSheet:getCellByPosition(4,LIN):SetFormula("=SUM("+LetraExcel(5)+LTRIM(STR(LIN1))+":"+LetraExcel(5)+LTRIM(STR(LIN))+")")
        oSheet:getCellRangeByPosition(4,LIN,4,LIN):NumberFormat:=4  //#.##0,00
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CharWeight:=150  //NEGRITA
        aMiColor:=MiColor("VERDEPALIDO")
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
        LIN2B:=LIN+1

        LIN:=LIN+2

        oSheet:getCellByPosition(3,LIN):SetString("Diferencia")
        oSheet:getCellByPosition(4,LIN):SetFormula("="+LetraExcel(5)+LTRIM(STR(LIN2A))+"-"+LetraExcel(5)+LTRIM(STR(LIN2B)) )
        oSheet:getCellRangeByPosition(4,LIN,4,LIN):NumberFormat:=4  //#.##0,00
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CharWeight:=150  //NEGRITA
        aMiColor:=MiColor("VERDEPALIDO")
        oSheet:getCellRangeByPosition(3,LIN,4,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])


        oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)

        oSheet:getCellByPosition(0,0):SetString(DIA(DATE(),10))
        oSheet:getCellByPosition(0,1):SetString(Nombre_Programa())
        oSheet:getCellByPosition(0,2):SetString(NOMEMPRESA)

        oSheet:getCellByPosition(0,3):SetString("Resumen de I.V.A.  del "+DTOC(W_IvaR.D_Fec1.value)+" al "+DTOC(W_IvaR.D_Fec2.value))
        oSheet:getCellByPosition(0,3):CharWeight:=150  //NEGRITA

    ELSE
        MsgInfo( "Error: OpenOffice not available. ", win_oleErrorText() )
    ENDIF

    QuitarEspera()

Return Nil



STATIC FUNCTION Ivalres_imp()

    IF W_IvaR.GR_Iva.ItemCount=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 160 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        aIMP:=Impresoras(EMP_IMPRESORA)
        @ 15,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @ 10,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @ 45,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @ 40,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @ 40, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @ 70, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @100, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION Ivalres_impi()

        @100,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil

STATIC FUNCTION Ivalres_impi()
    local oprint

    IF W_IvaR.GR_Iva.ItemCount=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
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
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    PAG:=0
    LIN:=0
    FOR N=1 TO W_IvaR.GR_Lis.ItemCount
        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)

            LIN:=55
        ENDIF

        oprint:printdata(LIN,30,W_IvaR.GR_Lis.Cell(N,1),"Courier",10,.T.,,"L",)

        LIN:=LIN+5

    NEXT

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

Return Nil

STATIC FUNCTION Ivalres_Conta()
    IF MSGYESNO("¿Desea crear el asiento de ajuste de IVA?")=.F.
        RETURN
    ENDIF

    F1:=W_IvaR.D_Fec1.value
    F2:=W_IvaR.D_Fec2.value
    DO CASE
    CASE DAY(F1)=1 .AND. MONTH(F1)=1 .AND. F2=DIAMESMAS(F1,3)-1
        CONCEPTO2:="Traspaso IVA 1T/"+STR(YEAR(F2),4)
    CASE DAY(F1)=1 .AND. MONTH(F1)=4 .AND. F2=DIAMESMAS(F1,3)-1
        CONCEPTO2:="Traspaso IVA 2T/"+STR(YEAR(F2),4)
    CASE DAY(F1)=1 .AND. MONTH(F1)=7 .AND. F2=DIAMESMAS(F1,3)-1
        CONCEPTO2:="Traspaso IVA 3T/"+STR(YEAR(F2),4)
    CASE DAY(F1)=1 .AND. MONTH(F1)=10 .AND. F2=DIAMESMAS(F1,3)-1
        CONCEPTO2:="Traspaso IVA 4T/"+STR(YEAR(F2),4)
    OTHERWISE
        CONCEPTO2:="Traspaso IVA "+DIA(F1,6)+" a "+DIA(F2,6)
    ENDCASE

    AbrirDBF("CUENTAS")
    AbrirDBF("APUNTES",,,,,1)
    GO BOTT
    NASI2:=NASI+1
    FASI2:=IF(W_IvaR.D_Fec2.value>DMA2,DMA2,W_IvaR.D_Fec2.value)
    TOT2:=0
    APU2:=1
    FOR N=1 TO W_IvaR.GR_Iva.ItemCount
        IF W_IvaR.GR_Iva.Cell(N,5)<>0
            IF W_IvaR.GR_Iva.Cell(N,6)="REQ"
                IF AT("EMITIDAS",UPPER(W_IvaR.GR_Iva.Cell(N,7)))<>0
                    CODCTA2:=47700100+ROUND(W_IvaR.GR_Iva.Cell(N,3),0)
                ELSE
                    CODCTA2:=47200100+ROUND(W_IvaR.GR_Iva.Cell(N,3),0)
                ENDIF
            ELSE
                IF AT("EMITIDAS",UPPER(W_IvaR.GR_Iva.Cell(N,7)))<>0
                    CODCTA2:=47700000+W_IvaR.GR_Iva.Cell(N,1)
                ELSE
                    CODCTA2:=47200000+W_IvaR.GR_Iva.Cell(N,1)
                ENDIF
            ENDIF
            APPEND BLANK
            IF RLOCK()
                REPLACE NASI WITH NASI2
                REPLACE APU WITH APU2++
                REPLACE NEMP WITH NUMEMP
                REPLACE FECHA WITH FASI2
                REPLACE CODCTA WITH CODCTA2
                REPLACE NOMAPU WITH CONCEPTO2
                IF AT("EMITIDAS",UPPER(W_IvaR.GR_Iva.Cell(N,7)))<>0
                    REPLACE DEBE WITH W_IvaR.GR_Iva.Cell(N,5)
                ELSE
                    REPLACE HABER WITH W_IvaR.GR_Iva.Cell(N,5)
                ENDIF
                DBCOMMIT()
                DBUNLOCK()
                Suizo_saldocuenta(CODCTA,DEBE,HABER)
            ENDIF
            TOT2:=TOT2+DEBE-HABER
        ENDIF
    NEXT
    IF TOT2<0
        CODCTA2:=47000000
        DEB2:=TOT2*-1
        HAB2:=0
    ELSE
        CODCTA2:=47500000
        DEB2:=0
        HAB2:=TOT2
    ENDIF
    APPEND BLANK
    IF RLOCK()
        REPLACE NASI WITH NASI2
        REPLACE APU WITH APU2++
        REPLACE NEMP WITH NUMEMP
        REPLACE FECHA WITH FASI2
        REPLACE CODCTA WITH CODCTA2
        REPLACE NOMAPU WITH CONCEPTO2
        REPLACE DEBE  WITH DEB2
        REPLACE HABER WITH HAB2
        DBCOMMIT()
        DBUNLOCK()
        Suizo_saldocuenta(CODCTA,DEBE,HABER)
    ENDIF

    MSGINFO("Asiento realizado con exito"+CHR(13)+ ;
        "Asiento: "+LTRIM(STR(NASI2))+CHR(13)+ ;
        "Fecha: "+DIA(FASI2,10) )

Return Nil
