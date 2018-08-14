#include "minigui.ch"

procedure RemIrem()

    DO CASE
    CASE UPPER(PROGRAMA())="SUICONTA"
        RUTAREMESA:=RUTAEMPRESA
        RUTAREMEMP:=RUTAPROGRAMA
    CASE FILE(RUTACONTABLE+"\REMESA.DBF")=.T.
        RUTAREMESA:=RUTACONTABLE
        RUTAREMEMP:=LEFT(RUTACONTABLE,RAT("\",RUTACONTABLE)-1)
    CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)+"\SUICONTA.EXE") = .T.
        RUTAREMEMP:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)
        AbrirDBF("empresa",,,,RUTAREMEMP,1)
        GO BOTT
        RUTAREMESA:=RUTAREMEMP+"\"+RTRIM(RUTA)
    CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)+"\REMESA.EXE") = .T.
        RUTAREMESA:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)
        RUTAREMEMP:=RUTAREMESA
    OTHERWISE
        RUTAREMESA:=RUTAEMPRESA
        RUTAREMEMP:=RUTAPROGRAMA
    ENDCASE

    aREMESAS:=Remesa_alis(RUTAREMESA)

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 900 HEIGHT 480 ;
        TITLE 'Imprimir: Remesa' ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @0,0 LABEL RutaRemesa VALUE RUTAREMESA AUTOSIZE TRANSPARENT INVISIBLE

        @10,10 GRID GR_Remesas ;
            HEIGHT 250 ;
            WIDTH 460 ;
            HEADERS {'Serie','Numero','Remesa','Fecha','Banco','Importe','Documentos','Asiento'} ;
            WIDTHS {0,0,75,80,75,100,50,50 } ;
            ITEMS aREMESAS ;
            VALUE 1 ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999','E'},{'TEXTBOX','NUMERIC'}} ;
            ON HEADCLICK {{|| Remesa_Grid_Ord(1)},{|| Remesa_Grid_Ord(2)},{|| Remesa_Grid_Ord(3)}, ;
            {|| Remesa_Grid_Ord(4)},{|| Remesa_Grid_Ord(5)},{|| Remesa_Grid_Ord(6)}, ;
            {|| Remesa_Grid_Ord(7)},{|| Remesa_Grid_Ord(8)}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
            BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            ON CHANGE RemesaActRec()

        @10,480 GRID GR_Recibos ;
            HEIGHT 250 ;
            WIDTH 410 ;
            HEADERS {'Factura','Cliente','Importe','Vencimiento'} ;
            WIDTHS {50,150,100,80 } ;
            ITEMS {} ;
            COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','DATE'}} ;
            ON HEADCLICK {{|| Remesa_Grid_RecOrd(1)},{|| Remesa_Grid_RecOrd(2)},{|| Remesa_Grid_RecOrd(3)},{|| Remesa_Grid_RecOrd(4)}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER}

        @275,010 LABEL L_Orden VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
        @270,110 RADIOGROUP R_Orden OPTIONS {'Codigo','Cliente','Factura','Vencimiento' } ;
            VALUE 1 WIDTH 85 HORIZONTAL

        @305,010 LABEL L_MarIzq VALUE 'Margen izquierdo' AUTOSIZE TRANSPARENT
        @300,110 SPINNER S_MarIzq RANGE 1,20 VALUE 20 WIDTH 90 HEIGHT 25

        @335,010 LABEL L_MarSup VALUE 'Margen superior' AUTOSIZE TRANSPARENT
        @330,110 SPINNER S_MarSup RANGE 1,300 VALUE 20 WIDTH 90 HEIGHT 25

        @365,010 LABEL L_Copia VALUE 'Copias' AUTOSIZE TRANSPARENT
        @360,110 SPINNER S_Copia RANGE 1,300 VALUE 1 WIDTH 90 HEIGHT 25



        LINW:=310
        COLW:=480
        draw rectangle in window W_Imp1 at LINW,COLW+10 to LINW+2,COLW+390 fillcolor{255,0,0}  //Rojo
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

        @LINW+100,COLW+10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION RemIremI("IMPRESORA")

        @LINW+100,COLW+110 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 90 HEIGHT 25 ;
            ACTION RemIremI("HOJACALC")

        @LINW+100,COLW+210 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

        RemesaActRec()

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION RemIremI(LLAMADA)
    //*** NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
    //*** NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
    //*** FREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,4)
    NSER2:= GetColValue("GR_Remesas","W_Imp1",1)
    NREM2:= GetColValue("GR_Remesas","W_Imp1",2)
    FREM2:= GetColValue("GR_Remesas","W_Imp1",4)

    W_Imp1.Title:="Imprimir: Remesa "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)+" "+DIA(FREM2,10)

    IF FILE("FIN.DBF")
        AbrirDBF("Fin","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    AbrirDBF("REMESA",,,,W_Imp1.RutaRemesa.value)
    COPY TO FIN FOR SERIE=NSER2 .AND. NREM=NREM2
    AbrirDBF("FIN","SIN_INDICE")

    DO CASE
    CASE W_Imp1.R_Orden.Value=1
        INDEX ON STR(CODCTA,8)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=2
        INDEX ON UPPER(NOMCTA)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=3
        INDEX ON NFRA+STR(CODCTA,8) TO FIN
    CASE W_Imp1.R_Orden.Value=4
        INDEX ON DTOS(FVTO)+STR(CODCTA,8)+NFRA TO FIN
    ENDCASE

    GO TOP
    IF LASTREC()=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF

    IF LLAMADA="HOJACALC"
        RemIremI_Calc()
    ELSE
        RemIremI_Imp()
    ENDIF

Return Nil

STATIC FUNCTION RemIremI_Imp()
    local oprint

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",4)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(TituloqImp(W_Imp1.Title))
    oprint:setpreviewsize(2)                        // tamaño del preview
    FOR nCopia=1 TO W_Imp1.S_Copia.Value
        oprint:beginpage()

        SELEC FIN
        GO TOP
        AbrirDBF("CUENTAS",,,,W_Imp1.RutaRemesa.value)
        SEEK FIN->CODBAN
        aDATBANCO:={CODCTA,NOMCTA,CTA_BAN_SUIZO(BANCTA,24)}

        MARGIZQ:=W_Imp1.S_MarIzq.Value
        MARGSUP:=W_Imp1.S_MarSup.Value

        TOT1:=0
        TOTREC2:=0
        PAG:=0
        LIN:=0
        SELEC FIN
        GO TOP
        DO WHILE .NOT. EOF()

            IF LIN>=260 .OR. PAG=0
                IF PAG<>0
                    oprint:printdata(LIN,MARGIZQ+140,'Suma',"times new roman",8,.F.,,"L",)
                    oprint:printdata(LIN,MARGIZQ+170,MIL(TOT1,15,2),"times new roman",8,.F.,,"R",)
                    LIN:=LIN+5
                    oprint:printdata(LIN+5,MARGIZQ+85,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
                    oprint:endpage()
                    oprint:beginpage()
                ENDIF
                PAG=PAG+1

                LIN:=MARGSUP
                oprint:printdata(LIN  ,MARGIZQ+0,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
                oprint:printdata(LIN  ,MARGIZQ+170,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
                oprint:printdata(LIN+5,MARGIZQ+0,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

                oprint:printdata(LIN+5 ,MARGIZQ+85,NOMEMPRESA,"times new roman",12,.F.,,"C",)
                oprint:printdata(LIN+12,MARGIZQ+0,"Banco: "+STR(aDATBANCO[1])+" "+aDATBANCO[3],"times new roman",12,.F.,,"L",)
                oprint:printdata(LIN+12,MARGIZQ+170,TituloqImp(W_Imp1.Title),"times new roman",12,.F.,,"R",)

                LIN:=LIN+20
                oprint:printdata(LIN,MARGIZQ+ 12,'Factura',"times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,MARGIZQ+ 14,'Codigo',"times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,MARGIZQ+ 27,'Descripcion',"times new roman",8,.F.,,"L",)
                IF NSER2=3 .OR. NSER2=4
                    oprint:printdata(LIN,MARGIZQ+ 75,'Banco',"times new roman",8,.F.,,"L",)
                    oprint:printdata(LIN,MARGIZQ+109,'Nº talon',"times new roman",8,.F.,,"L",)
                ELSE
                    oprint:printdata(LIN,MARGIZQ+ 82,'Cuenta bancaria',"times new roman",8,.F.,,"L",)
                ENDIF
                oprint:printdata(LIN,MARGIZQ+132,'Vencimiento',"times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,MARGIZQ+170,'Importe',"times new roman",8,.F.,,"R",)
                oprint:printline(LIN+4,MARGIZQ+0,LIN+4,MARGIZQ+170,,0.5)

                LIN:=LIN+4
            ENDIF

            oprint:printdata(LIN,MARGIZQ+ 12,RTRIM(NFRA),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,MARGIZQ+ 14,CODCTA,"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,MARGIZQ+ 27,RTRIM(NOMCTA),"times new roman",8,.F.,,"L",)
            IF NSER2=3 .OR. NSER2=4
                oprint:printdata(LIN,MARGIZQ+ 75,CTA_BAN_SUIZO(BANCTA,24),"times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,MARGIZQ+109,TALON,"times new roman",8,.F.,,"L",)
            ELSE
                oprint:printdata(LIN,MARGIZQ+ 82,CTA_BAN_SUIZO(BANCTA,24),"times new roman",8,.F.,,"L",)
            ENDIF
            oprint:printdata(LIN,MARGIZQ+132,DIA(FVTO,10),"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,MARGIZQ+170,MIL(IMPORTE,15,2),"times new roman",8,.F.,,"R",)

            TOT1:=TOT1+IMPORTE
            TOTREC2++

            LIN:=LIN+4
            SKIP

        ENDDO

        oprint:printdata(LIN,MARGIZQ+20,MIL(TOTREC2,15,0)+' Recibos',"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,MARGIZQ+140,'Total',"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,MARGIZQ+170,MIL(TOT1,15,2),"times new roman",8,.F.,,"R",)

        oprint:endpage()
    NEXT
    oprint:enddoc()
    oprint:RELEASE()


Return Nil




STATIC FUNCTION RemIremI_Calc()
    local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

    AbrirDBF("CUENTAS",,,,W_Imp1.RutaRemesa.value)
    SEEK FIN->CODBAN
    aDATBANCO:={CODCTA,NOMCTA,CTA_BAN_SUIZO(BANCTA,24)}
    SELEC FIN


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

    LIN:=3

    oSheet:getCellByPosition(0,LIN):SetString("Factura")
    oSheet:getCellByPosition(1,LIN):SetString("Codigo")
    oSheet:getCellByPosition(1,LIN):HoriJustify:=3
    oSheet:getCellByPosition(2,LIN):SetString("Descripcion")
    IF NSER2=3 .OR. NSER2=4
        oSheet:getCellByPosition(3,LIN):SetString("Codigo banco y nº talon")
    ELSE
        oSheet:getCellByPosition(3,LIN):SetString("Cuenta bancaria")
    ENDIF
    oSheet:getCellByPosition(4,LIN):SetString("Vencimiento")
    oSheet:getCellByPosition(5,LIN):SetString("Importe")
    oSheet:getCellRangeByPosition(5,LIN,5,LIN):HoriJustify:=3
    oSheet:getCellRangeByPosition(0,LIN,5,LIN):CharWeight:=150  //NEGRITA
    aMiColor:=MiColor("AMARILLOPALIDO")
    oSheet:getCellRangeByPosition(0,LIN,5,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

    LIN++
    LIN1:=LIN+1


    TOT1:=0
    TOTREC2:=0
    PAG:=0
    SELEC FIN
    GO TOP
    FREM2:=FREM
    DO WHILE .NOT. EOF()
        DO EVENTS
        oSheet:getCellByPosition(0,LIN):SetString(RTRIM(NFRA))
        oSheet:getCellByPosition(1,LIN):SetValue(CODCTA)
        oSheet:getCellByPosition(2,LIN):SetString(RTRIM(NOMCTA))
        IF NSER2=3 .OR. NSER2=4
            oSheet:getCellByPosition(3,LIN):SetString(LEFT(CTA_BAN_SUIZO(BANCTA,24),9)+" "+RTRIM(TALON))
        ELSE
            oSheet:getCellByPosition(3,LIN):SetString(CTA_BAN_SUIZO(BANCTA,24))
        ENDIF
        oSheet:getCellByPosition(4,LIN):SetString(DIA(FVTO,10))
        oSheet:getCellByPosition(5,LIN):SetValue(IMPORTE)
        oSheet:getCellRangeByPosition(2,LIN,5,LIN):NumberFormat:=4  //#.##0,00

        TOTREC2++

        LIN++
        SKIP

    ENDDO

    oSheet:getCellByPosition(1,LIN):SetValue(TOTREC2)
    oSheet:getCellByPosition(2,LIN):SetString('Recibos')
    oSheet:getCellByPosition(4,LIN):SetString("Total")
    oSheet:getCellByPosition(5,LIN):SetFormula("=SUM("+LetraExcel(6)+LTRIM(STR(LIN1))+":"+LetraExcel(6)+LTRIM(STR(LIN))+")")
    oSheet:getCellRangeByPosition(5,LIN,5,LIN):NumberFormat:=4  //#.##0,00
    oSheet:getCellRangeByPosition(1,LIN,5,LIN):CharWeight:=150  //NEGRITA
    aMiColor:=MiColor("VERDEPALIDO")
    oSheet:getCellRangeByPosition(1,LIN,5,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

    oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)

    oSheet:getCellByPosition(0,0):SetString(DIA(DATE(),10))
    oSheet:getCellByPosition(1,0):SetString(Nombre_Programa())
    oSheet:getCellByPosition(3,0):SetString(NOMEMPRESA)
    oSheet:getCellByPosition(0,1):SetString("Banco")
    oSheet:getCellByPosition(1,1):SetString(STR(aDATBANCO[1]))
    oSheet:getCellByPosition(2,1):SetString(aDATBANCO[3])

    oSheet:getCellByPosition(4,0):SetString(TituloqImp(W_Imp1.Title))
    oSheet:getCellByPosition(4,0):CharWeight:=150   //NEGRITA
    oSheet:getCellByPosition(4,1):SetString(DIA(FREM2,10))
    oSheet:getCellByPosition(4,1):CharWeight:=150   //NEGRITA


    QuitarEspera()

    Return Nil
