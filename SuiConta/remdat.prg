#include "minigui.ch"

Function RemDat()
    LOCAL Color_RemDat1:= { |x,nItem| IF( x[9]=.T. , RGB(255,  0,  0), RGB(255,255,255) ) }

    IF IsWindowDefined(WinRemDat)=.T.
        DECLARE WINDOW WinRemDat
        WinRemDat.restore
        WinRemDat.Show
        WinRemDat.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinRemDat ;
        AT 50,0     ;
        WIDTH 900  ;
        HEIGHT 600 ;
        TITLE "Remesas" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("AZULCLARO") ;
        ON RELEASE CloseTables()

        @000,100 TEXTBOX T_CodCta WIDTH 50 VALUE "" RIGHTALIGN INVISIBLE
        @010,100 TEXTBOX T_NomCta WIDTH 250 VALUE "" INVISIBLE
        @020,100 TEXTBOX T_NumFac WIDTH 80 VALUE 0 NUMERIC RIGHTALIGN INVISIBLE
        @030,100 TEXTBOX T_SerFac WIDTH 20 VALUE '' INVISIBLE


        @015,10 LABEL L_SerRem VALUE 'Tipo remesa' AUTOSIZE TRANSPARENT
        @010,100 RADIOGROUP R_SerRem ;
            OPTIONS REM_NOM3() ON CHANGE RemDatVerTipTra() ;
            VALUE 1 WIDTH 100 SPACING 20 NOTABSTOP

        @135,10 LABEL L_TipTra VALUE 'Tipo transf.' AUTOSIZE TRANSPARENT

        @130,100 RADIOGROUP R_TipTra OPTIONS {"Nomina","Pension","Otros"} ;
            VALUE 1 WIDTH 70 HORIZONTAL NOTABSTOP

        @015,310 LABEL L_NumRem VALUE 'Numero remesa' AUTOSIZE TRANSPARENT
        @010,420 TEXTBOX T_NumRem WIDTH 100 NUMERIC RIGHTALIGN

        @010,525 BUTTONEX Bt_BuscarRem CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION (br_rem1(WinRemDat.R_SerRem.Value,WinRemDat.T_NumRem.Value,"WinRemDat","R_SerRem","T_NumRem") , RemDatActualiz("BOTON") ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @045,310 LABEL L_FecRem VALUE 'Fecha remesa' AUTOSIZE TRANSPARENT
        @040,420 DATEPICKER D_FecRem WIDTH 100 VALUE DATE()

        @070,310 BUTTONEX Bt_BuscarCue CAPTION 'Banco' ICON icobus('buscar') ;
            ACTION br_cue1(VAL(WinRemDat.T_CodCtaBan.Value),"WinRemDat","T_CodCtaBan","T_NomCtaBan") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @070,420 TEXTBOX T_CodCtaBan WIDTH 100 VALUE "57200001" MAXLENGTH 8
        @070,520 TEXTBOX T_NomCtaBan WIDTH 200 VALUE Codigos_NomCta(WinRemDat.T_CodCtaBan.Value)

        @105,310 LABEL L_Asiento VALUE 'Asiento' AUTOSIZE TRANSPARENT
        @100,420 TEXTBOX T_Asiento WIDTH 75 NUMERIC RIGHTALIGN NOTABSTOP
        @100,520 BUTTON Bt_Asiento CAPTION 'Ver asiento' ;
            ACTION br_suizoasi(RUTAEMPRESA,WinRemDat.T_Asiento.Value,NUMEMP) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @010,720 BUTTONEX Bt_AltaVto CAPTION 'Alta vencimientos' ICON icobus('nuevo') ;
            ACTION RemDat_AltaVto1() ;
            WIDTH 150 HEIGHT 25 NOTABSTOP

/*
        @155,010 GRID GR_Remesa ;
            HEIGHT 340 ;
            WIDTH 870 ;
            HEADERS {'Linea','Concepto','Cuenta','Nombre','Banco','Talon','Fecha Vto','Importe','Eliminar','Empresa','Registro'} ;
            WIDTHS {0,70,70,200,170,150,100,100,0,0,0 } ;
            ITEMS {} ;
            DYNAMICBACKCOLOR  { Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1,Color_RemDat1 } ;
            EDIT INPLACE {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','CHARACTER'},{'DATEPICKER','DROPDOWN'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'CHECKBOX','Si','No'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'}} ;
            COLUMNVALID { ,,,,{ || RemDat_CtaBan1()},,,{ || RemDat_Importe()} } ;
            COLUMNWHEN  { { || .F.},{ || RemDat_BusFra1()},{ || RemDat_BusCue1()},,,, } ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT, ;
            BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            CELLNAVIGATION
*/

        @155,010 GRID GR_Remesa ;
            HEIGHT 340 ;
            WIDTH 870 ;
            HEADERS {'Linea','Concepto','Cuenta','Nombre','Banco','Talon','Fecha Vto','Importe','Eliminar','Empresa','Registro'} ;
            WIDTHS {0,70,70,200,170,150,100,100,0,0,0 } ;
            ITEMS {} ;
            EDIT INPLACE ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','CHARACTER'},{'DATEPICKER','DROPDOWN'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'CHECKBOX','Si','No'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'}} ;
            VALID { ,,,,{ || RemDat_CtaBan1()},,,{ || RemDat_Importe()} } ;
            COLUMNWHEN  { { || .F.},{ || RemDat_BusFra1()},{ || RemDat_BusCue1()},,,, } ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT, ;
            BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            NAVIGATEBYCELL

**      Menu_Grid("WinRemDat","GR_Remesa","MENU",,{"COPREGPP","COPTABPP"})

        DEFINE CONTEXT MENU CONTROL GR_Remesa OF WinRemDat
        ITEM "Añadir linea"            ACTION RemDatNueLin()
        ITEM "Eliminar linea"          ACTION RemDatEliLin()
        ITEM "Recuperar linea"         ACTION RemDatEliLin()
        SEPARATOR
        ITEM "Copiar celda"            ACTION Menu_Grid("WinRemDat","GR_Remesa","CopiarCelda",2)
        ITEM "Copiar registro"         ACTION Menu_Grid("WinRemDat","GR_Remesa","CopiarRegistro")
        ITEM "Copiar tabla"            ACTION Menu_Grid("WinRemDat","GR_Remesa","CopiarTabla")
    END MENU

    @505,720 LABEL L_Total VALUE 'Total' AUTOSIZE TRANSPARENT
    @500,770 TEXTBOX T_Total WIDTH 110 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


    @540,010 BUTTONEX Bt_Nuevo CAPTION 'Nuevo' ICON icobus('nuevo') ;
        ACTION RemDatNuevo("NUEVO") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') ;
        ACTION RemDatModif(.T.) ;
        WIDTH 95 HEIGHT 25

    @540,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') ;
        ACTION RemDatGuardar() ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') ;
        ACTION RemDatActualiz("CANCELAR") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,410 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
        ACTION RemDatEliminar() ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
        ACTION WinRemDat.Release ;
        WIDTH 95 HEIGHT 25 NOTABSTOP


    RemDatModif(.F.)

    END WINDOW
    VentanaCentrar("WinRemDat","Ventana1","Alinear")
    CENTER WINDOW WinRemDat
    ACTIVATE WINDOW WinRemDat

Return Nil

STATIC FUNCTION RemDat_CtaBan1()
    local cVal := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex

    IF CTA_BAN(This.CellValue,1)=0
        MsgStop("Cuenta bancaria erronea"+HB_OsNewLine()+ ;
            This.CellValue+HB_OsNewLine() )
    ELSE
        This.CellValue:=CTA_BAN_SUIZO(This.CellValue,24)
    ENDIF
Return Nil

STATIC FUNCTION RemDat_BusFra1()
    local cVal := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex

    Titulo     :='Asignar numero de factura'
    aLabels    :={'Forma asignar'}
    aFormats   :={ {'Directamente','Factura emitidas','Factura recibidas'} }
    aInitValues:={ IF(WinRemDat.R_SerRem.Value<5,2,3) }

    aBotones   :={"Aceptar","Cancelar"}
    aTipoAsig  :=InputWindow(Titulo, aLabels , aInitValues , aFormats,WinRemDat.Row+260,WinRemDat.Col+410 , .T. , aBotones)

    DO CASE
    CASE aTipoAsig[1]=1
        Titulo     :='Asignar numero de factura'
        aLabels    :={'Factura'}
        aFormats   :={ 15 }
        aInitValues:={ cVal }
        aBotones   :={"Aceptar","Cancelar"}
        aResFac    :=InputWindow(Titulo, aLabels , aInitValues , aFormats,WinRemDat.Row+260,WinRemDat.Col+410 , .T. , aBotones)
        IF aResFac[1] <> Nil
            WinRemDat.GR_Remesa.Cell(nRow,2):=aResFac[1]
            WinRemDat.GR_Remesa.Cell(nRow,10):=0
            WinRemDat.GR_Remesa.Cell(nRow,11):=0
        ENDIF
    CASE aTipoAsig[1]=2
        SFAC2:=SUBSTR(cVal,AT("-",cVal)+1,1)
        NFAC2:=LEFT(cVal,AT("-",cVal)-1)
        IF VAL(NFAC2)<>0
            WinRemDat.T_SerFac.Value:=SFAC2
            WinRemDat.T_NumFac.Value:=VAL(NFAC2)
        ELSE
            WinRemDat.T_SerFac.Value:=" "
            WinRemDat.T_NumFac.Value:=0
        ENDIF
        aNFAC:=br_emibus(RUTAPROGRAMA,NUMEMP,WinRemDat.T_SerFac.Value,WinRemDat.T_NumFac.Value,WinRemDat.GR_Remesa.Cell(nRow,4))
        IF aNFAC[1]<>0
            AbrirDBF("FAC92",,,,aNFAC[2],1)
            GO aNFAC[1]
            IF NFAC<>WinRemDat.T_NumFac.Value .OR. SERFAC<>WinRemDat.T_SerFac.Value
                WinRemDat.GR_Remesa.Cell(nRow,2):=LTRIM(STR(NFAC))+"-"+SERFAC
                WinRemDat.GR_Remesa.Cell(nRow,3):=CODCTA
                WinRemDat.GR_Remesa.Cell(nRow,4):=RTRIM(CLIENTE)
                WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value-WinRemDat.GR_Remesa.Cell(nRow,8)+IF(PEND<>0,PEND,TFAC)
                WinRemDat.GR_Remesa.Cell(nRow,8):=IF(PEND<>0,PEND,TFAC)
                WinRemDat.GR_Remesa.Cell(nRow,10):=aNFAC[3]
                WinRemDat.GR_Remesa.Cell(nRow,11):=NFAC
                AbrirDBF("CUENTAS",,,,,1)
                SEEK WinRemDat.GR_Remesa.Cell(nRow,3)
                IF .NOT. EOF()
                    WinRemDat.GR_Remesa.Cell(nRow,5):=CTA_BAN_SUIZO(BANCTA,24)
                    IF WinRemDat.R_SerRem.Value<>3 .AND. WinRemDat.R_SerRem.Value<>4
                        WinRemDat.GR_Remesa.Cell(nRow,6):=CODPOS+"-"+RTRIM(POBCTA)
                    ENDIF
                ENDIF
            ENDIF
            RemDat_BusDatAnt(nRow)
        ENDIF
    CASE aTipoAsig[1]=3
        WinRemDat.T_SerFac.Value:=" "
        WinRemDat.T_NumFac.Value:=0
        aNFAC:=br_recbus(RUTAPROGRAMA,NUMEMP,WinRemDat.T_NumFac.Value,WinRemDat.GR_Remesa.Cell(nRow,4))
        IF aNFAC[1]<>0
            AbrirDBF("FACREB",,,,aNFAC[2],1)
            GO aNFAC[1]
**         IF NFAC<>WinRemDat.T_NumFac.Value
            WinRemDat.GR_Remesa.Cell(nRow,2):=RTRIM(REF)
            WinRemDat.GR_Remesa.Cell(nRow,3):=CODIGO
            WinRemDat.GR_Remesa.Cell(nRow,4):=RTRIM(NOMCTA)
            WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value-WinRemDat.GR_Remesa.Cell(nRow,8)+IF(PEND<>0,PEND,TFAC)
            WinRemDat.GR_Remesa.Cell(nRow,8):=IF(PEND<>0,PEND,TFAC)
            WinRemDat.GR_Remesa.Cell(nRow,10):=aNFAC[3]
            WinRemDat.GR_Remesa.Cell(nRow,11):=NREG
            AbrirDBF("CUENTAS",,,,,1)
            SEEK WinRemDat.GR_Remesa.Cell(nRow,3)
            IF .NOT. EOF()
                WinRemDat.GR_Remesa.Cell(nRow,5):=CTA_BAN_SUIZO(BANCTA,24)
                IF WinRemDat.R_SerRem.Value<>3 .AND. WinRemDat.R_SerRem.Value<>4
                    WinRemDat.GR_Remesa.Cell(nRow,6):=CODPOS+"-"+RTRIM(POBCTA)
                ENDIF
            ENDIF
**         ENDIF
            RemDat_BusDatAnt(nRow)
        ENDIF
    ENDCASE

RETURN .F.



STATIC FUNCTION RemDat_BusDatAnt(nRow)
    AbrirDBF("REMESA")
    IF LEN(RTRIM(STRTRAN(WinRemDat.GR_Remesa.Cell(nRow,5),"-"," ")))=0  //BANCTA
        SET FILTER TO CODCTA=WinRemDat.GR_Remesa.Cell(nRow,3)
        FREMCLI:=CTOD("  -  -  ")
        GO TOP
        IF .NOT. EOF()
            DO WHILE .NOT. EOF()
                IF FREMCLI<WinRemDat.D_FecRem.Value
                    FREMCLI:=WinRemDat.D_FecRem.Value
                    WinRemDat.GR_Remesa.Cell(nRow,5):=CTA_BAN_SUIZO(BANCTA,24)
                    IF LEN(RTRIM(WinRemDat.GR_Remesa.Cell(nRow,6)))=0
                        WinRemDat.GR_Remesa.Cell(nRow,6):=RTRIM(TALON)
                    ENDIF
                ENDIF
                SKIP
            ENDDO
        ELSE
            NUMEJE2:=EJERANY-1
            DO WHILE NUMEJE2>=EJERANY-5 .AND. LEN(RTRIM(STRTRAN(WinRemDat.GR_Remesa.Cell(nRow,5),"-"," ")))=0
                RUTA2:=EMPRUTANY(NUMEMP,NUMEJE2--)
                IF .NOT. FILE(RUTA2+"\REMESA.DBF") .OR. .NOT. FILE(RUTA2+"\REMESA.CDX")
                    EXIT
                ELSE
                    AbrirDBF("REMESA",,,,RUTA2)
                    SET FILTER TO CODCTA=WinRemDat.GR_Remesa.Cell(nRow,3)
                    GO TOP
                    DO WHILE .NOT. EOF()
                        IF FREMCLI<WinRemDat.D_FecRem.Value
                            FREMCLI:=WinRemDat.D_FecRem.Value
                            WinRemDat.GR_Remesa.Cell(nRow,5):=CTA_BAN_SUIZO(BANCTA,24)
                            IF LEN(RTRIM(WinRemDat.GR_Remesa.Cell(nRow,6)))=0
                                WinRemDat.GR_Remesa.Cell(nRow,6):=RTRIM(TALON)
                            ENDIF
                        ENDIF
                        SKIP
                    ENDDO
                ENDIF
            ENDDO
            AbrirDBF("REMESA")
        ENDIF
        SET FILTER TO
    ENDIF
Return Nil

STATIC FUNCTION RemDat_BusCue1()
    local CODCUE2 := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex
    IF nRow>1 .AND. CODCUE2=0
        CODCUE2:=WinRemDat.GR_Remesa.Cell(nRow-1,3)
    ENDIF
    WinRemDat.T_CodCta.Value:=""
    br_cue1(CODCUE2,"WinRemDat","T_CodCta","T_NomCta")
    IF LEN(LTRIM(WinRemDat.T_CodCta.Value))=0
        RETURN .F.
    ENDIF
    WinRemDat.GR_Remesa.Cell(nRow,3):=VAL(WinRemDat.T_CodCta.Value)
    WinRemDat.GR_Remesa.Cell(nRow,4):=WinRemDat.T_NomCta.Value
    IF LEN(RTRIM(WinRemDat.GR_Remesa.Cell(nRow,5)))=0
        SiCambiar:=.T.
    ELSE
        SiCambiar:=MsgYesNo("¿Desea cambiar la cuenta bancaria?")
    ENDIF
    IF SiCambiar=.T.
        AbrirDBF("CUENTAS",,,,,1)
        SEEK WinRemDat.GR_Remesa.Cell(nRow,3)
        IF .NOT. EOF()
            WinRemDat.GR_Remesa.Cell(nRow,5):=CTA_BAN_SUIZO(BANCTA,24)
            IF WinRemDat.R_SerRem.Value<>3 .AND. WinRemDat.R_SerRem.Value<>4
                WinRemDat.GR_Remesa.Cell(nRow,6):=CODPOS+"-"+RTRIM(POBCTA)
            ENDIF
        ENDIF
        RemDat_BusDatAnt(nRow)
    ENDIF

RETURN .F.



STATIC FUNCTION RemDat_Importe()
    local nVal := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex
    WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value-WinRemDat.GR_Remesa.Cell(nRow,nCol)+nVal
Return Nil



STATIC FUNCTION RemDatNuevo(LLAMADA)
**   WinRemDat.R_SerRem.Value:=1
    WinRemDat.T_NumRem.Value:=0
    WinRemDat.R_TipTra.Value:=1
    RemDatVerTipTra()
    WinRemDat.D_FecRem.Value:=DATE()
    WinRemDat.T_CodCtaBan.Value:="57200001"
    WinRemDat.T_NomCtaBan.Value:=Codigos_NomCta(WinRemDat.T_CodCtaBan.Value)
    WinRemDat.T_Asiento.Value:=0
    WinRemDat.T_Total.Value:=0
    WinRemDat.GR_Remesa.DeleteAllItems
    WinRemDat.GR_Remesa.Refresh
    RemDatModif(.T.)
Return Nil

STATIC FUNCTION RemDatVerTipTra()
    WinRemDat.R_TipTra.Visible:=IF(WinRemDat.R_SerRem.Value=6,.T.,.F.)
    WinRemDat.L_TipTra.Visible:=IF(WinRemDat.R_SerRem.Value=6,.T.,.F.)
    WinRemDat.Bt_AltaVto.Visible:=IF(WinRemDat.R_SerRem.Value=1,.T.,.F.)
Return Nil

STATIC FUNCTION RemDatModif(Modif)
    SoloVer:=IF(Modif=.T.,.F.,.T.)
    WinRemDat.R_SerRem.Enabled:=SoloVer
    WinRemDat.T_NumRem.ReadOnly:=.T.
    WinRemDat.R_TipTra.Enabled:=Modif
    WinRemDat.Bt_AltaVto.Enabled:=Modif
    RemDatVerTipTra()
    WinRemDat.Bt_BuscarRem.Enabled:=SoloVer
    WinRemDat.D_FecRem.Enabled:=Modif
    WinRemDat.Bt_BuscarCue.Enabled:=Modif
    WinRemDat.T_CodCtaBan.ReadOnly:=SoloVer
    WinRemDat.T_NomCtaBan.ReadOnly:=SoloVer
    WinRemDat.T_Asiento.ReadOnly:=.T.
    WinRemDat.Bt_Asiento.Enabled:=SoloVer
    WinRemDat.T_Total.ReadOnly:=.T.
    WinRemDat.GR_Remesa.Enabled:=Modif
    IF WinRemDat.R_SerRem.Value=3 .OR. WinRemDat.R_SerRem.Value=4
        WinRemDat.GR_Remesa.Header(6):="Talon"
    ELSE
        WinRemDat.GR_Remesa.Header(6):="Poblacion"
    ENDIF

    IF Modif = .T.
        WinRemDat.Bt_Nuevo.Enabled   := .F.
        WinRemDat.Bt_Modif.Enabled   := .F.
        WinRemDat.Bt_Guardar.Enabled := .T.
        WinRemDat.Bt_Cancelar.Enabled:= .T.
        WinRemDat.Bt_Eliminar.Enabled:= .T.
        WinRemDat.Bt_Salir.Enabled   := .F.
        ON KEY ADD OF WinRemDat        ACTION RemDatNueLin("NUEVO")
        WinRemDat.R_SerRem.SetFocus
    ELSE
        WinRemDat.T_Asiento.Enabled  := .T.
        WinRemDat.Bt_Nuevo.Enabled   := .T.
        WinRemDat.Bt_Modif.Enabled   := IF(WinRemDat.T_NumRem.Value=0,.F.,.T.)
        WinRemDat.Bt_Guardar.Enabled := .F.
        WinRemDat.Bt_Cancelar.Enabled:= .F.
        WinRemDat.Bt_Eliminar.Enabled:= .F.
        WinRemDat.Bt_Salir.Enabled   := .T.
**   RELEASE KEY ADD OF WinRemDat
        ON KEY ADD OF WinRemDat        ACTION NIL
        WinRemDat.GR_Remesa.SetFocus
    ENDIF
Return Nil

STATIC FUNCTION RemDatGuardar()
    IF WinRemDat.D_FecRem.Value<DMA1 .OR. WinRemDat.D_FecRem.Value>DMA2
        IF MSGYESNO("La fecha de la remesa esta fuera del ejercicio"+HB_OsNewLine()+"¿Desea grabarla?")=.F.
            RETURN
        ENDIF
    ENDIF

    CuentaMal:=0
    FOR N=1 TO WinRemDat.GR_Remesa.ItemCount
        DO EVENTS
        IF CTA_BAN(WinRemDat.GR_Remesa.Cell(N,5),1)=0
            WinRemDat.GR_Remesa.Value:={N,5}
            MsgStop("Cuenta bancaria erronea"+HB_OsNewLine()+ ;
                "Concepto: "+WinRemDat.GR_Remesa.Cell(N,2)+HB_OsNewLine()+ ;
                "Cuenta: "+LTRIM(STR(WinRemDat.GR_Remesa.Cell(N,3)))+" "+WinRemDat.GR_Remesa.Cell(N,4)+HB_OsNewLine()+ ;
                "Banco: "+WinRemDat.GR_Remesa.Cell(N,5)+HB_OsNewLine()+ ;
                "Talon: "+WinRemDat.GR_Remesa.Cell(N,6)+HB_OsNewLine()+ ;
                "Vencimiento: "+DIA(WinRemDat.GR_Remesa.Cell(N,7),10)+HB_OsNewLine()+ ;
                "Importe: "+LTRIM(MIL(WinRemDat.GR_Remesa.Cell(N,8),15,2)) )
            CuentaMal:=1
        ENDIF
    NEXT
    IF CuentaMal=1
        IF MSGYESNO("Existen cuentas bancarias erroneas"+HB_OsNewLine()+"¿Desea grabarla?")=.F.
            WinRemDat.GR_Remesa.SetFocus
            RETURN
        ENDIF
    ENDIF

    AbrirDBF("REMESA")
    IF WinRemDat.T_NumRem.Value=0
        SET SOFTSEEK ON
        SEEK (WinRemDat.R_SerRem.Value+1)*100000
        SKIP -1
        IF WinRemDat.R_SerRem.Value<>SERIE
            WinRemDat.T_NumRem.Value:=1
        ELSE
            WinRemDat.T_NumRem.Value:=NREM+1
        ENDIF
        SET SOFTSEEK OFF
    ELSE
        SEEK (WinRemDat.R_SerRem.Value*100000)+WinRemDat.T_NumRem.Value
        DO WHILE SERIE=WinRemDat.R_SerRem.Value .AND. NREM=WinRemDat.T_NumRem.Value .AND. .NOT. EOF()
            DO EVENTS
            IF RLOCK()
                DELETE
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
            SKIP
        ENDDO
    ENDIF

    PonerEspera("Guardando remesa....")

    AbrirDBF("REMESA")
    LREM2:=0
    FOR N=1 TO WinRemDat.GR_Remesa.ItemCount
        DO EVENTS
        IF WinRemDat.GR_Remesa.Cell(N,9)=.F.
            APPEND BLANK
            IF RLOCK()
                REPLACE SERIE WITH WinRemDat.R_SerRem.Value
                REPLACE NREM WITH WinRemDat.T_NumRem.Value
                REPLACE FREM WITH WinRemDat.D_FecRem.Value
                REPLACE CODBAN WITH VAL(WinRemDat.T_CodCtaBan.Value)
                REPLACE TRANOMI WITH WinRemDat.R_TipTra.Value
                REPLACE NASI WITH WinRemDat.T_Asiento.Value
                REPLACE NFRA WITH WinRemDat.GR_Remesa.Cell(N,2)
                REPLACE CODCTA WITH WinRemDat.GR_Remesa.Cell(N,3)
                REPLACE NOMCTA WITH WinRemDat.GR_Remesa.Cell(N,4)
                REPLACE BANCTA WITH WinRemDat.GR_Remesa.Cell(N,5)
                REPLACE TALON WITH WinRemDat.GR_Remesa.Cell(N,6)
                REPLACE FVTO WITH WinRemDat.GR_Remesa.Cell(N,7)
                REPLACE IMPORTE WITH WinRemDat.GR_Remesa.Cell(N,8)
                REPLACE NEMPREG WITH WinRemDat.GR_Remesa.Cell(N,10)
                REPLACE NREG    WITH WinRemDat.GR_Remesa.Cell(N,11)
                IF WinRemDat.GR_Remesa.Cell(N,1)=0
                    LREM2++
                ELSE
                    LREM2:=WinRemDat.GR_Remesa.Cell(N,1)
                ENDIF
                REPLACE LREM WITH LREM2
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
        ELSE
            AbrirDBF("REMLIN")
            SEEK STR(WinRemDat.R_SerRem.Value,2)+STR(WinRemDat.T_NumRem.Value,5)+STR(WinRemDat.GR_Remesa.Cell(N,1),3)
            DO WHILE SERIE=WinRemDat.R_SerRem.Value .AND. NREM=WinRemDat.T_NumRem.Value .AND. WinRemDat.GR_Remesa.Cell(N,1)=LREM .AND. .NOT. EOF()
                DO EVENTS
                IF RLOCK()
                    DELETE
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
                SKIP
            ENDDO
            AbrirDBF("REMESA")
        ENDIF
    NEXT

    QuitarEspera()

    MsgInfo('Los datos han sido guardados','Datos guardados')
    RemDatModif(.F.)

Return Nil

STATIC FUNCTION RemDatActualiz(LLAMADA)
    TOTREM2:=0

    WinRemDat.GR_Remesa.DeleteAllItems
    AbrirDBF("REMESA")
    SEEK (WinRemDat.R_SerRem.Value*100000)+WinRemDat.T_NumRem.Value
    IF .NOT. EOF()
        WinRemDat.D_FecRem.Value:=FREM
        WinRemDat.T_CodCtaBan.Value:=LTRIM(STR(CODBAN))
        WinRemDat.T_NomCtaBan.Value:=Codigos_NomCta(WinRemDat.T_CodCtaBan.Value)
        WinRemDat.T_Asiento.Value:=NASI
    ELSE
        WinRemDat.D_FecRem.Value:=DATE()
        WinRemDat.T_CodCtaBan.Value:=""
        WinRemDat.T_NomCtaBan.Value:=""
        WinRemDat.T_Asiento.Value:=0
    ENDIF
    DO WHILE SERIE=WinRemDat.R_SerRem.Value .AND. NREM=WinRemDat.T_NumRem.Value .AND. .NOT. EOF()
        DO EVENTS
        WinRemDat.GR_Remesa.AddItem({LREM,RTRIM(NFRA),CODCTA,RTRIM(NOMCTA),CTA_BAN_SUIZO(BANCTA,24),RTRIM(TALON),FVTO,IMPORTE,.F.,NEMPREG,NREG})
        TOTREM2:=TOTREM2+IMPORTE
        SKIP
    ENDDO
    WinRemDat.GR_Remesa.Value:={WinRemDat.GR_Remesa.ItemCount,2}
    WinRemDat.GR_Remesa.Refresh
    WinRemDat.T_Total.Value:=TOTREM2
    RemDatModif(.F.)
Return Nil

STATIC FUNCTION RemDatEliminar()
    IF MSGYESNO("¿Desea eliminar la remesa?")=.F.
        RETURN
    ENDIF

    PonerEspera("Procesando los datos....")
    AbrirDBF("REMESA")
    SEEK (WinRemDat.R_SerRem.Value*100000)+WinRemDat.T_NumRem.Value
    DO WHILE SERIE=WinRemDat.R_SerRem.Value .AND. NREM=WinRemDat.T_NumRem.Value .AND. .NOT. EOF()
        DO EVENTS
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
        SKIP
    ENDDO
    AbrirDBF("REMLIN")
    SEEK STR(WinRemDat.R_SerRem.Value,2)+STR(WinRemDat.T_NumRem.Value,5)
    DO WHILE SERIE=WinRemDat.R_SerRem.Value .AND. NREM=WinRemDat.T_NumRem.Value .AND. .NOT. EOF()
        DO EVENTS
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
        SKIP
    ENDDO
    RemDatActualiz()
    QuitarEspera()
Return Nil

STATIC FUNCTION RemDatNueLin(LLAMADA)
      ***{'Linea','Concepto','Cuenta','Nombre','Banco','Talon','Fecha Vto','Importe','Eliminar','Empresa','Registro'}
    aREMESA:={   0   ,   ''     ,   0    ,   ''   ,  ''   ,  ''   ,  DATE()   ,    0    ,   .F.    ,    0    ,     0    }
    WinRemDat.GR_Remesa.AddItem(aREMESA)
    WinRemDat.GR_Remesa.Value:={WinRemDat.GR_Remesa.ItemCount,2}
    WinRemDat.GR_Remesa.Refresh
Return Nil

STATIC FUNCTION RemDatEliLin()
    local aRow := WinRemDat.GR_Remesa.Value
    local nRow := aRow[1]
    IF WinRemDat.GR_Remesa.Cell(nRow,9)=.F.
        IF MSGYESNO("¿Desea eliminar el registro activo?")=.T.
            WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value-WinRemDat.GR_Remesa.Cell(nRow,8)
            WinRemDat.GR_Remesa.Cell(nRow,9):=.T.
        ENDIF
    ELSE
        IF MSGYESNO("¿Desea recuperar el registro activo?")=.T.
            WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value+WinRemDat.GR_Remesa.Cell(nRow,8)
            WinRemDat.GR_Remesa.Cell(nRow,9):=.F.
        ENDIF
    ENDIF
    WinRemDat.GR_Remesa.Refresh

Return Nil

STATIC FUNCTION RemDat_AltaVto1()
    DEFINE WINDOW WinRemDatVto ;
        AT 10,10 ;
        WIDTH 350 HEIGHT 190 ;
        TITLE 'Alta por vencimiento' ;
        MODAL      ;
        NOSIZE

        @ 15,10 LABEL L_Vto1 VALUE 'Desde vencimiento' AUTOSIZE TRANSPARENT
        @ 10,120 DATEPICKER D_Vto1 WIDTH 100 VALUE DATE()
        @ 45,10 LABEL L_Vto2 VALUE 'Hasta vencimiento' AUTOSIZE TRANSPARENT
        @ 40,120 DATEPICKER D_Vto2 WIDTH 100 VALUE DATE()

        aFpago:={"  Todas las formas de pago"}
        FOR N=0 TO 8
            AADD(aFpago, LTRIM(STR(N))+"-"+VENCI_NC(N,30) )
        NEXT
        @ 75,10 LABEL L_Fpago VALUE 'Forma de Pago' AUTOSIZE TRANSPARENT
        @ 70,120 COMBOBOX C_Fpago WIDTH 200 ITEMS aFpago VALUE 5

        @100,10 LABEL L_Progres VALUE 'Progreso' AUTOSIZE TRANSPARENT
        @100,120 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 200 HEIGHT 20 SMOOTH

        @130,010 BUTTONEX Bt_Alta CAPTION 'Alta' ICON icobus('nuevo') ;
            ACTION RemDat_AltaVto2() ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @130,110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION WinRemDatVto.release

    END WINDOW
    VentanaCentrar("WinRemDatVto","Ventana1")
    ACTIVATE WINDOW WinRemDatVto
Return Nil



STATIC FUNCTION RemDat_AltaVto2()
    aAltaVto:={}
    LREM2:=0
    IF WinRemDat.GR_Remesa.ItemCount>0
        LREM2:=WinRemDat.GR_Remesa.Cell(WinRemDat.GR_Remesa.ItemCount,1)
    ENDIF
*** INCLUIR EJERCICIOS ANTORIORES ***
    FOR ANY2=EJERANY-1 TO EJERANY+1
        WinRemDatVto.L_Progres.Value:=STR(ANY2,4)
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        IF FILE(RUTA2+"\FAC92.DBF") .AND. ;
            FILE(RUTA2+"\FAC92.CDX")
            AbrirDBF("FAC92",,,,RUTA2)
            WinRemDatVto.L_Progres.Value:=LTRIM(STR(ANY2))
            WinRemDatVto.P_Progres.RangeMax:=LASTREC()
            WinRemDatVto.P_Progres.Value:=0
            GO TOP
            DO WHILE .NOT. EOF()
                WinRemDatVto.P_Progres.Value:=WinRemDatVto.P_Progres.Value+1
                DO EVENTS
                IF WinRemDatVto.C_Fpago.value<>0 .AND. VAL(LEFT(LTRIM(STR(FPAGO)),1))<>WinRemDatVto.C_Fpago.value-2
                    SKIP
                    LOOP
                ENDIF
                SFAC2:=SERFAC
                NFAC2:=NFAC
                NSFAC2:=LTRIM(STR(NFAC2))+"-"+SFAC2
                FFAC2:=FFAC
                CODCTA2:=CODCTA
                NOMCLI2:=RTRIM(CLIENTE)
                TFAC2:=TFAC
                PEND2:=PEND
                ACTA2:=ACTA
                FPAGO2:=FPAGO
                DPAGO2:=DPAGO
                IMPVTO2:=0
                EMPSUI2:=NEMP

                NVTOS:=VAL(SUBSTR(LTRIM(STR(FPAGO)),2,1))
                DIAVTO1=VAL(LEFT(DPAGO,2))
                DIAVTO2=VAL(RIGHT(DPAGO,2))

                AbrirDBF("CUENTAS",,,,RUTA2,1)
                SEEK CODCTA2
                BANCTA2:=BANCTA
                POBCTA2:=CODPOS+"-"+RTRIM(POBCTA)
                SELEC FAC92

                IF NVTOS=0 .AND. FFAC>=WinRemDatVto.D_Vto1.value .AND. FFAC<=WinRemDatVto.D_Vto2.value
                    AADD(aAltaVto,{++LREM2,NSFAC2,CODCTA2,NOMCLI2,BANCTA2,POBCTA2,FFAC2,TFAC2-ACTA2,.F.,EMPSUI2,NFAC2})
                ENDIF
                IMPVTO2:=IMPVTO2+ACTA
                IF ACTA<>0 .AND. FFAC>=WinRemDatVto.D_Vto1.value .AND. FFAC<=WinRemDatVto.D_Vto2.value
                    AADD(aAltaVto,{++LREM2,NSFAC2,CODCTA2,NOMCLI2,BANCTA2,POBCTA2,FFAC2,ACTA2,.F.,EMPSUI2,NFAC2})
                ENDIF
                FOR NVTO=1 TO NVTOS
                    FVTO2:=VENCI(2,FPAGO,FFAC,,DIAVTO1,DIAVTO2,NVTO)
                    TVTO2:=VENCI(3,FPAGO,,TFAC-ACTA,,,NVTO)
                    IMPVTO2:=IMPVTO2+TVTO2
                    IF FVTO2>=WinRemDatVto.D_Vto1.value  .AND. FVTO2<=WinRemDatVto.D_Vto2.value
                        AADD(aAltaVto,{++LREM2,NSFAC2,CODCTA2,NOMCLI2,BANCTA2,POBCTA2,FVTO2,TVTO2,.F.,EMPSUI2,NFAC2})
                    ENDIF
                NEXT
                SKIP
            ENDDO
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***

    ASORT(aAltaVto,,, { |x, y| x[3] < y[3] })
    FOR N=1 TO LEN(aAltaVto)
        WinRemDat.GR_Remesa.AddItem(aAltaVto[N])
        WinRemDat.T_Total.Value:=WinRemDat.T_Total.Value+aAltaVto[N,8]
    NEXT
    WinRemDat.GR_Remesa.Value:={WinRemDat.GR_Remesa.ItemCount,2}
    WinRemDat.GR_Remesa.Refresh

    WinRemDatVto.release

    Return Nil