#include "minigui.ch"

Function AsiDat()
    IF IsWindowDefined(WinAsiDat1)=.T.
        DECLARE WINDOW WinAsiDat1
        WinAsiDat1.restore
        WinAsiDat1.Show
        WinAsiDat1.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinAsiDat1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Asientos" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("ARENA") ;
        ON RELEASE CloseTables()

        @015,10 LABEL L_Asiento VALUE 'Asiento' AUTOSIZE TRANSPARENT
        @010,75 TEXTBOX T_Asiento WIDTH 75 NUMERIC RIGHTALIGN ON LOSTFOCUS AsiDatActualiz("TEXTBOX")
        @010,155 BUTTONEX Bt_BuscarAsi ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION (br_asi1(WinAsiDat1.T_Asiento.Value,"WinAsiDat1","T_Asiento", ;
            WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value)) , ;
            AsiDatActualiz("BOTON") ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @015,310 LABEL L_FecApu VALUE 'Fecha' AUTOSIZE TRANSPARENT
        @010,375 DATEPICKER D_FecApu WIDTH 100 VALUE DATE()

        @010,550 RADIOGROUP R_Cierre ;
            OPTIONS { 'Apuntes' , 'Cierre' } ;
            VALUE 1 WIDTH 75 ;
            ON CHANGE AsiDatActualiz("CIERRE") ;
            HORIZONTAL NOTABSTOP

        /*
        @040,010 GRID GR_Apuntes ;
            HEIGHT 400 ;
            WIDTH 750 ;
            HEADERS {'Linea','Cuenta','Descripcion','Apunte','Debe','Haber'} ;
            WIDTHS {50,100,150,200,100,100 } ;
            ITEMS {} ;
            EDIT INPLACE {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
            COLUMNVALID { ,{ || AsiDat_BusCueCod1()},,,{ || AsiDat_Importe()},{ || AsiDat_Importe()} } ;
            COLUMNWHEN  { { || .F.},,{ || .F.},,, } ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT, ;
            BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            ON CHANGE AsiDat_VerCue1() ;
            CELLNAVIGATION
 */

        @040,010 GRID GR_Apuntes ;
            HEIGHT 400 ;
            WIDTH 750 ;
            HEADERS {'Linea','Cuenta','Descripcion','Apunte','Debe','Haber'} ;
            WIDTHS {50,100,150,200,100,100 } ;
            ITEMS {} ;
            INPLACE ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
            VALID { ,{ || AsiDat_BusCueCod1()},,,{ || AsiDat_Importe()},{ || AsiDat_Importe()} } ;
            COLUMNWHEN  { { || .F.},,{ || .F.},,, } ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT, ;
            BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            ON CHANGE AsiDat_VerCue1() ;
            NAVIGATEBYCELL

**      Menu_Grid("WinAsiDat1","GR_Apuntes","MENU",,{"COPREGPP","COPTABPP"})

        DEFINE CONTEXT MENU CONTROL GR_Apuntes OF WinAsiDat1
        ITEM "Añadir linea"            ACTION AsiDatNueLin()
        ITEM "Eliminar linea"          ACTION AsiDatEliLin()
        SEPARATOR
        ITEM "Buscar cuenta"           ACTION AsiDat_BusCueBrow1()
        SEPARATOR
        ITEM "Copiar celda"            ACTION Menu_Grid("WinAsiDat1","GR_Apuntes","CopiarCelda",2)
        ITEM "Copiar registro"         ACTION Menu_Grid("WinAsiDat1","GR_Apuntes","CopiarRegistro")
        ITEM "Copiar tabla"            ACTION Menu_Grid("WinAsiDat1","GR_Apuntes","CopiarTabla")
    END MENU

    @450,010 BUTTONEX Bt_NueLin CAPTION 'Linea' ICON icobus('nuevo') ;
        ACTION AsiDatNueLin() ;
        WIDTH 90 HEIGHT 25 NOTABSTOP

    @450,110 BUTTONEX Bt_EliLin CAPTION 'Linea' ICON icobus('eliminar') ;
        ACTION AsiDatEliLin() ;
        WIDTH 90 HEIGHT 25 NOTABSTOP

    @480,010 TEXTBOX T_CodCta WIDTH 100 MAXLENGTH 8 READONLY
    @480,120 TEXTBOX T_NomCta WIDTH 200 VALUE '' READONLY
    @480,330 TEXTBOX T_ImpCta WIDTH 110 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN
    @480,450 BUTTON Bt_Mayor CAPTION 'Mayor' ;
        ACTION br_suizoext(RUTAPROGRAMA,VAL(WinAsiDat1.T_CodCta.Value), DMA1,DMA2,STR(NUMEMP)) ;
        WIDTH 90 HEIGHT 25 NOTABSTOP


    @455,510 LABEL L_Cuadre VALUE 'Cuadre asiento' AUTOSIZE TRANSPARENT
    @450,610 TEXTBOX T_Cuadre WIDTH 110 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


    @540,010 BUTTONEX Bt_Nuevo CAPTION 'Nuevo' ICON icobus('nuevo') ;
        ACTION AsiDatNuevo("NUEVO") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') ;
        ACTION AsiDatModif(.T.) ;
        WIDTH 95 HEIGHT 25

    @540,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') ;
        ACTION AsiDatGuardar() ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') ;
        ACTION AsiDatActualiz("CANCELAR") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,410 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
        ACTION AsiDatEliminar("SIPREGUNTA") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @540,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
        ACTION WinAsiDat1.Release ;
        WIDTH 95 HEIGHT 25 NOTABSTOP


    AsiDatModif(.F.)

END WINDOW

VentanaCentrar("WinAsiDat1","Ventana1","Alinear")
CENTER WINDOW WinAsiDat1
ACTIVATE WINDOW WinAsiDat1

Return Nil

STATIC FUNCTION AsiDat_BusCueCod1()
    local CODCUE2 := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex
    local CODCUE3 := WinAsiDat1.GR_Apuntes.Cell(nRow,2)
    AbrirDBF("CUENTAS",,,,,1)
    SEEK VAL(PCODCTA3(CODCUE2))
    IF .NOT. EOF()
        This.CellValue:=STR(CODCTA)
        WinAsiDat1.GR_Apuntes.Cell(nRow,3):=RTRIM(NOMCTA)
        AsiDat_VerCue1()
    ELSE
        This.CellValue:=CODCUE3
    ENDIF
RETURN Nil


STATIC FUNCTION AsiDat_BusCueBrow1()
    local nCol := 2
    local aRow := WinAsiDat1.GR_Apuntes.Value
    local nRow := aRow[1]
    local CODCUE2 := WinAsiDat1.GR_Apuntes.Cell(nRow,2)
    IF nRow>1 .AND. VAL(CODCUE2)=0
        CODCUE2:=WinAsiDat1.GR_Apuntes.Cell(nRow-1,2)
    ENDIF
    WinAsiDat1.T_CodCta.Value:=""
    br_cue1(CODCUE2,"WinAsiDat1","T_CodCta","T_NomCta")
    IF LEN(LTRIM(WinAsiDat1.T_CodCta.Value))=0
        RETURN
    ENDIF
    WinAsiDat1.GR_Apuntes.Cell(nRow,2):=WinAsiDat1.T_CodCta.Value
    WinAsiDat1.GR_Apuntes.Cell(nRow,3):=WinAsiDat1.T_NomCta.Value
    AsiDat_VerCue1()
RETURN Nil


STATIC FUNCTION AsiDat_VerCue1()
    local aRow := WinAsiDat1.GR_Apuntes.Value
    local nRow := aRow[1]
    IF WinAsiDat1.GR_Apuntes.ItemCount=0
        WinAsiDat1.T_CodCta.Value:=""
        WinAsiDat1.T_NomCta.Value:=""
        WinAsiDat1.T_ImpCta.Value:=0
    ELSE
        AbrirDBF("CUENTAS",,,,,1)
        SEEK VAL(WinAsiDat1.GR_Apuntes.Cell(nRow,2))
        IF .NOT. EOF()
            WinAsiDat1.T_CodCta.Value:=LTRIM(STR(CODCTA))
            WinAsiDat1.T_NomCta.Value:=NOMCTA
            WinAsiDat1.T_ImpCta.Value:=SALDO
        ELSE
            WinAsiDat1.T_CodCta.Value:=""
            WinAsiDat1.T_NomCta.Value:=""
            WinAsiDat1.T_ImpCta.Value:=0
        ENDIF
    ENDIF
Return Nil

STATIC FUNCTION AsiDat_Importe()
    local nVal := This.CellValue
    local nCol := This.CellColIndex
    local nRow := This.CellRowIndex
    IF nCol=5
        WinAsiDat1.T_Cuadre.Value:=WinAsiDat1.T_Cuadre.Value-WinAsiDat1.GR_Apuntes.Cell(nRow,nCol)+nVal
        IF nVal<>0 .AND. WinAsiDat1.GR_Apuntes.Cell(nRow,6)<>0
            WinAsiDat1.T_Cuadre.Value:=WinAsiDat1.T_Cuadre.Value+WinAsiDat1.GR_Apuntes.Cell(nRow,6)
            WinAsiDat1.GR_Apuntes.Cell(nRow,6):=0
        ENDIF
    ENDIF
    IF nCol=6
        WinAsiDat1.T_Cuadre.Value:=WinAsiDat1.T_Cuadre.Value+WinAsiDat1.GR_Apuntes.Cell(nRow,nCol)-nVal
        IF WinAsiDat1.GR_Apuntes.Cell(nRow,5)<>0 .AND. nVal<>0
            WinAsiDat1.T_Cuadre.Value:=WinAsiDat1.T_Cuadre.Value-WinAsiDat1.GR_Apuntes.Cell(nRow,5)
            WinAsiDat1.GR_Apuntes.Cell(nRow,5):=0
        ENDIF
    ENDIF
    AsiDatCuadre("COLOR")
Return Nil



STATIC FUNCTION AsiDatNuevo(LLAMADA)
    WinAsiDat1.T_Asiento.Value:=0
    AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
    GO BOTT
    IF FECHA>=DMA1 .AND. FECHA<=DMA2
        WinAsiDat1.D_FecApu.Value:=FECHA
    ELSE
        WinAsiDat1.D_FecApu.Value:=IF(DATE()<DMA2,DATE(),DMA2)
    ENDIF
    WinAsiDat1.GR_Apuntes.DeleteAllItems
    WinAsiDat1.GR_Apuntes.Refresh
    AsiDatCuadre()
    AsiDatModif(.T.)
    IF LLAMADA="NUEVO"
        WinAsiDat1.D_FecApu.SetFocus
    ENDIF
Return Nil

STATIC FUNCTION AsiDatModif(Modif)
    SoloVer:=IF(Modif=.T.,.F.,.T.)
    //*** WinAsiDat1.D_FecApu.ReadOnly:=SoloVer
    WinAsiDat1.GR_Apuntes.Enabled:=Modif
    IF Modif = .T.
        WinAsiDat1.T_Asiento.Enabled  := .F.
        WinAsiDat1.Bt_Nuevo.Enabled   := .F.
        WinAsiDat1.Bt_Modif.Enabled   := .F.
        WinAsiDat1.Bt_Guardar.Enabled := .T.
        WinAsiDat1.Bt_Cancelar.Enabled:= .T.
        WinAsiDat1.Bt_Eliminar.Enabled:= .T.
        WinAsiDat1.Bt_Salir.Enabled   := .F.
        WinAsiDat1.Bt_NueLin.Enabled  := .T.
        WinAsiDat1.Bt_EliLin.Enabled  := .T.
        ON KEY ADD OF WinAsiDat1       ACTION AsiDatNueLin()
        WinAsiDat1.D_FecApu.SetFocus
    ELSE
        WinAsiDat1.T_Asiento.Enabled  := .T.
        WinAsiDat1.Bt_Nuevo.Enabled   := .T.
        WinAsiDat1.Bt_Modif.Enabled   := .T.
        WinAsiDat1.Bt_Guardar.Enabled := .F.
        WinAsiDat1.Bt_Cancelar.Enabled:= .F.
        WinAsiDat1.Bt_Eliminar.Enabled:= .F.
        WinAsiDat1.Bt_Salir.Enabled   := .T.
        WinAsiDat1.Bt_NueLin.Enabled  := .F.
        WinAsiDat1.Bt_EliLin.Enabled  := .F.
**   RELEASE KEY ADD OF WinAsiDat1
        ON KEY ADD OF WinAsiDat1       ACTION NIL
        WinAsiDat1.GR_Apuntes.SetFocus
    ENDIF
Return Nil

STATIC FUNCTION AsiDatGuardar()
    FOR N=1 TO WinAsiDat1.GR_Apuntes.ItemCount
        IF VAL(WinAsiDat1.GR_Apuntes.Cell(N,2))=0
            MsgStop("No se puede grabar el asiento"+HB_OsNewLine()+"Existen apuntes sin cuenta","error")
            RETURN
        ENDIF
    NEXT
    IF WinAsiDat1.T_Cuadre.Value<>0
        IF MSGYESNO("El asiento esta descuadrado"+HB_OsNewLine()+"¿Desea grabarlo?")=.F.
            RETURN
        ENDIF
    ENDIF
    IF WinAsiDat1.D_FecApu.Value<DMA1 .OR. WinAsiDat1.D_FecApu.Value>DMA2
        IF MSGYESNO("La fecha del asiento esta fuera del ejercicio"+HB_OsNewLine()+"¿Desea grabarlo?")=.F.
            RETURN
        ENDIF
    ENDIF

    AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
    IF WinAsiDat1.T_Asiento.Value<>0
        AsiDatEliminar("NOPREGUNTA")
    ELSE
        GO BOTT
        IF .NOT. EOF()
            WinAsiDat1.T_Asiento.Value:=NASI+1
        ELSE
            IF WinAsiDat1.R_Cierre.Value=1
                WinAsiDat1.T_Asiento.Value:=2
            ELSE
                WinAsiDat1.T_Asiento.Value:=999001
            ENDIF
        ENDIF
    ENDIF

    PonerEspera("Contabilizando asiento....")

    FOR N=1 TO WinAsiDat1.GR_Apuntes.ItemCount
        AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH WinAsiDat1.T_Asiento.Value
            REPLACE APU  WITH N
            REPLACE FECHA WITH WinAsiDat1.D_FecApu.Value
            REPLACE CODCTA WITH VAL(WinAsiDat1.GR_Apuntes.Cell(N,2))
            REPLACE NOMAPU WITH WinAsiDat1.GR_Apuntes.Cell(N,4)
            REPLACE DEBE  WITH WinAsiDat1.GR_Apuntes.Cell(N,5)
            REPLACE HABER WITH WinAsiDat1.GR_Apuntes.Cell(N,6)
            REPLACE NEMP WITH NUMEMP
            DBCOMMIT()
            DBUNLOCK()
            SALDO2:=DEBE-HABER

            AbrirDBF("CUENTAS",,,,,1)
            SEEK VAL(WinAsiDat1.GR_Apuntes.Cell(N,2))
            IF .NOT. EOF()
                IF RLOCK()
                    REPLACE SALDO WITH SALDO+SALDO2
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
            ENDIF
        ENDIF
    NEXT

    QuitarEspera()

    MsgInfo('Los datos han sido guardados','Datos guardados')
    AsiDatModif(.F.)

Return Nil

STATIC FUNCTION AsiDatActualiz(LLAMADA)
    DO CASE
    CASE WinAsiDat1.T_Asiento.Value=0 .AND. LLAMADA="TEXTBOX"
        RETURN
    CASE WinAsiDat1.T_Asiento.Value=0 .AND. LLAMADA="BOTON"
        RETURN
    CASE WinAsiDat1.T_Asiento.Value=0 .AND. LLAMADA="CIERRE"
        RETURN
    CASE WinAsiDat1.T_Asiento.Value=0 .AND. LLAMADA<>"CANCELAR"
        AsiDatNuevo()
        RETURN
    ENDCASE

    AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
    SEEK WinAsiDat1.T_Asiento.Value
    IF .NOT. EOF()
        WinAsiDat1.D_FecApu.Value:=FECHA
    ELSE
        IF DATE()<DMA1
            WinAsiDat1.D_FecApu.Value:=DATE()
        ELSE
            WinAsiDat1.D_FecApu.Value:=DMA1
        ENDIF
    ENDIF
    WinAsiDat1.GR_Apuntes.DeleteAllItems
    AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
    SEEK WinAsiDat1.T_Asiento.Value
    DO WHILE NASI=WinAsiDat1.T_Asiento.Value .AND. .NOT. EOF()
*     'Linea','Cuenta','Descripcion','Apunte','Debe','Haber'
        aAPU:={}
        AADD(aAPU,WinAsiDat1.GR_Apuntes.ItemCount+1)
        AADD(aAPU,STR(CODCTA))
        AbrirDBF("CUENTAS",,,,,1)
        SEEK APUNTES->CODCTA
        AADD(aAPU,RTRIM(NOMCTA))
        AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
        AADD(aAPU,RTRIM(NOMAPU))
        AADD(aAPU,DEBE)
        AADD(aAPU,HABER)
        WinAsiDat1.GR_Apuntes.AddItem(aAPU)
        SKIP
    ENDDO
    WinAsiDat1.GR_Apuntes.Value:={WinAsiDat1.GR_Apuntes.ItemCount,2}
    WinAsiDat1.GR_Apuntes.Refresh
    AsiDatCuadre()
    AsiDatModif(.F.)
Return Nil

STATIC FUNCTION AsiDatEliminar(LLAMADA)
    IF LLAMADA="SIPREGUNTA"
        IF MSGYESNO("¿Desea eliminar el asiento activo?")=.F.
            RETURN
        ENDIF
    ENDIF

    PonerEspera("Procesando los datos....")
    AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
    SEEK WinAsiDat1.T_Asiento.Value
    DO WHILE NASI=WinAsiDat1.T_Asiento.Value .AND. .NOT. EOF()
        DO EVENTS
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
        CODCTA2:=CODCTA
        SALDO2:=DEBE-HABER
        AbrirDBF("CUENTAS",,,,,1)
        SEEK CODCTA2
        IF .NOT. EOF()
            IF RLOCK()
                REPLACE SALDO WITH SALDO-SALDO2
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
        ENDIF
        AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
        SKIP
    ENDDO
    IF LLAMADA="SIPREGUNTA"
        AsiDatActualiz()
    ENDIF
    QuitarEspera()
Return Nil

STATIC FUNCTION AsiDatNueLin(LLAMADA)
          **{'Linea','Cuenta','Descripcion','Apunte','Debe','Haber'}
    aAPUNTE:={ 0     , ''     ,''           ,''      , 0    , 0     }

    aAPUNTE[1]:=WinAsiDat1.GR_Apuntes.ItemCount+1
    IF WinAsiDat1.T_Asiento.Value=1
        aAPUNTE[4]:="Asiento de apertura"
    ELSE
        IF WinAsiDat1.GR_Apuntes.ItemCount>0
            aAPUNTE[4]:=WinAsiDat1.GR_Apuntes.Cell(WinAsiDat1.GR_Apuntes.ItemCount,4)
        ELSE
            AbrirDBF(WinAsiDat1.R_Cierre.Caption(WinAsiDat1.R_Cierre.Value),,,,,1)
            GO BOTT
            IF .NOT. EOF()
                aAPUNTE[4]:=RTRIM(NOMAPU)
            ENDIF
        ENDIF
    ENDIF
    IF WinAsiDat1.T_Cuadre.Value<0
        aAPUNTE[5]:=WinAsiDat1.T_Cuadre.Value*-1
        aAPUNTE[6]:=0
    ELSE
        aAPUNTE[5]:=0
        aAPUNTE[6]:=WinAsiDat1.T_Cuadre.Value
    ENDIF
    WinAsiDat1.GR_Apuntes.AddItem(aAPUNTE)
    WinAsiDat1.GR_Apuntes.Value:={WinAsiDat1.GR_Apuntes.ItemCount,2}
    WinAsiDat1.GR_Apuntes.Refresh
    AsiDatCuadre()
Return Nil

STATIC FUNCTION AsiDatEliLin()
    local aRow := WinAsiDat1.GR_Apuntes.Value
    local nRow := aRow[1]
    IF MSGYESNO("¿Desea eliminar el apunte activo?")=.T.
        WinAsiDat1.GR_Apuntes.DeleteItem(nRow)
    ENDIF
   ***MODIFICAR Nº DE LINEA***
    FOR N=1 TO WinAsiDat1.GR_Apuntes.ItemCount
        WinAsiDat1.GR_Apuntes.Cell(N,1):=N
    NEXT
   ***FIN MODIFICAR Nº DE LINEA***
    IF nRow<WinAsiDat1.GR_Apuntes.ItemCount
        WinAsiDat1.GR_Apuntes.Value:={nRow,aRow[2]}
    ELSE
        WinAsiDat1.GR_Apuntes.Value:={WinAsiDat1.GR_Apuntes.ItemCount,aRow[2]}
    ENDIF
    WinAsiDat1.GR_Apuntes.Refresh
    AsiDatCuadre()
    AsiDat_VerCue1()
Return Nil

STATIC FUNCTION AsiDatCuadre(LLAMADA)
    LLAMADA:=IF(LLAMADA=NIL,"TODO",LLAMADA)
    IF LLAMADA="TODO"
        SALDO2:=0
        FOR N=1 TO WinAsiDat1.GR_Apuntes.ItemCount
            SALDO2:=SALDO2+WinAsiDat1.GR_Apuntes.Cell(N,5)-WinAsiDat1.GR_Apuntes.Cell(N,6)
        NEXT
        WinAsiDat1.T_Cuadre.Value:=SALDO2
    ENDIF
    IF LLAMADA="COLOR" .OR. LLAMADA="TODO"
        IF WinAsiDat1.T_Cuadre.Value=0
            WinAsiDat1.T_Cuadre.BackColor:=MICOLOR("VERDECLARO")
        ELSE
            WinAsiDat1.T_Cuadre.BackColor:=MICOLOR("ROJOCLARO")
        ENDIF
    ENDIF
Return Nil
