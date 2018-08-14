#include "minigui.ch"

Function br_asi1(CODIGO1,VENTANA1,CONTROL1,FICHERO,RUTAFIC)
    LOCAL Color_asi1 := { || IIF( NASI%2<>0 , RGB(200,255,200), RGB(200,200,255)) }
    FICHERO:=IF(FICHERO=NIL,"APUNTES",FICHERO)
    RUTAFIC:=IF(RUTAFIC=NIL,RUTAEMP,RUTAFIC)
    CODBR1:=1

    IF AT("GRUPOSP",UPPER(RUTAFIC))<>0
        FICHERO:="DIARIO"
        Color_asi1 := { || IIF( ASIEN%2<>0 , RGB(200,255,200), RGB(200,200,255)) }
        aCAMPOS:={ 'ASIEN','FECHA','VAL(SUBCTA)','CONCEPTO','MIL(EURODEBE,12,2)','MIL(EUROHABER,12,2)' }
        nINDICE:=3
    ELSE
        aCAMPOS:={ 'NASI','FECHA','CODCTA','NOMAPU','MIL(DEBE,12,2)','MIL(HABER,12,2)' }
        nINDICE:=1
    ENDIF

    AbrirDBF(FICHERO,,,,RUTAFIC,nINDICE)

    SEEK CODIGO1
    IF .NOT. EOF()
        CODBR1:=RECNO()
    ELSE
        GO BOTT
        IF .NOT. EOF()
            CODBR1:=RECNO()
        ENDIF
    ENDIF

    DEFINE WINDOW WinBRasi1 ;
        AT 0,0     ;
        WIDTH 735  ;
        HEIGHT 500 ;
        TITLE "Consulta de asientos" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("AMARILLOHUEVO")

        @ 10,10 BROWSE br_asi1 ;
            HEIGHT 380 ;
            WIDTH 685 ;
            HEADERS {'Asiento','Fecha','Cuenta','Apunte','Debe','Haber'} ;
            WIDTHS {50,85,100,200,100,100 } ;
            WORKAREA &FICHERO ;
            FIELDS aCAMPOS ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT, ;
            BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            DYNAMICBACKCOLOR { Color_asi1,Color_asi1,Color_asi1,Color_asi1,Color_asi1,Color_asi1 } ;
            VALUE CODBR1 ;
            ON DBLCLICK Terminar_BusAsi1(VENTANA1,CONTROL1,FICHERO,RUTAFIC)

        @405,010 LABEL L_Buscar VALUE 'Buscar' AUTOSIZE TRANSPARENT
        @400,110 TEXTBOX T_BuscarNom WIDTH 250 VALUE '' MAXLENGTH 30 ;
            ON CHANGE { || Bus_SuiCodCta(FICHERO,RUTAFIC) }
/*
        @400,400 BUTTONEX Bt_Duplicar CAPTION 'Duplicar' ICON icobus('guardar') WIDTH 90 HEIGHT 25 ;
            ACTION Menu_WinBRasi1("Duplicarasiento",FICHERO,RUTAFIC) NOTABSTOP
*/

        DEFINE CONTEXT MENU CONTROL br_asi1 OF WinBRasi1
        ITEM "Copiar descripcion"      ACTION Menu_WinBRasi1("CopiarDescripcion",FICHERO,RUTAFIC)
        ITEM "Duplicar asiento"        ACTION Menu_WinBRasi1("Duplicarasiento",FICHERO,RUTAFIC)
    END MENU

    WinBRasi1.T_BuscarNom.SetFocus

END WINDOW
VentanaCentrar("WinBRasi1","Ventana1")
CENTER WINDOW WinBRasi1
ACTIVATE WINDOW WinBRasi1

Return Nil

STATIC FUNCTION Menu_WinBRasi1(LLAMADA,FICHERO,RUTAFIC)
    DO CASE
    CASE LLAMADA="CopiarDescripcion"
        AbrirDBF(FICHERO,,,,RUTAFIC)
        GO WinBRasi1.br_asi1.Value
        //*** CopyToClipboard(RTRIM(aCAMPOS[4]))
        SetClipboardText(RTRIM(aCAMPOS[4]))
        MSGBOX(RTRIM(NOMAPU),"Texto del portapapeles")
    CASE LLAMADA="Duplicarasiento"
        AbrirDBF(FICHERO,,,,RUTAFIC)
        GO WinBRasi1.br_asi1.Value
        NASIDUP1:=FIELDGET(FIELDPOS(aCAMPOS[1]))
        NASIDUP2:=Asi_Duplicar(RUTAFIC,NASIDUP1,0)
        AbrirDBF(FICHERO,,,,RUTAFIC)
        SEEK IF(NASIDUP2=0,NASIDUP1,NASIDUP2)
        WinBRasi1.br_asi1.Value:=RECNO()
        WinBRasi1.br_asi1.Refresh
    ENDCASE

Return Nil

STATIC FUNCTION Bus_SuiCodCta(FICHERO,RUTAFIC)
    AbrirDBF(FICHERO,,,,RUTAFIC,1)
    SEEK VAL(WinBRasi1.T_BuscarNom.Value)
    IF .NOT. EOF()
        WinBRasi1.br_asi1.Value := Recno()
        WinBRasi1.br_asi1.Refresh
    ENDIF
Return Nil

STATIC FUNCTION Terminar_BusAsi1(VENTANA1,CONTROL1,FICHERO,RUTAFIC)
    IF VENTANA1<>NIL
        AbrirDBF(FICHERO,,,,RUTAFIC)
        GO WinBRasi1.br_asi1.Value
        IF FICHERO="DIARIO"
            SetProperty(VENTANA1,CONTROL1,"Value",ASIEN)
        ELSE
            SetProperty(VENTANA1,CONTROL1,"Value",NASI)
        ENDIF
    ENDIF
    WinBRasi1.Release

Return Nil
