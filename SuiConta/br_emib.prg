#include "minigui.ch"
** RUTAEB:=RUTAPROGRAMA

Function br_emibus( RUTAEB,NEMP2,SFAC2,NFAC2,NOMCLI2 )
    LOCAL Color_fac1 := { || IIF( PEND<>0 , RGB(255,200,200), RGB(255,255,255)) }
    Local CODFACBR:=1        // MigSoft
    DEFAULT NOMCLI2 TO ""

    NEMP2:=IF(NEMP2=0,NUMEMP,NEMP2)

    aNFAC:={0,"",""}

    IREMPRESA := 1
    aRUTAEMP1 := {}
    aRUTAEMP2 := {}
    aRUTAEMP3 := {}
    aRUTAEMP  := aRUTAEMP(RUTAEB,"SUICONTA")

    IF LEN(aRUTAEMP)>=1
        FOR N=1 TO LEN(aRUTAEMP)
            AADD(aRUTAEMP1,aRUTAEMP[N,1])
            AADD(aRUTAEMP2,aRUTAEMP[N,2])
            AADD(aRUTAEMP3,LTRIM(aRUTAEMP[N,4]))
            IF VAL(aRUTAEMP[N,4])=NEMP2
                IREMPRESA=N
            ENDIF
        NEXT
    ENDIF

    AbrirDBF("fac92",,,,RUTAEB+"\"+aRUTAEMP2[IREMPRESA])

    IF NFAC2<>0
        DBSETORDER(1)
        SEEK SERIE(SFAC2,NFAC2)
        CODFACBR:=RECNO()
    ELSE
        IF LEN(RTRIM(NOMCLI2))<>0
            DBSETORDER(2)
            SET SOFTSEEK ON
            SEEK NOMCLI2
            SET SOFTSEEK OFF
            CODRECBR:=RECNO()
            CODFACBR:=CODRECBR    // MigSoft
        ELSE
            CODFACBR:=1
        ENDIF
    ENDIF

    DBSETORDER(2)

    DEFINE WINDOW Winfacbus1 ;
        AT 50,0     ;
        WIDTH 580  ;
        HEIGHT 430 ;
        TITLE "Facturas emitidas" ;
        MODAL ;
        NOSIZE BACKCOLOR MiColor("NARANJA")

        @0,0 TEXTBOX RutaEB WIDTH 250 VALUE RUTAEB INVISIBLE

        @ 10,10 BROWSE BR_fac1 ;
            HEIGHT 300 ;
            WIDTH 550 ;
            HEADERS {'factura','Fecha','Nombre','Importe','Pendiente','Asiento'} ;
            WIDTHS { 50,70,170,80,80,75 } ;
            WORKAREA fac92 ;
            FIELDS {'LTRIM(STR(fac92->Nfac))+fac92->SERFAC','DIA(fac92->FFAC,8)','fac92->CLIENTE', ;
            'LTRIM(MIL(fac92->Tfac,12,2))','LTRIM(MIL(fac92->Pend,12,2))','fac92->ASIENTO'} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            DYNAMICBACKCOLOR { Color_fac1,Color_fac1,Color_fac1,Color_fac1,Color_fac1,Color_fac1 } ;
            VALUE CODFACBR ;
            ON DBLCLICK br_emibusFin() ;
            ON HEADCLICK { {|| AbrirDBF("fac92",,,,RUTAEB+"\"+Winfacbus1.C_RutaEmp.DisplayValue),DBSETORDER(1), Winfacbus1.BR_fac1.Refresh} , , ;
            {|| AbrirDBF("fac92",,,,RUTAEB+"\"+Winfacbus1.C_RutaEmp.DisplayValue),DBSETORDER(2), Winfacbus1.BR_fac1.Refresh} }

        DEFINE CONTEXT MENU CONTROL BR_fac1 OF Winfacbus1
        ITEM "Copiar registro"         ACTION Winfacbus1_CopReg("REG1")
        ITEM "Copiar registros cliente" ACTION Winfacbus1_CopReg("REGCLI")
    END MENU

    @330,010 BUTTONEX Bt_Localizar CAPTION 'Localizar' ICON icobus('buscar') ;
        ACTION br_emibusLoc() ;
        WIDTH 90 HEIGHT 25 NOTABSTOP
    @330,110 TEXTBOX T_BusNom WIDTH 300 VALUE '' MAXLENGTH 30 ;
        ON CHANGE br_emibusfac("BUSCAR")

    @365,010 LABEL L_Empresa VALUE 'Empresa:' AUTOSIZE TRANSPARENT
    @360,110 COMBOBOX C_Empresa WIDTH 300 ITEMS aRUTAEMP1 VALUE IREMPRESA NOTABSTOP ;
        ON CHANGE (Winfacbus1.C_RutaEmp.Value:=Winfacbus1.C_Empresa.Value , ;
        Winfacbus1.C_CodEmp.Value :=Winfacbus1.C_Empresa.Value , br_emibusfac("RUTA") )
    @360,110 COMBOBOX C_RutaEmp WIDTH 300 ITEMS aRUTAEMP2 VALUE IREMPRESA INVISIBLE
    @360,110 COMBOBOX C_CodEmp  WIDTH  50 ITEMS aRUTAEMP3 VALUE IREMPRESA INVISIBLE

    @360,450 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
        ACTION Winfacbus1.Release ;
        WIDTH 90 HEIGHT 25 NOTABSTOP

    Winfacbus1.T_BusNom.SetFocus

    END WINDOW

    VentanaCentrar("Winfacbus1","Ventana1")
    CENTER WINDOW Winfacbus1
    ACTIVATE WINDOW Winfacbus1

RETURN(aNFAC)

STATIC FUNCTION br_emibusFin()
    aNFAC:={Winfacbus1.BR_fac1.Value ,;
        Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue ,;
        VAL(Winfacbus1.C_CodEmp.DisplayValue) }
    Winfacbus1.Release
Return Nil

STATIC FUNCTION br_emibusfac(LLAMADA)
    AbrirDBF("fac92",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
    IF LLAMADA="RUTA"
        Winfacbus1.BR_fac1.Refresh
        RETURN
    ENDIF

    SET SOFTSEEK ON
    IF INDEXORD()=1
        TBusNom2:=STRTRAN(Winfacbus1.T_BusNom.Value,"-","")
        IF ISALPHA(RIGHT(TBusNom2,1))
            SFACBUS2:=RIGHT(TBusNom2,1)
            NFACBUS2:=VAL(LEFT(TBusNom2,LEN(TBusNom2)-1))
        ELSE
            SFACBUS2:="A"
            NFACBUS2:=VAL(TBusNom2)
        ENDIF
        SEEK SERIE(SFACBUS2,NFACBUS2)
    ELSE
        IF VAL(Winfacbus1.T_BusNom.Value)<>0
            NUMIND2:=INDEXORD()
            DBSETORDER(1)
            SEEK SERIE("A",VAL(Winfacbus1.T_BusNom.Value))
            DBSETORDER(NUMIND2)
        ELSE
            SEEK UPPER(Winfacbus1.T_BusNom.Value)
        ENDIF
    ENDIF
    SET SOFTSEEK OFF
    Winfacbus1.BR_fac1.Value:=RECNO()
    Winfacbus1.BR_fac1.Refresh

Return Nil

STATIC FUNCTION br_emibusLoc()
    NOMBUSCAR:=RTRIM(Winfacbus1.T_BusNom.Value)
    IF LEN(NOMBUSCAR)=0
        RETURN Nil
    ENDIF
    AbrirDBF("fac92",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
    GO Winfacbus1.BR_fac1.Value
    DO WHILE .T.
        DO EVENTS
        SKIP
        IF EOF()
            GO TOP
        ENDIF
        IF AT(UPPER(NOMBUSCAR),UPPER(CLIENTE))<>0 .OR. ;
            AT(UPPER(NOMBUSCAR),UPPER(STR(TFAC)))<>0
            Winfacbus1.BR_fac1.Value:=recno()
            Winfacbus1.BR_fac1.Refresh
            EXIT
        ENDIF
        IF RECNO()=Winfacbus1.BR_fac1.Value
            MSGSTOP("No se ha localizado ningun registro")
            EXIT
        ENDIF
    ENDDO
Return Nil

STATIC FUNCTION Winfacbus1_CopReg(LLAMADA)
    TEXTO2:=""
    AbrirDBF("fac92",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
    GO Winfacbus1.BR_fac1.Value
    DO CASE
    CASE LLAMADA="REG1"
        TEXTO2:=LTRIM(STR(NFAC))+SERFAC+CHR(9)+LTRIM(STR(CODCTA))+CHR(9)+RTRIM(CLIENTE)+CHR(9)+ ;
            DIA(FFAC)+CHR(9)+STRTRAN(LTRIM(STR(TFAC)),".",",")+CHR(9)+STRTRAN(LTRIM(STR(PEND)),".",",")
    CASE LLAMADA="REGCLI"
        PonerEspera("Copiando datos al portapapeles")
        CODCTA2:=CODCTA
        DO WHILE CODCTA=CODCTA2 .AND. .NOT. BOF()
            SKIP -1
        ENDDO
        SKIP
        DO WHILE CODCTA=CODCTA2 .AND. .NOT. EOF()
            TEXTO2:=TEXTO2+LTRIM(STR(NFAC))+SERFAC+CHR(9)+LTRIM(STR(CODCTA))+CHR(9)+RTRIM(CLIENTE)+CHR(9)+ ;
                DIA(FFAC)+CHR(9)+STRTRAN(LTRIM(STR(TFAC)),".",",")+CHR(9)+STRTRAN(LTRIM(STR(PEND)),".",",")+HB_OsNewLine()
            SKIP
        ENDDO
        QuitarEspera()
    ENDCASE
    //*** CopyToClipboard(TEXTO2)
    SetClipboardText(TEXTO2)
    MsgInfo("El texto se ha copiado al portapapeles")

Return Nil
