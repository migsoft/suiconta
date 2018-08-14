#include "minigui.ch"

Function CueDat()
    IF IsWindowDefined(WinCueDat1)=.T.
        DECLARE WINDOW WinCueDat1
        WinCueDat1.restore
        WinCueDat1.Show
        WinCueDat1.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinCueDat1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Subcuentas" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("AZULCLARO") ;
        ON RELEASE CloseTables()

        @015,010 LABEL L_CodCta VALUE 'Cuenta' AUTOSIZE TRANSPARENT
        @010,100 TEXTBOX T_CodCta WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS (WinCueDat1.T_CodCta.Value:=PCODCTA3(WinCueDat1.T_CodCta.Value), CueDat_Actualiz("CODIGO") ) NOTABSTOP
        @010,210 BUTTONEX Bt_BuscarCue ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION (br_cue1(VAL(WinCueDat1.T_CodCta.Value),"WinCueDat1","T_CodCta") , CueDat_Actualiz("CODIGO") ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @010,310 BUTTON Bt_Mayor CAPTION 'Mayor' ;
            ACTION br_suizoext(RUTAPROGRAMA,VAL(WinCueDat1.T_CodCta.Value),DMA1,DMA2,STR(NUMEMP)) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @010,410 BUTTONEX Bt_CopiarDir CAPTION 'portapapeles' ICON icobus('portapapeles') ;
            ACTION CopiaPP("TEXTO","WinCueDat1") ;
            WIDTH 100 HEIGHT 25 NOTABSTOP

        @045,010 LABEL L_Grupo VALUE 'Grupo' AUTOSIZE TRANSPARENT
        @040,100 TEXTBOX T_Grupo WIDTH 300

        @075,010 LABEL L_Nombre VALUE 'Descripcion' AUTOSIZE TRANSPARENT
        @070,100 TEXTBOX T_Nombre WIDTH 300 MAXLENGTH 40
        @070,410 BUTTON Bt_NomTit CAPTION 'Titulo' WIDTH 90 HEIGHT 25 NOTABSTOP ;
            ACTION WinCueDat1.T_Nombre.Value:=NomPropio(WinCueDat1.T_Nombre.Value)

        @105,010 LABEL L_Direccion VALUE 'Direccion' AUTOSIZE TRANSPARENT
        @100,100 TEXTBOX T_Direccion WIDTH 250 MAXLENGTH 30

        @135,010 LABEL L_Poblacion VALUE 'Poblacion' AUTOSIZE TRANSPARENT
        @130,100 TEXTBOX T_CodPos WIDTH 50 MAXLENGTH 5 ON CHANGE CodPostal("WinCueDat1")
        @130,155 TEXTBOX T_Poblacion WIDTH 250 MAXLENGTH 30

        @165,010 LABEL L_Provincia VALUE 'Provincia' AUTOSIZE TRANSPARENT
        @160,100 TEXTBOX T_Provincia WIDTH 250 MAXLENGTH 30

        @195,010 LABEL L_Pais VALUE 'Pais' AUTOSIZE TRANSPARENT
        @190,100 TEXTBOX T_Pais WIDTH 250 MAXLENGTH 30

        @225,010 LABEL L_Regimen VALUE 'Regimen' AUTOSIZE TRANSPARENT
        @220,100 COMBOBOX C_Regimen WIDTH 250 ITEMS REGIVAREC("ARRAY") VALUE 1 NOTABSTOP

        @255,010 LABEL L_CifCtaMal VALUE 'Nif erroneo' AUTOSIZE FONTCOLOR MICOLOR("ROJO") TRANSPARENT INVISIBLE
        @255,010 LABEL L_CifCta VALUE 'Cif/Nif' AUTOSIZE TRANSPARENT
        @250,100 TEXTBOX T_CifCta WIDTH 125 MAXLENGTH 15 ;
            ON LOSTFOCUS ( WinCueDat1.T_CifCta.Value:=DNI_LET(WinCueDat1.T_CifCta.Value) , CueDat_VerCifBan() )

        @285,010 LABEL L_Telefono VALUE 'Telefono' AUTOSIZE TRANSPARENT
        @280,100 TEXTBOX T_Telefono WIDTH 150 MAXLENGTH 15

        @315,010 LABEL L_Fax VALUE 'Fax' AUTOSIZE TRANSPARENT
        @310,100 TEXTBOX T_Fax WIDTH 150 MAXLENGTH 15

        @345,010 LABEL L_CtaBanMal VALUE 'Cuenta erronea' AUTOSIZE FONTCOLOR MICOLOR("ROJO") TRANSPARENT INVISIBLE
        @345,010 LABEL L_CtaBan VALUE 'Cuenta banco' AUTOSIZE TRANSPARENT
        @340,100 TEXTBOX T_CtaBan WIDTH 250 MAXLENGTH 30 ;
            ON LOSTFOCUS ( WinCueDat1.T_CtaBan.Value:=CTA_BAN_SUIZO(WinCueDat1.T_CtaBan.Value,24) , ;
            CTA_BAN(WinCueDat1.T_CtaBan.Value) , CueDat_VerCifBan() )

        @375,010 LABEL L_Saldo VALUE 'Saldo cuenta' AUTOSIZE TRANSPARENT
        @370,100 TEXTBOX T_Saldo WIDTH 110 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


        @525,010 BUTTONEX Bt_Nuevo CAPTION 'Nuevo' ICON icobus('nuevo') ;
            ACTION CueDat_Nuevo("NUEVO") ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') ;
            ACTION CueDat_Modif(.t.) ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') ;
            ACTION CueDat_Guardar() ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') ;
            ACTION CueDat_Actualiz("CANCELAR") ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,410 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
            ACTION CueDat_Eliminar() ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,510 BUTTONEX Bt_DirImp CAPTION 'Direccion' ICON icobus('imprimir') ;
            ACTION CueDat_DirImp() ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        @525,610 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
            ACTION WinCueDat1.Release ;
            WIDTH 95 HEIGHT 25 NOTABSTOP

        CueDat_Modif(.F.)

    END WINDOW
    VentanaCentrar("WinCueDat1","Ventana1","Alinear")
    CENTER WINDOW WinCueDat1
    ACTIVATE WINDOW WinCueDat1

Return Nil

STATIC FUNCTION CueDat_Nuevo(LLAMADA)
    WinCueDat1.T_Nombre.Value:=""
    WinCueDat1.T_Direccion.Value:=""
    WinCueDat1.T_CodPos.Value:=""
    WinCueDat1.T_Poblacion.Value:=""
    WinCueDat1.T_Provincia.Value:=""
    WinCueDat1.T_Pais.Value:=""
    WinCueDat1.C_Regimen.Value:=""
    WinCueDat1.T_CifCta.Value:=""
    WinCueDat1.T_Telefono.Value:=""
    WinCueDat1.T_Fax.Value:=""
    WinCueDat1.T_CtaBan.Value:=""
    WinCueDat1.T_Saldo.Value:=0
    WinCueDat1.T_Grupo.Value:=""
    IF LLAMADA="NUEVO"
        WinCueDat1.T_CodCta.Value:=""
        CueDat_Modif(.T.)
        WinCueDat1.T_CodCta.SetFocus
    ENDIF
Return Nil

STATIC FUNCTION CueDat_Modif(Modif)
    SoloVer:=IF(Modif=.T.,.F.,.T.)
    IF VAL(WinCueDat1.T_CodCta.Value)=0
        WinCueDat1.T_CodCta.Enabled:=Modif
        WinCueDat1.T_CodCta.SetFocus
    ELSE
        WinCueDat1.T_CodCta.Enabled:=.F.
        WinCueDat1.T_Nombre.SetFocus
    ENDIF
    WinCueDat1.T_Nombre.ReadOnly:=SoloVer
    WinCueDat1.T_Direccion.ReadOnly:=SoloVer
    WinCueDat1.T_CodPos.ReadOnly:=SoloVer
    WinCueDat1.T_Poblacion.ReadOnly:=SoloVer
    WinCueDat1.T_Provincia.ReadOnly:=SoloVer
    WinCueDat1.T_Pais.ReadOnly:=SoloVer
    WinCueDat1.C_Regimen.Enabled:=Modif
    WinCueDat1.T_CifCta.ReadOnly:=SoloVer
    WinCueDat1.T_Telefono.ReadOnly:=SoloVer
    WinCueDat1.T_Fax.ReadOnly:=SoloVer
    WinCueDat1.T_CtaBan.ReadOnly:=SoloVer
    WinCueDat1.T_Saldo.Enabled:=.F.
    WinCueDat1.T_Grupo.Enabled:=.F.

    WinCueDat1.Bt_BuscarCue.Enabled:= IF(Modif=.T.,.F.,.T.)
    IF Modif = .T.
        WinCueDat1.Bt_Nuevo.Enabled    := .F.
        WinCueDat1.Bt_Modif.Enabled    := .F.
        WinCueDat1.Bt_Guardar.Enabled  := .T.
        WinCueDat1.Bt_Cancelar.Enabled := .T.
        WinCueDat1.Bt_Eliminar.Enabled := .F.
        WinCueDat1.Bt_DirImp.Enabled   := .F.
        WinCueDat1.Bt_Salir.Enabled    := .F.
    ELSE
        WinCueDat1.Bt_Nuevo.Enabled    := .T.
        WinCueDat1.Bt_Modif.Enabled    := .T.
        WinCueDat1.Bt_Guardar.Enabled  := .F.
        WinCueDat1.Bt_Cancelar.Enabled := .F.
        WinCueDat1.Bt_Eliminar.Enabled := .T.
        WinCueDat1.Bt_DirImp.Enabled   := .T.
        WinCueDat1.Bt_Salir.Enabled    := .T.
    ENDIF
    CueDat_VerRegimen()

Return Nil


STATIC FUNCTION CueDat_VerRegimen()
    IF VAL(WinCueDat1.T_CodCta.Value)>=40000000 .AND. ;
        VAL(WinCueDat1.T_CodCta.Value)<=44009999
        WinCueDat1.C_Regimen.Visible:=.T.
    ELSE
        WinCueDat1.C_Regimen.Visible:=.F.
    ENDIF

Return Nil


STATIC FUNCTION CueDat_Guardar()

    IF VAL(WinCueDat1.T_CodCta.Value)=0 .OR. ;
        LEN(WinCueDat1.T_CodCta.Value)<>8
        MSGBOX("El codigo de la cuentas no es valido","error")
        RETURN
    ENDIF

    AbrirDBF("CUENTAS")
    DBSETORDER(1)
    SEEK VAL(WinCueDat1.T_CodCta.Value)

    IF EOF()
        APPEND BLANK
    ENDIF

    IF RLOCK()
        REPLACE CODCTA WITH VAL(WinCueDat1.T_CodCta.Value)
        REPLACE NOMCTA WITH WinCueDat1.T_Nombre.Value
        REPLACE DIRCTA WITH WinCueDat1.T_Direccion.Value
        REPLACE POBCTA WITH WinCueDat1.T_Poblacion.Value
        REPLACE CODPOS WITH WinCueDat1.T_CodPos.Value
        REPLACE PROCTA WITH WinCueDat1.T_Provincia.Value
        REPLACE PAIS WITH WinCueDat1.T_Pais.Value
        REPLACE REGIMEN WITH WinCueDat1.C_Regimen.Value-1
        REPLACE CIF WITH WinCueDat1.T_CifCta.Value
        REPLACE TEL1 WITH WinCueDat1.T_Telefono.Value
        REPLACE FAX1 WITH WinCueDat1.T_Fax.Value
        REPLACE BANCTA WITH WinCueDat1.T_CtaBan.Value
        DBCOMMIT()
        DBUNLOCK()

        CP_GRABAR(WinCueDat1.T_CodPos.Value,WinCueDat1.T_Poblacion.Value,WinCueDat1.T_Provincia.Value)

        MsgInfo('Los datos han sido guardados','Datos guardados')
    ELSE
        MsgStop('No se han guardado los datos - '+HB_OsNewLine()+ ;
            'El registro esta siendo utilizado por otro usuario'+HB_OsNewLine()+ ;
            'Por favor, intentelo mas tarde','Error')
    ENDIF

    CueDat_Modif(.F.)

Return Nil

STATIC FUNCTION CueDat_Actualiz(LLAMADA)
    IF VAL(WinCueDat1.T_CodCta.Value)=0
        CueDat_Nuevo("ACTUALIZAR")
        CueDat_Modif(.F.)
        RETURN
    ENDIF

    AbrirDBF("CUENTAS",,,,,1)
    SEEK VAL(WinCueDat1.T_CodCta.Value)

    IF .NOT. EOF()
        WinCueDat1.T_Nombre.Value:=NOMCTA
        WinCueDat1.T_Direccion.Value:=DIRCTA
        WinCueDat1.T_CodPos.Value:=CODPOS
        WinCueDat1.T_Poblacion.Value:=POBCTA
        WinCueDat1.T_Provincia.Value:=PROCTA
        WinCueDat1.T_Pais.Value:=PAIS
        WinCueDat1.C_Regimen.Value:=REGIMEN+1
        WinCueDat1.T_CifCta.Value:=CIF
        WinCueDat1.T_Telefono.Value:=TEL1
        WinCueDat1.T_Fax.Value:=FAX1
        WinCueDat1.T_CtaBan.Value:=BANCTA
        WinCueDat1.T_Saldo.Value:=SALDO

        aPGC:=PGCNOM(WinCueDat1.T_CodCta.Value,4)
        IF aPGC[1]=SPACE(4)
            WinCueDat1.T_Grupo.Value:="Grupo no dado de alta"
        ELSE
            WinCueDat1.T_Grupo.Value:=aPGC[1]+" "+aPGC[2]
        ENDIF
    ELSE
        CueDat_Nuevo("ACTUALIZAR")
    ENDIF

    IF LLAMADA="CANCELAR"
        CueDat_Modif(.F.)
    ENDIF
    CueDat_VerCifBan()
    CueDat_VerRegimen()

Return Nil


STATIC FUNCTION CueDat_Eliminar()
    AbrirDBF("CUENTAS")
    DBSETORDER(1)
    SEEK VAL(WinCueDat1.T_CodCta.Value)

    IF EOF()
        MSGSTOP("La cuenta no esta dada de alta","error")
        RETURN
    ENDIF

    IF MSGYESNO("¿Desea eliminar la cuenta?","Atencion")=.T.
        SIESTA2:=0
        SIESTAT:="Movimientos de la cuenta:"+HB_OsNewLine()
        AbrirDBF("APUNTES")
        LOCATE FOR CODCTA=VAL(WinCueDat1.T_CodCta.Value)
        IF EOF()
            SIESTAT:=SIESTAT+"Apuntes: NO"+HB_OsNewLine()
        ELSE
            SIESTAT:=SIESTAT+"Apuntes: SI"+HB_OsNewLine()
            SIESTA2:=1
        ENDIF
        AbrirDBF("FAC92")
        LOCATE FOR CODCTA=VAL(WinCueDat1.T_CodCta.Value)
        IF EOF()
            SIESTAT:=SIESTAT+"Facturas emitidas: NO"+HB_OsNewLine()
        ELSE
            SIESTAT:=SIESTAT+"Facturas emitidas: SI"+HB_OsNewLine()
            SIESTA2:=1
        ENDIF
        AbrirDBF("FACREB")
        LOCATE FOR CODIGO=VAL(WinCueDat1.T_CodCta.Value)
        IF EOF()
            SIESTAT:=SIESTAT+"Facturas recibidas: NO"+HB_OsNewLine()
        ELSE
            SIESTAT:=SIESTAT+"Facturas recibidas: SI"+HB_OsNewLine()
            SIESTA2:=1
        ENDIF
        IF SIESTA2=1
            MSGSTOP(SIESTAT+"No se puede eliminar la cuenta")
            RETURN
        ELSE
            IF MSGYESNO(SIESTAT+"¿Esta seguro?")=.T.
                AbrirDBF("CUENTAS")
                DBSETORDER(1)
                SEEK VAL(WinCueDat1.T_CodCta.Value)
                IF RLOCK()
                    DELETE
                    DBCOMMIT()
                    DBUNLOCK()
                    MsgInfo('La cuenta ha sido eliminada')
                ELSE
                    MsgStop('No se ha eliminado el registro'+HB_OsNewLine()+ ;
                        'El registro esta siendo utilizado por otro usuario'+HB_OsNewLine()+ ;
                        'Por favor, intentelo mas tarde','Error')
                ENDIF
            ENDIF
        ENDIF
    ENDIF

Return Nil

STATIC FUNCTION CueDat_VerCifBan()
    IF DNI_LET(WinCueDat1.T_CifCta.Value,2)=.T.
        WinCueDat1.T_CifCta.BackColor:=MICOLOR("ROJOCLARO")
        WinCueDat1.L_CifCta.Visible:=.F.
        WinCueDat1.L_CifCtaMal.Visible:=.T.
    ELSE
        WinCueDat1.T_CifCta.BackColor:=COLORSYS("BTNFACE")
        WinCueDat1.L_CifCta.Visible:=.T.
        WinCueDat1.L_CifCtaMal.Visible:=.F.
    ENDIF

    IF CTA_BAN(WinCueDat1.T_CtaBan.Value,1)=0
        WinCueDat1.T_CtaBan.BackColor:=MICOLOR("ROJOCLARO")
        WinCueDat1.L_CtaBan.Visible:=.F.
        WinCueDat1.L_CtaBanMal.Visible:=.T.
    ELSE
        WinCueDat1.T_CtaBan.BackColor:=COLORSYS("BTNFACE")
        WinCueDat1.L_CtaBan.Visible:=.T.
        WinCueDat1.L_CtaBanMal.Visible:=.F.
    ENDIF

Return Nil



STATIC FUNCTION CueDat_DirImp()
    aDireccionImp:={}
    AADD(aDireccionImp, ;
        {WinCueDat1.T_CodCta.Value, ;
        WinCueDat1.T_Nombre.Value, ;
        WinCueDat1.T_Direccion.Value, ;
        WinCueDat1.T_CodPos.Value, ;
        WinCueDat1.T_Poblacion.Value, ;
        WinCueDat1.T_Provincia.Value, ;
        WinCueDat1.T_Pais.Value} )

    AbrirDBF("empresa",,,,RUTAPROGRAMA)
    SEEK NUMEMP
    aDireccionRte:={ ;
        RTRIM(EMP), ;
        RTRIM(Direccion), ;
        RTRIM(CodPos), ;
        RTRIM(Poblacion), ;
        RTRIM(Provincia), ;
        RTRIM(Pais) }

    DireccionImp(1,aDireccionImp,aDireccionRte)

Return Nil
