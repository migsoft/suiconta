#include "minigui.ch"

Function BR_empresa1()
    IF IsWindowDefined(WinBRemp1)=.T.
        DECLARE WINDOW WinBRemp1
        WinBRemp1.restore
        WinBRemp1.Show
        WinBRemp1.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF

    AbrirDBF("empresa",,,,RUTAPROGRAMA)

    DEFINE WINDOW WinBRemp1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Empresas" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("TURQUESA") ;
        ON RELEASE CloseTables()

        @ 10,10 BROWSE BR_emp1 ;
            HEIGHT 170 ;
            WIDTH 770 ;
            HEADERS {'Codigo','Empresa','Ejercicio','cierre','Ruta'} ;
            WIDTHS { 50,300,50,50,200 } ;
            WORKAREA empresa ;
            FIELDS {'empresa->NEMP','empresa->EMP','empresa->ejercicio','empresa->cierrejer','empresa->ruta'} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT} ;
            VALUE RECNOEMPRESA ;
            ON CHANGE ( BR_empresa1_Actualiz("BROWSE") ) ;
            ON HEADCLICK { {|| AbrirDBF("empresa",,,,RUTAPROGRAMA,1), WinBRemp1.BR_emp1.Refresh} , ;
            {|| AbrirDBF("empresa",,,,RUTAPROGRAMA,2), WinBRemp1.BR_emp1.Refresh} }

        @0,0 TEXTBOX RegNuevo WIDTH 100 VALUE 0 NUMERIC RIGHTALIGN INVISIBLE

        @195,10 LABEL L_Codemp VALUE 'Codigo' AUTOSIZE TRANSPARENT
        @190,100 TEXTBOX T_Codemp WIDTH 100 NUMERIC RIGHTALIGN

        @225,10 LABEL L_Nomemp VALUE 'Nombre' AUTOSIZE TRANSPARENT
        @220,100 TEXTBOX T_nomemp WIDTH 250 MAXLENGTH 30

        @255,10 LABEL L_Diremp VALUE 'Direccion' AUTOSIZE TRANSPARENT
        @250,100 TEXTBOX T_Diremp WIDTH 250 MAXLENGTH 30

        @285,10 LABEL L_Poblacion VALUE 'Poblacion' AUTOSIZE TRANSPARENT
        @280,100 TEXTBOX T_CodPos WIDTH 50 TOOLTIP 'Codigo Postal' MAXLENGTH 5 ;
            ON LOSTFOCUS CodPostal("WinBRemp1")
        @280,155 TEXTBOX T_Poblacion WIDTH 195 MAXLENGTH 30

        @315,10 LABEL L_Provincia VALUE 'Provincia' AUTOSIZE TRANSPARENT
        @310,100 TEXTBOX T_Provincia WIDTH 250 MAXLENGTH 30

        @345,10 LABEL L_Cifemp VALUE 'CIF' AUTOSIZE TRANSPARENT
        @340,100 TEXTBOX T_cifemp WIDTH 250 MAXLENGTH 15

        @375,10 LABEL L_Banco VALUE 'Cuenta banco' AUTOSIZE TRANSPARENT
        @370,100 TEXTBOX T_Banco WIDTH 250 MAXLENGTH 30

        @405,10 LABEL L_email VALUE 'email' AUTOSIZE TRANSPARENT
        @400,100 TEXTBOX T_email WIDTH 250 MAXLENGTH 40

        @435,10 LABEL L_Obs VALUE 'Observaciones' AUTOSIZE TRANSPARENT
        @430,100 TEXTBOX T_Obs WIDTH 250 MAXLENGTH 30

        aIMP:=Impresoras(EMP_IMPRESORA)
        AADD(aIMP[1],"Impresora predeterminada")
        @465,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @460,100 COMBOBOX C_Impresora ;
            WIDTH 250 ;
            ITEMS aIMP[1] ;
            VALUE aIMP[3] ;
            TOOLTIP 'Impresora' NOTABSTOP


***************************

        @195,400 LABEL L_ejercicio VALUE 'Ejercicio' AUTOSIZE TRANSPARENT
        @190,500 TEXTBOX T_Ejercicio WIDTH 100 NUMERIC RIGHTALIGN

        @225,400 LABEL L_Telemp VALUE 'Telefono' AUTOSIZE TRANSPARENT
        @220,500 TEXTBOX T_Telemp WIDTH 150 MAXLENGTH 15

        @255,400 LABEL L_Faxemp VALUE 'Fax' AUTOSIZE TRANSPARENT
        @250,500 TEXTBOX T_Faxemp WIDTH 150 MAXLENGTH 15

        @285,400 LABEL L_Movemp VALUE 'Movil' AUTOSIZE TRANSPARENT
        @280,500 TEXTBOX T_Movemp WIDTH 150 MAXLENGTH 15

        @315,400 LABEL L_Empant VALUE 'Empresa anterior' AUTOSIZE TRANSPARENT
        @310,525 TEXTBOX T_Empant WIDTH 50 NUMERIC RIGHTALIGN

        @345,400 LABEL L_Empsig VALUE 'Empresa siguiente' AUTOSIZE TRANSPARENT
        @340,525 TEXTBOX T_Empsig WIDTH 50 NUMERIC RIGHTALIGN

        @375,400 LABEL L_Empbus VALUE 'Buscar empresas' AUTOSIZE TRANSPARENT
        @370,525 TEXTBOX T_Empbus WIDTH 50 NUMERIC RIGHTALIGN

        @405,400 LABEL L_ruta VALUE 'Ruta empresa' AUTOSIZE TRANSPARENT
        @400,500 TEXTBOX T_Ruta WIDTH 200 MAXLENGTH 15

        @460,400 CHECKBOX C_cierre CAPTION 'Ejercicio cerrado' WIDTH 120 VALUE .F. ;
            BACKCOLOR MiColor("TURQUESA")


        @500,010 BUTTONEX Bt_CopiarDir CAPTION 'Copiar al portapapeles' ICON icobus('portapapeles') ;
            ACTION BR_empresa1_CopiarPP() ;
            WIDTH 200 HEIGHT 25 NOTABSTOP

        @530,010 BUTTONEX Bt_Nueva CAPTION 'Nueva' ICON icobus('nuevo') WIDTH 95 HEIGHT 25 ;
            ACTION BR_empresa1_Nueva("NUEVO")

        @530,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') WIDTH 95 HEIGHT 25 ;
            ACTION BR_empresa1_Modif(.t.)

        @530,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') WIDTH 95 HEIGHT 25 ;
            ACTION BR_empresa1_Guardar()

        @530,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 95 HEIGHT 25 ;
            ACTION BR_empresa1_Actualiz("CANCELAR")

        @530,410 BUTTONEX Bt_Elim CAPTION 'Eliminar' ICON icobus('eliminar') WIDTH 95 HEIGHT 25 ;
/*
        @530,510 BUTTONEX Bt_Selec CAPTION 'Seleccionar' WIDTH 90 HEIGHT 25 ;
            ACTION BR_empresa1_Selec("EMPRESA",WinBRemp1.BR_emp1.Value)
*/
        @530,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 95 HEIGHT 25 ;
            ACTION WinBRemp1.Release ;

        WinBRemp1.Bt_Elim.Enabled:=.F.

    END WINDOW
    VentanaCentrar("WinBRemp1","Ventana1","Alinear")

    CENTER WINDOW WinBRemp1
    ACTIVATE WINDOW WinBRemp1

Return nil


STATIC FUNCTION BR_empresa1_Nueva(LLAMADA)
    SiApertura:=Box3botones("Nuevo ejercicio",{"¿Desea aperturar el siguiente ejercicio?", ;
        WinBRemp1.T_Nomemp.Value+" "+STR(WinBRemp1.T_Ejercicio.Value+1,4)}, ;
        "Aperturar","Nuevo","Cancelar")
    DO CASE
    CASE SiApertura=1
        IF WinBRemp1.T_Empsig.Value<>0
            MSGSTOP("El siguiente ejercicio esta aperturado","error")
            RETURN
        ENDIF
        EmpApe('EMPRESA')
        AbrirDBF("empresa",,,,RUTAPROGRAMA)
        WinBRemp1.BR_emp1.Refresh
    CASE SiApertura=2
        NuevaEmpresaVacia()
    ENDCASE
Return Nil



STATIC FUNCTION BR_empresa1_Modif(Modif)
    SoloVer:=IF(Modif=.T.,.F.,.T.)
    WinBRemp1.T_Codemp.Enabled   := .F.
    WinBRemp1.T_nomemp.ReadOnly  := SoloVer
    WinBRemp1.T_Diremp.ReadOnly  := SoloVer
    WinBRemp1.T_CodPos.ReadOnly  := SoloVer
    WinBRemp1.T_Poblacion.ReadOnly  := SoloVer
    WinBRemp1.T_Provincia.ReadOnly  := SoloVer
    WinBRemp1.T_cifemp.ReadOnly  := SoloVer
    WinBRemp1.T_Banco.ReadOnly   := SoloVer
    WinBRemp1.T_email.ReadOnly   := SoloVer
    WinBRemp1.T_Obs.ReadOnly     := SoloVer
    WinBRemp1.C_Impresora.Enabled:= Modif
    WinBRemp1.T_Ejercicio.Enabled:= .F.
    WinBRemp1.T_Telemp.ReadOnly  := SoloVer
    WinBRemp1.T_Faxemp.ReadOnly  := SoloVer
    WinBRemp1.T_Movemp.ReadOnly  := SoloVer
    WinBRemp1.T_Empant.Enabled   := .F.
    WinBRemp1.T_Empsig.Enabled   := .F.
    WinBRemp1.T_Empbus.ReadOnly  := SoloVer
    WinBRemp1.T_Ruta.Enabled     := .F.
    WinBRemp1.C_cierre.Enabled   := Modif
    IF Modif = .T.
        WinBRemp1.Bt_Nueva.Enabled   := .F.
        WinBRemp1.Bt_Modif.Enabled   := .F.
        WinBRemp1.Bt_Guardar.Enabled := .T.
        WinBRemp1.Bt_Cancelar.Enabled:= .T.
        WinBRemp1.Bt_Elim.Enabled    := IF(WinBRemp1.BR_emp1.Value=0,.T.,.F.)
**   WinBRemp1.Bt_Selec.Enabled   := .F.
        WinBRemp1.Bt_Salir.Enabled   := .F.
    ELSE
        WinBRemp1.Bt_Nueva.Enabled   := .T.
        WinBRemp1.Bt_Modif.Enabled   := IF(WinBRemp1.BR_emp1.Value=0,.F.,.T.)
        WinBRemp1.Bt_Guardar.Enabled := .F.
        WinBRemp1.Bt_Cancelar.Enabled:= .F.
        WinBRemp1.Bt_Elim.Enabled    := .F.
**   WinBRemp1.Bt_Selec.Enabled   := .T.
        WinBRemp1.Bt_Salir.Enabled   := .T.
    ENDIF
Return Nil

STATIC FUNCTION BR_empresa1_Guardar()

    AbrirDBF("empresa",,,,RUTAPROGRAMA)
    IF WinBRemp1.BR_emp1.Value=0
        MSGSTOP("No se ha seleccionado ninguna empresa")
        RETURN
    ENDIF
    GO WinBRemp1.BR_emp1.Value

    IF RLOCK()
        REPLACE EMP       WITH WinBRemp1.T_Nomemp.Value
        REPLACE DIRECCION WITH WinBRemp1.T_Diremp.Value
        IF CAMPO("CODPOS")=.T.
            REPLACE CODPOS    WITH WinBRemp1.T_CodPos.Value
        ENDIF
        REPLACE POBLACION WITH WinBRemp1.T_Poblacion.Value
        REPLACE PROVINCIA WITH WinBRemp1.T_Provincia.Value
        REPLACE CIF       WITH WinBRemp1.T_cifemp.Value
        REPLACE BANCTA    WITH WinBRemp1.T_Banco.Value
        REPLACE EMAIL     WITH WinBRemp1.T_email.Value
        REPLACE OBS       WITH WinBRemp1.T_Obs.Value
        REPLACE EJERCICIO WITH WinBRemp1.T_Ejercicio.Value
        REPLACE TEL1      WITH WinBRemp1.T_Telemp.Value
        REPLACE FAX1      WITH WinBRemp1.T_Faxemp.Value
        REPLACE MOVIL     WITH WinBRemp1.T_Movemp.Value
        REPLACE NEMPANT   WITH WinBRemp1.T_Empant.Value
        REPLACE NEMPSIG   WITH WinBRemp1.T_Empsig.Value
        REPLACE NEMPBUS   WITH WinBRemp1.T_Empbus.Value
        REPLACE RUTA      WITH WinBRemp1.T_Ruta.Value
        REPLACE CIERREJER WITH IF(WinBRemp1.C_cierre.Value=.T.,1,0)
        IF CAMPO("IMPRESORA")
            REPLACE IMPRESORA WITH WinBRemp1.C_Impresora.Item(WinBRemp1.C_Impresora.Value)
            EMP_IMPRESORA:=IMPRESORA
        ENDIF
        DBCOMMIT()
        DBUNLOCK()
        MsgInfo('Los datos han sido guardados en el codigo: '+LTRIM(STR(NEMP)),'Datos guardados')
        WinBRemp1.BR_emp1.Refresh

    ELSE
        MsgStop('No se han guardado los datos - '+HB_OsNewLine()+ ;
            'El registro esta siendo utilizado por otro usuario'+HB_OsNewLine()+ ;
            'Por favor, intentelo mas tarde','Error')
    ENDIF

    BR_empresa1_Modif(.f.)
return nil


STATIC FUNCTION BR_empresa1_Actualiz(LLAMADA)

    AbrirDBF("empresa",,,,RUTAPROGRAMA)
    GO WinBRemp1.BR_emp1.Value

    IF .NOT. EOF()
        WinBRemp1.T_Codemp.Value   := NEMP
        WinBRemp1.T_nomemp.Value   := EMP
        WinBRemp1.T_Diremp.Value   := DIRECCION
        WinBRemp1.T_CodPos.Value   := IF(CAMPO("CODPOS"),CODPOS,"")
        WinBRemp1.T_Poblacion.Value:= POBLACION
        WinBRemp1.T_Provincia.Value:= PROVINCIA
        WinBRemp1.T_cifemp.Value   := CIF
        WinBRemp1.T_Banco.Value    := BANCTA
        WinBRemp1.T_email.Value    := EMAIL
        WinBRemp1.T_Obs.Value      := OBS
        WinBRemp1.T_Ejercicio.Value:= EJERCICIO
        WinBRemp1.T_Telemp.Value   := TEL1
        WinBRemp1.T_Faxemp.Value   := FAX1
        WinBRemp1.T_Movemp.Value   := MOVIL
        WinBRemp1.T_Empant.Value   := NEMPANT
        WinBRemp1.T_Empsig.Value   := NEMPSIG
        WinBRemp1.T_Empbus.Value   := NEMPBUS
        WinBRemp1.T_Ruta.Value     := RUTA
        WinBRemp1.C_cierre.Value   := IF(CIERREJER=1,.T.,.F.)
        IF CAMPO("IMPRESORA")
            FOR N=1 TO WinBRemp1.C_Impresora.ItemCount
                IF PADR(WinBRemp1.C_Impresora.Item(N),LEN(IMPRESORA)," ")==IMPRESORA
                    WinBRemp1.C_Impresora.Value:=N
                    EXIT
                ENDIF
            NEXT
        ENDIF
    ELSE
        WinBRemp1.T_Codemp.Value   := 0
        WinBRemp1.T_nomemp.Value   := ""
        WinBRemp1.T_Diremp.Value   := ""
        WinBRemp1.T_CodPos.Value   := ""
        WinBRemp1.T_Poblacion.Value   := ""
        WinBRemp1.T_Provincia.Value   := ""
        WinBRemp1.T_cifemp.Value   := ""
        WinBRemp1.T_Banco.Value    := ""
        WinBRemp1.T_email.Value    := ""
        WinBRemp1.T_Obs.Value      := ""
        WinBRemp1.C_Impresora.Value:= 1
        WinBRemp1.T_Ejercicio.Value:= 0
        WinBRemp1.T_Telemp.Value   := ""
        WinBRemp1.T_Faxemp.Value   := ""
        WinBRemp1.T_Movemp.Value   := ""
        WinBRemp1.T_Empant.Value   := 0
        WinBRemp1.T_Empsig.Value   := 0
        WinBRemp1.T_Empbus.Value   := 0
        WinBRemp1.T_Ruta.Value     := ""
        WinBRemp1.C_cierre.Value   := .F.
    ENDIF

    IF LLAMADA<>"BROWSE"
        WinBRemp1.BR_emp1.Value:=recno()
        WinBRemp1.BR_emp1.Refresh
    ENDIF

    BR_empresa1_Modif(.f.)

Return Nil

FUNCTION NuevaEmpresaVacia()
    IF MSGYESNO("¿Desea aperturar una empresa vacia?")=.F.
        RETURN
    ENDIF
    AbrirDBF("empresa",,,,RUTAPROGRAMA,1)
    NUMEMP2:=LASTREC()+1

    DO WHILE .T.
        RESULTADO2:=InputWindow("Aperturar una empresa vacia",{"Numero empresa"},{NUMEMP2},{"Numeric"},300,300,.T.,{"Aperturar","Cancelar"})
        IF VALTYPE(RESULTADO2[1])<>"N"
            RETURN
        ENDIF
        SEEK RESULTADO2[1]
        IF .NOT. EOF()
            MSGSTOP("La empresa ya esta aperturada"+HB_OsNewLine()+STR(NEMP)+"-"+LTRIM(EMP)+" "+STR(EJERCICIO))
        ELSE
            EXIT
        ENDIF
    ENDDO

    AbrirDBF("empresa",,,,RUTAPROGRAMA,1)
    APPEND BLANK
    IF RLOCK()
        REPLACE NEMP WITH RESULTADO2[1]
        REPLACE EMP  WITH "Empresa nueva"
        REPLACE RUTA WITH "Suizo"+STRZERO(RESULTADO2[1],3)
        REPLACE EJERCICIO WITH YEAR(DATE())
        DBCOMMIT()
        DBUNLOCK()
        CreateFolder(RUTAPROGRAMA+"\Suizo"+STRZERO(RESULTADO2[1],3))
    ENDIF

    RUTAEMPRESA2APE:=RUTAEMPRESA
    RUTAEMPRESA:=RUTAPROGRAMA+"\Suizo"+STRZERO(RESULTADO2[1],3)
    SetCurrentFolder(RUTAEMPRESA)
    set default to &RUTAEMPRESA
    W_RegficConta("TODOS","SINVENTANA")
    RUTAEMPRESA:=RUTAEMPRESA2APE
    SetCurrentFolder(RUTAEMPRESA)
    set default to &RUTAEMPRESA

    AbrirDBF("empresa",,,,RUTAPROGRAMA,1)
    WinBRemp1.BR_emp1.Refresh
    MSGBOX("La empresa ha sido aperturada con exito")

RETURN Nil




STATIC FUNCTION BR_empresa1_CopiarPP()
    Texto1:=RTRIM(WinBRemp1.T_nomemp.Value)+HB_OsNewLine()+ ;
        RTRIM(WinBRemp1.T_Diremp.Value)+HB_OsNewLine()+ ;
        RTRIM(WinBRemp1.T_CodPos.Value)+"-"+RTRIM(WinBRemp1.T_Poblacion.Value)+HB_OsNewLine()+ ;
        RTRIM(WinBRemp1.T_Provincia.Value)+HB_OsNewLine()+ ;
        "Tel."+RTRIM(WinBRemp1.T_Telemp.Value)+HB_OsNewLine()+ ;
        "Fax."+RTRIM(WinBRemp1.T_Faxemp.Value)+HB_OsNewLine()+ ;
        "Movil "+RTRIM(WinBRemp1.T_Movemp.Value)+HB_OsNewLine()+ ;
        "Cif "+RTRIM(WinBRemp1.T_cifemp.Value)+HB_OsNewLine()+ ;
        "email "+RTRIM(WinBRemp1.T_email.Value)+HB_OsNewLine()+ ;
        RTRIM(WinBRemp1.T_Banco.Value)
    //*** CopyToClipboard(Texto1)
    SetClipboardText(Texto1)
    AutoMsgInfo(Texto1,"Portapapeles")
Return Nil
