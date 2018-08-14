#include "minigui.ch"

Function Cuentas_Actualizar(LLAMADA,CODIGOACT)

    DEFINE WINDOW WinCueAct ;
        AT 0,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Actualizar cuentas" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("CIAN") ;
        ON RELEASE CloseTables()

        aPROG:={"Contaplus -Subcuentas", ;
            "Facturaplus -Clientes", ;
            "Facturaplus -Proveedores", ;
            "Suizo Contable -Subcuentas", ;
            "Suizo Comercial -Clientes", ;
            "Suizo Comercial -Proveedores", ;
            "Suizo Remesa -Subcuentas", ;
            "Suizo Escola -Alumnos", ;
            "Suizo Escola -Educadoras"}

        @0,0 CHECKBOX Calcular CAPTION '' VALUE .F. INVISIBLE

        @015,010 LABEL L_Programa1 VALUE 'Programa' AUTOSIZE TRANSPARENT
        @010,100 COMBOBOX C_Programa1 WIDTH 250 ITEMS aPROG ;
            VALUE 2 ON CHANGE CuentasActRuta(1,WinCueAct.T_Ruta1.Value)

        @040,010 BUTTONEX Bt_ruta1 CAPTION 'Ruta' ICON icobus('buscar') ;
            ACTION ( WinCueAct.T_Ruta1.Value:=Getfolder( 'Buscar ruta', WinCueAct.T_Ruta1.Value ) , ;
            CuentasActRuta(1,WinCueAct.T_Ruta1.Value) ) ;
            WIDTH 50 HEIGHT 25
        @040,060 BUTTON Bt_ruta1r CAPTION 'Red' ;
            ACTION ( WinCueAct.T_Ruta1.Value:=BrowseForFolder(0) , ;
            CuentasActRuta(1,WinCueAct.T_Ruta1.Value) ) ;
            WIDTH 40 HEIGHT 25
        @040,100 TEXTBOX T_Ruta1 WIDTH 250 VALUE GetCurrentFolder() READONLY ON CHANGE CuentasActEmp(1,WinCueAct.T_Ruta1.Value)

        @075,010 LABEL L_Empresa1 VALUE 'Empresa' AUTOSIZE TRANSPARENT
        @070,100 COMBOBOX C_Empresa1 WIDTH 250 ITEMS {} ON CHANGE WinCueAct.C_RutaEmp1.Value:=WinCueAct.C_Empresa1.Value
        @100,100 COMBOBOX C_RutaEmp1 WIDTH 250 ITEMS {} INVISIBLE


        @015,410 LABEL L_Programa2 VALUE 'Programa' AUTOSIZE TRANSPARENT
        @010,500 COMBOBOX C_Programa2 WIDTH 250 ITEMS aPROG ;
            VALUE 1 ON CHANGE CuentasActRuta(2,WinCueAct.T_Ruta2.Value)

        @040,410 BUTTONEX Bt_ruta2 CAPTION 'Ruta' ICON icobus('buscar') ;
            ACTION ( WinCueAct.T_Ruta2.Value:=Getfolder( 'Buscar ruta', WinCueAct.T_Ruta2.Value ) , ;
            CuentasActRuta(2,WinCueAct.T_Ruta2.Value) ) ;
            WIDTH 50 HEIGHT 25
        @040,460 BUTTON Bt_ruta2r CAPTION 'Red' ;
            ACTION ( WinCueAct.T_Ruta2.Value:=BrowseForFolder(0) , ;
            CuentasActRuta(2,WinCueAct.T_Ruta2.Value) ) ;
            WIDTH 40 HEIGHT 25
        @040,500 TEXTBOX T_Ruta2 WIDTH 250 VALUE GetCurrentFolder() READONLY ON CHANGE CuentasActEmp(2,WinCueAct.T_Ruta2.Value)

        @075,410 LABEL L_Empresa2 VALUE 'Empresa' AUTOSIZE TRANSPARENT
        @070,500 COMBOBOX C_Empresa2 WIDTH 250 ITEMS {} ON CHANGE WinCueAct.C_RutaEmp2.Value:=WinCueAct.C_Empresa2.Value
        @100,500 COMBOBOX C_RutaEmp2 WIDTH 250 ITEMS {} INVISIBLE

        DRAW LINE IN WINDOW WinCueAct AT 100,010 TO 100,790

        @115,010 LABEL L_Codigo VALUE 'Codigo' AUTOSIZE TRANSPARENT
        @110,100 TEXTBOX T_DatoA1 WIDTH 100 VALUE 0 READONLY NUMERIC RIGHTALIGN
        @110,500 TEXTBOX T_DatoB1 WIDTH 100 VALUE 0 READONLY NUMERIC RIGHTALIGN

        @145,010 LABEL L_Nombre VALUE 'Nombre' AUTOSIZE TRANSPARENT
        @140,100 TEXTBOX T_DatoA2 WIDTH 250 VALUE "" READONLY
        @140,500 TEXTBOX T_DatoB2 WIDTH 250 VALUE "" READONLY

        @175,010 LABEL L_Direccion1 VALUE 'Direccion' AUTOSIZE TRANSPARENT
        @170,100 TEXTBOX T_DatoA3 WIDTH 250 VALUE "" READONLY
        @170,500 TEXTBOX T_DatoB3 WIDTH 250 VALUE "" READONLY

        @205,010 LABEL L_Poblacion1 VALUE 'Poblacion' AUTOSIZE TRANSPARENT
        @200,100 TEXTBOX T_DatoA4 WIDTH 250 VALUE "" READONLY
        @200,500 TEXTBOX T_DatoB4 WIDTH 250 VALUE "" READONLY

        @235,010 LABEL L_Provincia1 VALUE 'Provincia' AUTOSIZE TRANSPARENT
        @230,100 TEXTBOX T_DatoA5 WIDTH 250 VALUE "" READONLY
        @230,500 TEXTBOX T_DatoB5 WIDTH 250 VALUE "" READONLY

        @265,010 LABEL L_CodPos1 VALUE 'Codigo Postal' AUTOSIZE TRANSPARENT
        @260,100 TEXTBOX T_DatoA6 WIDTH 250 VALUE "" READONLY
        @260,500 TEXTBOX T_DatoB6 WIDTH 250 VALUE "" READONLY

        @295,010 LABEL L_Cif1 VALUE 'CIF' AUTOSIZE TRANSPARENT
        @290,100 TEXTBOX T_DatoA7 WIDTH 250 VALUE "" READONLY
        @290,500 TEXTBOX T_DatoB7 WIDTH 250 VALUE "" READONLY

        @325,010 LABEL L_Telefono1 VALUE 'Telefono' AUTOSIZE TRANSPARENT
        @320,100 TEXTBOX T_DatoA8 WIDTH 250 VALUE "" READONLY
        @320,500 TEXTBOX T_DatoB8 WIDTH 250 VALUE "" READONLY

        @355,010 LABEL L_Fax1 VALUE 'Fax' AUTOSIZE TRANSPARENT
        @350,100 TEXTBOX T_DatoA9 WIDTH 250 VALUE "" READONLY
        @350,500 TEXTBOX T_DatoB9 WIDTH 250 VALUE "" READONLY

        @385,010 LABEL L_email1 VALUE 'email' AUTOSIZE TRANSPARENT
        @380,100 TEXTBOX T_DatoA10 WIDTH 250 VALUE "" READONLY
        @380,500 TEXTBOX T_DatoB10 WIDTH 250 VALUE "" READONLY

        @415,010 LABEL L_CtaBan1 VALUE 'Cuenta banco' AUTOSIZE TRANSPARENT
        @410,100 TEXTBOX T_DatoA11 WIDTH 250 VALUE "" READONLY
        @410,500 TEXTBOX T_DatoB11 WIDTH 250 VALUE "" READONLY

        @110,375 CHECKBOX C_Si1 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @140,375 CHECKBOX C_Si2 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @170,375 CHECKBOX C_Si3 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @200,375 CHECKBOX C_Si4 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @230,375 CHECKBOX C_Si5 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @260,375 CHECKBOX C_Si6 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @290,375 CHECKBOX C_Si7 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @320,375 CHECKBOX C_Si8 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @350,375 CHECKBOX C_Si9 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @380,375 CHECKBOX C_Si10 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP
        @410,375 CHECKBOX C_Si11 CAPTION 'Actualizar' WIDTH 100 VALUE .F. TRANSPARENT NOTABSTOP

        @440,100 TEXTBOX T_Registro1  WIDTH 100 VALUE 0 READONLY NUMERIC
        @440,250 TEXTBOX T_Registro1t WIDTH 100 VALUE 0 READONLY NUMERIC
        @440,500 TEXTBOX T_Registro2  WIDTH 100 VALUE 0 READONLY NUMERIC
        @440,650 TEXTBOX T_Registro2t WIDTH 100 VALUE 0 READONLY NUMERIC

        @470,100 PROGRESSBAR Progress_1 WIDTH 650 HEIGHT 20 SMOOTH


*      @505,350 LABEL L_DirAct VALUE 'Direccion a actualizar ----------->>>' AUTOSIZE TRANSPARENT


        @500,150 BUTTONEX Bt_Alreves CAPTION 'Actualizar' ICON icobus('aceptar') WIDTH 90 HEIGHT 25 ;
            ACTION ( CuentasActFic("ALREVES") , IF(CODIGOACT=NIL,,WinCueAct.Release) ) ;
            NOTABSTOP

        @500,550 BUTTONEX Bt_Actualizar CAPTION 'Actualizar' ICON icobus('aceptar') WIDTH 90 HEIGHT 25 ;
            ACTION ( CuentasActFic("ACTUALIZAR") , IF(CODIGOACT=NIL,,WinCueAct.Release) ) ;
            NOTABSTOP



        @530,210 BUTTONEX Bt_Inicio CAPTION 'Inicio' ICON icobus('consultar') WIDTH 90 HEIGHT 25 ;
            ACTION CuentasActComprobar("INICIO") ;
            NOTABSTOP

        @530,310 BUTTONEX Bt_Continuar CAPTION 'Continuar' ICON icobus('consultar') WIDTH 90 HEIGHT 25 ;
            ACTION CuentasActComprobar("COMPROBAR") ;
            NOTABSTOP

        @530,410 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION CuentasActComprobar("CANCELAR") ;
            NOTABSTOP

        @530,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION WinCueAct.Release ;
            NOTABSTOP

        DO CASE
        CASE LLAMADA="SUICONTA"
            WinCueAct.C_Programa1.Value:=5
            WinCueAct.C_Programa2.Value:=4
        CASE LLAMADA="CAJA"
            WinCueAct.C_Programa1.Value:=2
            WinCueAct.C_Programa2.Value:=1
        CASE LLAMADA="ALUMNOS"
            WinCueAct.C_Programa1.Value:=8
            WinCueAct.C_Programa2.Value:=4
        CASE LLAMADA="EDUCADORAS"
            WinCueAct.C_Programa1.Value:=9
            WinCueAct.C_Programa2.Value:=4
        ENDCASE
        WinCueAct.Calcular.Value:=.T.
        CuentasActRuta(1,WinCueAct.T_Ruta1.Value)
        CuentasActRuta(2,WinCueAct.T_Ruta2.Value)

        CuentasActEmp(1,WinCueAct.T_Ruta1.Value)
        CuentasActEmp(2,WinCueAct.T_Ruta2.Value)

        VerBotones("INICIO")

        IF !CODIGOACT=NIL
            WinCueAct.Bt_Inicio.Enabled   :=.F.
            WinCueAct.Bt_Continuar.Enabled:=.F.
            WinCueAct.Bt_Cancelar.Enabled :=.F.
            CuentasActComprobar("INICIO",CODIGOACT)
        ENDIF

    END WINDOW

    VentanaCentrar("WinCueAct","Ventana1","Alinear")
    CENTER WINDOW WinCueAct
    ACTIVATE WINDOW WinCueAct

Return Nil

STATIC FUNCTION VerBotones(LLAMADA)
    DO CASE
    CASE LLAMADA="INICIO"
        WinCueAct.Bt_ruta1.Enabled     :=.T.
        WinCueAct.Bt_ruta2.Enabled     :=.T.
        WinCueAct.Bt_Alreves.Enabled   :=.F.
        WinCueAct.Bt_Actualizar.Enabled:=.F.
        WinCueAct.Bt_Inicio.Enabled    :=.T.
        WinCueAct.Bt_Continuar.Enabled :=.T.
        WinCueAct.Bt_Cancelar.Enabled  :=.T.
        WinCueAct.Bt_Salir.Enabled     :=.T.
    CASE LLAMADA="ALTA"
        WinCueAct.Bt_ruta1.Enabled     :=.F.
        WinCueAct.Bt_ruta2.Enabled     :=.F.
        WinCueAct.Bt_Alreves.Enabled   :=.F.
        WinCueAct.Bt_Actualizar.Enabled:=.T.
        WinCueAct.Bt_Inicio.Enabled    :=.T.
        WinCueAct.Bt_Continuar.Enabled :=.T.
        WinCueAct.Bt_Cancelar.Enabled  :=.T.
        WinCueAct.Bt_Salir.Enabled     :=.T.
    CASE LLAMADA="ACTUALIZAR"
        WinCueAct.Bt_ruta1.Enabled     :=.F.
        WinCueAct.Bt_ruta2.Enabled     :=.F.
        WinCueAct.Bt_Alreves.Enabled   :=.T.
        WinCueAct.Bt_Actualizar.Enabled:=.T.
        WinCueAct.Bt_Inicio.Enabled    :=.T.
        WinCueAct.Bt_Continuar.Enabled :=.T.
        WinCueAct.Bt_Cancelar.Enabled  :=.T.
        WinCueAct.Bt_Salir.Enabled     :=.T.
    CASE LLAMADA="COMPROBANDO"
        WinCueAct.Bt_ruta1.Enabled     :=.F.
        WinCueAct.Bt_ruta2.Enabled     :=.F.
        WinCueAct.Bt_Alreves.Enabled   :=.F.
        WinCueAct.Bt_Actualizar.Enabled:=.F.
        WinCueAct.Bt_Inicio.Enabled    :=.F.
        WinCueAct.Bt_Continuar.Enabled :=.F.
        WinCueAct.Bt_Cancelar.Enabled  :=.F.
        WinCueAct.Bt_Salir.Enabled     :=.F.
    ENDCASE

Return Nil


STATIC FUNCTION CuentasActRuta(NIVEL2,RUTA2)
    IF NIVEL2=1
        PROGRAMA2:=WinCueAct.C_Programa1.Value
    ELSE
        PROGRAMA2:=WinCueAct.C_Programa2.Value
    ENDIF
***EVITAR RUTA VACIA***
    IF LEN(RTRIM(RUTA2))=0
        RUTA2:=GetCurrentFolder()
    ENDIF
***FIN EVITAR RUTA VACIA***

    RUTA4:=""
    DO CASE
    CASE PROGRAMA2=1                                //Contaplus
        IF AT("GRUPOSP",UPPER(RUTA2))<>0
            DO CASE
            CASE FILE(RUTA2+"\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\Empresa.cdx")
                RUTA4:=RUTA2
            CASE FILE(RUTA2+"\Emp\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\Emp\Empresa.cdx")
                RUTA4:=RUTA2+"\Emp"
            CASE FILE(LEFT(RUTA2,AT("EMP",UPPER(RUTA2))+2)+"\Empresa.dbf") .AND. ;
                FILE(LEFT(RUTA2,AT("EMP",UPPER(RUTA2))+2)+"\Empresa.cdx")
                RUTA4:=LEFT(RUTA2,AT("EMP",UPPER(RUTA2))+2)
            OTHERWISE
                RUTA2B:=LEFT(RUTA2,AT("GRUPOSP",UPPER(RUTA2))+6)
                aCARPETAS2:=DIRECTORY(RUTA2B+"\CO*.","D")
                IF LEN(aCARPETAS2)>0
                    ASORT( aCARPETAS2,,, {| x, y | DTOS(x[3]) > DTOS(y[3]) } )
                    FOR N=1 TO LEN(aCARPETAS2)
                        IF FILE(RUTA2B+"\"+aCARPETAS2[N,1]+"\Emp\Empresa.dbf")  .AND. ;
                            FILE(RUTA2B+"\"+aCARPETAS2[N,1]+"\Emp\Empresa.cdx")
                            RUTA4:=RUTA2B+"\"+aCARPETAS2[N,1]+"\Emp"
                            EXIT
                        ENDIF
                    NEXT
                ENDIF
            ENDCASE
        ENDIF
    CASE PROGRAMA2=2 .OR. PROGRAMA2=3               //Facturaplus
        IF AT("GRUPOSP",UPPER(RUTA2))<>0
            DO CASE
            CASE FILE(RUTA2+"\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\Empresa.cdx")
                RUTA4:=RUTA2
            CASE FILE(RUTA2+"\dbf\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\dbf\Empresa.cdx")
                RUTA4:=RUTA2+"\dbf"
            CASE FILE(LEFT(RUTA2,AT("DBF",UPPER(RUTA2))+2)+"\Empresa.dbf") .AND. ;
                FILE(LEFT(RUTA2,AT("DBF",UPPER(RUTA2))+2)+"\Empresa.cdx")
                RUTA4:=LEFT(RUTA2,AT("DBF",UPPER(RUTA2))+2)
            OTHERWISE
                RUTA2B:=LEFT(RUTA2,AT("GRUPOSP",UPPER(RUTA2))+6)
                aCARPETAS2:=DIRECTORY(RUTA2B+"\FA*.","D")
                IF LEN(aCARPETAS2)>0
                    ASORT( aCARPETAS2,,, {| x, y | DTOS(x[3]) > DTOS(y[3]) } )
                    FOR N=1 TO LEN(aCARPETAS2)
                        IF FILE(RUTA2B+"\"+aCARPETAS2[N,1]+"\dbf\Empresa.dbf")  .AND. ;
                            FILE(RUTA2B+"\"+aCARPETAS2[N,1]+"\dbf\Empresa.cdx")
                            RUTA4:=RUTA2B+"\"+aCARPETAS2[N,1]+"\dbf"
                            EXIT
                        ENDIF
                    NEXT
                ENDIF
            ENDCASE
        ENDIF
    CASE PROGRAMA2=4                                //Suizo Contable
        DO CASE
        CASE AT("SUICONTA",UPPER(RUTA2))<>0
            DO CASE
            CASE FILE(RUTA2+"\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\Empresa.cdx")
                RUTA4:=RUTA2
            CASE FILE(LEFT(RUTA2,AT("SUICONTA",UPPER(RUTA2))+7)+"\Empresa.dbf") .AND. ;
                FILE(LEFT(RUTA2,AT("SUICONTA",UPPER(RUTA2))+7)+"\Empresa.cdx")
                RUTA4:=LEFT(RUTA2,AT("SUICONTA",UPPER(RUTA2))+7)
            ENDCASE
        CASE AT("SUIZOWIN",UPPER(RUTA2))<>0
            RUTA5:=SUBSTR(RUTA2,AT("SUIZOWIN",UPPER(RUTA2)),LEN(RUTA2))
            RUTA5:=LEFT(RUTA5,AT("\",RUTA5)-1)
            RUTA4:=LEFT(RUTA2,AT("SUIZOWIN",UPPER(RUTA2))-1)+RUTA5+"\Suiconta"
        ENDCASE
    CASE PROGRAMA2=5 .OR. PROGRAMA2=6               //Suizo Comercial
        DO CASE
        CASE AT("SUIALMA",UPPER(RUTA2))<>0
            DO CASE
            CASE FILE(RUTA2+"\Empcon.dbf") .AND. ;
                FILE(RUTA2+"\Empcon.cdx")
                RUTA4:=RUTA2
            CASE FILE(LEFT(RUTA2,AT("SUIALMA",UPPER(RUTA2))+6)+"\Empcon.dbf") .AND. ;
                FILE(LEFT(RUTA2,AT("SUIALMA",UPPER(RUTA2))+6)+"\Empcon.cdx")
                RUTA4:=LEFT(RUTA2,AT("SUIALMA",UPPER(RUTA2))+6)
            ENDCASE
        CASE AT("SUIZOWIN",UPPER(RUTA2))<>0
            RUTA5:=SUBSTR(RUTA2,AT("SUIZOWIN",UPPER(RUTA2)),LEN(RUTA2))
            RUTA5:=LEFT(RUTA5,AT("\",RUTA5)-1)
            RUTA4:=LEFT(RUTA2,AT("SUIZOWIN",UPPER(RUTA2))-1)+RUTA5+"\SuiAlma"
        ENDCASE
    CASE PROGRAMA2=7                                //Suizo Remesa
        IF FILE(RUTA2+"\cuentas.dbf") .AND. ;
            FILE(RUTA2+"\cuentas.cdx") .AND. ;
            AT("REMESA",UPPER(RUTA2))<>0
            RUTA4:=RUTA2
        ENDIF
    CASE PROGRAMA2=8 .OR. PROGRAMA2=9               //Suizo Escola
        DO CASE
        CASE AT("ESCOLA",UPPER(RUTA2))<>0
            DO CASE
            CASE FILE(RUTA2+"\Empresa.dbf") .AND. ;
                FILE(RUTA2+"\Empresa.cdx")
                RUTA4:=RUTA2
            CASE FILE(LEFT(RUTA2,AT("ESCOLA",UPPER(RUTA2))+5)+"\Empresa.dbf") .AND. ;
                FILE(LEFT(RUTA2,AT("ESCOLA",UPPER(RUTA2))+5)+"\Empresa.cdx")
                RUTA4:=LEFT(RUTA2,AT("ESCOLA",UPPER(RUTA2))+5)
            ENDCASE
        CASE AT("SUIZOWIN",UPPER(RUTA2))<>0
            RUTA5:=SUBSTR(RUTA2,AT("SUIZOWIN",UPPER(RUTA2)),LEN(RUTA2))
            RUTA5:=LEFT(RUTA5,AT("\",RUTA5)-1)
            RUTA4:=LEFT(RUTA2,AT("SUIZOWIN",UPPER(RUTA2))-1)+RUTA5+"\escola"
        ENDCASE
    ENDCASE

    IF LEN(RTRIM(RUTA4))>0
        IF NIVEL2=1
            WinCueAct.T_Ruta1.Value:=RUTA4
        ELSE
            WinCueAct.T_Ruta2.Value:=RUTA4
        ENDIF
    ENDIF
Return Nil

STATIC FUNCTION CuentasActEmp(NIVEL2,RUTA2)
    IF WinCueAct.Calcular.Value=.F.
        RETURN Nil
    ENDIF
    IF NIVEL2=1
        PROGRAMA2:=WinCueAct.C_Programa1.Value
        WinCueAct.T_Registro1.Value:=0
    ELSE
        PROGRAMA2:=WinCueAct.C_Programa2.Value
        WinCueAct.T_Registro2.Value:=0
    ENDIF

    aEMP:={}
    aRUT:={}
    DO CASE
    CASE PROGRAMA2=1 .OR. PROGRAMA2=2 .OR. PROGRAMA2=3
        IF FILE(RUTA2+"\Empresa.dbf") .AND. ;
            FILE(RUTA2+"\Empresa.cdx")
            AbrirDBF("Empresa",,,,RUTA2)
            GO TOP
            DO WHILE .NOT. EOF()
                DO CASE
                CASE FIELDPOS("COD")<>0 .AND. FIELDPOS("NOMBRE")<>0 .AND. FIELDPOS("EJERCICIO")<>0
                    AADD(aEMP,COD+"-"+RTRIM(NOMBRE)+" ("+ALLTRIM(EJERCICIO)+")")
                    AADD(aRUT,COD)
                CASE FIELDPOS("CODEMP")<>0 .AND. FIELDPOS("CNOMBRE")<>0
                    AADD(aEMP,CODEMP+"-"+RTRIM(CNOMBRE))
                    AADD(aRUT,CODEMP)
                ENDCASE
                SKIP
            ENDDO
            DBCLOSEAREA()
        ENDIF
    CASE PROGRAMA2=4
        IF FILE(RUTA2+"\Empresa.dbf") .AND. ;
            FILE(RUTA2+"\Empresa.cdx")
            AbrirDBF("Empresa",,,,RUTA2)
            GO TOP
            DO WHILE .NOT. EOF()
                AADD(aEMP,STR(NEMP)+"-"+RTRIM(EMP)+" ("+STR(EJERCICIO,4)+")")
                AADD(aRUT,"\"+RTRIM(RUTA))
                SKIP
            ENDDO
            DBCLOSEAREA()
        ENDIF
    CASE PROGRAMA2=5 .OR. PROGRAMA2=6
        IF FILE(RUTA2+"\Empcon.dbf") .AND. ;
            FILE(RUTA2+"\Empcon.cdx")
            AbrirDBF("Empcon",,,,RUTA2)
            GO TOP
            DO WHILE .NOT. EOF()
                AADD(aEMP,STR(NEMP)+"-"+RTRIM(EMP)+" ("+STR(EJERCICIO,4)+")")
                AADD(aRUT,"\"+RTRIM(RUTA))
                SKIP
            ENDDO
            DBCLOSEAREA()
        ENDIF
    CASE PROGRAMA2=7
        AADD(aEMP,"Subcuentas Remesa")
        AADD(aRUT,"")
    CASE PROGRAMA2=8 .OR. PROGRAMA2=9
        IF FILE(RUTA2+"\Empresa.dbf") .AND. ;
            FILE(RUTA2+"\Empresa.cdx")
            AbrirDBF("Empresa",,,,RUTA2)
            GO TOP
            DO WHILE .NOT. EOF()
                AADD(aEMP,LTRIM(STR(CODEMP))+"-"+RTRIM(NOMEMP)+" ("+STR(EJERCICIO,4)+")")
                AADD(aRUT,"\"+RTRIM(RUTA))
                SKIP
            ENDDO
            DBCLOSEAREA()
        ENDIF
    ENDCASE

    IF NIVEL2=1
        WinCueAct.C_Empresa1.DeleteAllItems
        WinCueAct.C_RutaEmp1.DeleteAllItems
    ELSE
        WinCueAct.C_Empresa2.DeleteAllItems
        WinCueAct.C_RutaEmp2.DeleteAllItems
    ENDIF
    IF LEN(aEMP)>0
        IF NIVEL2=1
            FOR N=1 TO LEN(aEMP)
                WinCueAct.C_Empresa1.AddItem(aEMP[N])
                WinCueAct.C_RutaEmp1.AddItem(aRUT[N])
            NEXT
            IF WinCueAct.C_Empresa1.ItemCount>0
                WinCueAct.C_Empresa1.Value:=WinCueAct.C_Empresa1.ItemCount
                WinCueAct.C_RutaEmp1.Value:=WinCueAct.C_Empresa1.ItemCount
            ENDIF
        ELSE
            FOR N=1 TO LEN(aEMP)
                WinCueAct.C_Empresa2.AddItem(aEMP[N])
                WinCueAct.C_RutaEmp2.AddItem(aRUT[N])
            NEXT
            IF WinCueAct.C_Empresa2.ItemCount>0
                WinCueAct.C_Empresa2.Value:=WinCueAct.C_Empresa2.ItemCount
                WinCueAct.C_RutaEmp2.Value:=WinCueAct.C_Empresa2.ItemCount
            ENDIF
        ENDIF
    ENDIF

Return Nil

STATIC FUNCTION CuentasActComprobar(LLAMADA,CODIGOACT)
    DO EVENTS
    VerBotones("COMPROBANDO")

    IF LLAMADA="INICIO"
        WinCueAct.T_Registro1.Value:=0
    ENDIF
    RUTA1:=WinCueAct.T_Ruta1.Value+WinCueAct.C_RutaEmp1.DisplayValue
    RUTA2:=WinCueAct.T_Ruta2.Value+WinCueAct.C_RutaEmp2.DisplayValue
    aDATOS1:={}
    DO CASE
    CASE WinCueAct.C_Programa1.Value=1
        AbrirDBF("SUBCTA",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF VAL(COD)=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,VAL(COD))
        AADD(aDATOS1,RTRIM(TITULO))
        AADD(aDATOS1,RTRIM(DOMICILIO))
        AADD(aDATOS1,RTRIM(POBLACION))
        AADD(aDATOS1,RTRIM(PROVINCIA))
        AADD(aDATOS1,RTRIM(CODPOSTAL))
        AADD(aDATOS1,RTRIM(NIF))
        AADD(aDATOS1,RTRIM(TELEF01))
        AADD(aDATOS1,RTRIM(FAX01))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,"")
    CASE WinCueAct.C_Programa1.Value=2
        AbrirDBF("CLIENTES",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF VAL(CSUBCTA)=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,IF(VAL(CSUBCTA)<>0,VAL(CSUBCTA),VAL(CCODCLI)+43000000))
        AADD(aDATOS1,RTRIM(CNOMCLI))
        AADD(aDATOS1,RTRIM(CDIRCLI))
        AADD(aDATOS1,RTRIM(CPOBCLI))
        AADD(aDATOS1,RTRIM(CP_PROV(CPTLCLI)))
        AADD(aDATOS1,RTRIM(CPTLCLI))
        AADD(aDATOS1,RTRIM(CDNICIF))
        AADD(aDATOS1,RTRIM(CTFO1CLI))
        AADD(aDATOS1,RTRIM(CFAXCLI))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN(CENTIDAD+"-"+CAGENCIA+"-  -"+CCUENTA,2))
    CASE WinCueAct.C_Programa1.Value=3
        AbrirDBF("PROVEEDO",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF VAL(CSUBCTA)=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,IF(VAL(CSUBCTA)<>0,VAL(CSUBCTA),VAL(CCODPRO)+40000000))
        AADD(aDATOS1,RTRIM(CNOMPRO))
        AADD(aDATOS1,RTRIM(CDIRPRO))
        AADD(aDATOS1,RTRIM(CPOBPRO))
        AADD(aDATOS1,RTRIM(CP_PROV(CPTLPRO)))
        AADD(aDATOS1,RTRIM(CPTLPRO))
        AADD(aDATOS1,RTRIM(CNIFDNI))
        AADD(aDATOS1,RTRIM(CTFO1PRO))
        AADD(aDATOS1,RTRIM(CFAX))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN(CENTIDAD+"-"+CAGENCIA+"-  -"+CCUENTA,2))
    CASE WinCueAct.C_Programa1.Value=4
        AbrirDBF("CUENTAS",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF CODCTA=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,CODCTA)
        AADD(aDATOS1,RTRIM(NOMCTA))
        AADD(aDATOS1,RTRIM(DIRCTA))
        AADD(aDATOS1,RTRIM(POBCTA))
        AADD(aDATOS1,RTRIM(PROCTA))
        AADD(aDATOS1,RTRIM(CODPOS))
        AADD(aDATOS1,RTRIM(CIF))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(FAX1))
        AADD(aDATOS1,"")
        AADD(aDATOS1,CTA_BAN_SUIZO(BANCTA,2))
    CASE WinCueAct.C_Programa1.Value=5
        AbrirDBF("CLIENTES",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF SUBCTA=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,IF(SUBCTA<>0,SUBCTA,COD+43000000))
        AADD(aDATOS1,RTRIM(CLIENTE))
        AADD(aDATOS1,RTRIM(DIRECCION))
        AADD(aDATOS1,RTRIM(POBLACION))
        AADD(aDATOS1,RTRIM(PROVINCIA))
        AADD(aDATOS1,RTRIM(CODPOS))
        AADD(aDATOS1,RTRIM(CIF))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(FAX))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN_SUIZO(BANCTA,2))
    CASE WinCueAct.C_Programa1.Value=6
        AbrirDBF("PROVEE",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF SUBCTA=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,IF(SUBCTA<>0,SUBCTA,COD+40000000))
        AADD(aDATOS1,RTRIM(PROVEE))
        AADD(aDATOS1,RTRIM(DIRECCION))
        AADD(aDATOS1,RTRIM(POBLACION))
        AADD(aDATOS1,RTRIM(PROVINCIA))
        AADD(aDATOS1,RTRIM(CODPOS))
        AADD(aDATOS1,RTRIM(CIF))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(FAX1))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN_SUIZO(BANCTA,2))
    CASE WinCueAct.C_Programa1.Value=7
        AbrirDBF("CUENTAS",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF CODCTA=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,CODCTA)
        AADD(aDATOS1,HB_OEMtoANSI(RTRIM(NOMCTA)))
        AADD(aDATOS1,HB_OEMtoANSI(RTRIM(DIRCTA)))
        AADD(aDATOS1,HB_OEMtoANSI(RTRIM(POBCTA)))
        AADD(aDATOS1,HB_OEMtoANSI(RTRIM(SUBSTR(PROCTA,7,30))))
        AADD(aDATOS1,RTRIM(LEFT(PROCTA,5)))
        AADD(aDATOS1,RTRIM(CIF))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(FAX1))
        AADD(aDATOS1,"")
        AADD(aDATOS1,CTA_BAN_SUIZO(BANCTA,2))
    CASE WinCueAct.C_Programa1.Value=8
        AbrirDBF("ALUMNOS",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF CODALU+43000000=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        AADD(aDATOS1,CODALU+43000000)
        AADD(aDATOS1,RTRIM(NOMALU))
        AADD(aDATOS1,RTRIM(DIRECCION))
        AADD(aDATOS1,RTRIM(POBLACION))
        AADD(aDATOS1,RTRIM(PROVINCIA))
        AADD(aDATOS1,RTRIM(CODPOS))
        AADD(aDATOS1,RTRIM(IF(FRAPADRE<=1,NIFPADRE,NIFMADRE)))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(TEL2))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN_SUIZO(CTABAN,2))
    CASE WinCueAct.C_Programa1.Value=9
        AbrirDBF("EDUCADORAS",,"BASE1",,RUTA1)
        IF WinCueAct.T_Registro1.Value=0
            WinCueAct.Progress_1.RangeMin:=1
            WinCueAct.Progress_1.RangeMax:=LASTREC()
            WinCueAct.Progress_1.Value:=1
            GO TOP
            WinCueAct.T_Registro1.Value:=RECNO()
            IF !CODIGOACT=NIL
                DO WHILE .NOT. EOF()
                    DO EVENTS
                    IF CODCTA=CODIGOACT
                        WinCueAct.T_Registro1.Value:=RECNO()
                        EXIT
                    ENDIF
                    WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
                    SKIP
                ENDDO
            ENDIF
        ELSE
            GO WinCueAct.T_Registro1.Value
            SKIP 1
            WinCueAct.T_Registro1.Value:=RECNO()
        ENDIF
        WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        DO WHILE CODCTA=0 .AND. .NOT. EOF()
            SKIP
            WinCueAct.Progress_1.Value:=WinCueAct.Progress_1.Value+1
        ENDDO
        AADD(aDATOS1,CODCTA)
        AADD(aDATOS1,RTRIM(NOMEDU))
        AADD(aDATOS1,RTRIM(DIRECCION))
        AADD(aDATOS1,RTRIM(POBLACION))
        AADD(aDATOS1,RTRIM(PROVINCIA))
        AADD(aDATOS1,RTRIM(CODPOS))
        AADD(aDATOS1,RTRIM(NIFEDU))
        AADD(aDATOS1,RTRIM(TEL1))
        AADD(aDATOS1,RTRIM(TEL2))
        AADD(aDATOS1,RTRIM(EMAIL))
        AADD(aDATOS1,CTA_BAN_SUIZO(CTABAN,2))
    ENDCASE

    IF LEN(aDATOS1)=11
        WinCueAct.T_DatoA1.Value:=aDATOS1[1]
        WinCueAct.T_DatoA2.Value:=aDATOS1[2]
        WinCueAct.T_DatoA3.Value:=aDATOS1[3]
        WinCueAct.T_DatoA4.Value:=aDATOS1[4]
        WinCueAct.T_DatoA5.Value:=aDATOS1[5]
        WinCueAct.T_DatoA6.Value:=aDATOS1[6]
        WinCueAct.T_DatoA7.Value:=aDATOS1[7]
        WinCueAct.T_DatoA8.Value:=aDATOS1[8]
        WinCueAct.T_DatoA9.Value:=aDATOS1[9]
        WinCueAct.T_DatoA10.Value:=aDATOS1[10]
        WinCueAct.T_DatoA11.Value:=aDATOS1[11]
    ENDIF

    IF .NOT. EOF()
        WinCueAct.T_Registro1.Value:=RECNO()
        WinCueAct.T_Registro1t.Value:=LASTREC()
    ELSE
        MSGBOX("Actualizacion de ficheros terminada")
        VerBotones("INICIO")
        RETURN
    ENDIF


    aDATOS2:={}
    DO CASE
    CASE WinCueAct.C_Programa2.Value=1
        AbrirDBF("SUBCTA",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK PADR(LTRIM(STR(aDATOS1[1])),12," ")
        IF .NOT. EOF()
            AADD(aDATOS2,VAL(COD))
            AADD(aDATOS2,RTRIM(TITULO))
            AADD(aDATOS2,RTRIM(DOMICILIO))
            AADD(aDATOS2,RTRIM(POBLACION))
            AADD(aDATOS2,RTRIM(PROVINCIA))
            AADD(aDATOS2,RTRIM(CODPOSTAL))
            AADD(aDATOS2,RTRIM(NIF))
            AADD(aDATOS2,RTRIM(TELEF01))
            AADD(aDATOS2,RTRIM(FAX01))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,"")
        ENDIF
    CASE WinCueAct.C_Programa2.Value=2
        AbrirDBF("CLIENTES",,"BASE2",,RUTA2)
        DBSETORDER(5)
        SEEK PADR(LTRIM(STR(aDATOS1[1])),12," ")
        IF .NOT. EOF()
            AADD(aDATOS2,IF(VAL(CSUBCTA)<>0,VAL(CSUBCTA),VAL(CCODCLI)+43000000))
            AADD(aDATOS2,RTRIM(CNOMCLI))
            AADD(aDATOS2,RTRIM(CDIRCLI))
            AADD(aDATOS2,RTRIM(CPOBCLI))
            AADD(aDATOS2,RTRIM(CP_PROV(CPTLCLI)))
            AADD(aDATOS2,RTRIM(CPTLCLI))
            AADD(aDATOS2,RTRIM(CDNICIF))
            AADD(aDATOS2,RTRIM(CTFO1CLI))
            AADD(aDATOS2,RTRIM(CFAXCLI))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN(CENTIDAD+"-"+CAGENCIA+"-  -"+CCUENTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=3
        AbrirDBF("PROVEEDO",,"BASE2",,RUTA2)
        DBSETORDER(4)
        SEEK PADR(LTRIM(STR(aDATOS1[1])),12," ")
        IF .NOT. EOF()
            AADD(aDATOS2,IF(VAL(CSUBCTA)<>0,VAL(CSUBCTA),VAL(CCODPRO)+40000000))
            AADD(aDATOS2,RTRIM(CNOMPRO))
            AADD(aDATOS2,RTRIM(CDIRPRO))
            AADD(aDATOS2,RTRIM(CPOBPRO))
            AADD(aDATOS2,RTRIM(CP_PROV(CPTLPRO)))
            AADD(aDATOS2,RTRIM(CPTLPRO))
            AADD(aDATOS2,RTRIM(CNIFDNI))
            AADD(aDATOS2,RTRIM(CTFO1PRO))
            AADD(aDATOS2,RTRIM(CFAX))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN(CENTIDAD+"-"+CAGENCIA+"-  -"+CCUENTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=4
        AbrirDBF("CUENTAS",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK aDATOS1[1]
        IF .NOT. EOF()
            AADD(aDATOS2,CODCTA)
            AADD(aDATOS2,RTRIM(NOMCTA))
            AADD(aDATOS2,RTRIM(DIRCTA))
            AADD(aDATOS2,RTRIM(POBCTA))
            AADD(aDATOS2,RTRIM(PROCTA))
            AADD(aDATOS2,RTRIM(CODPOS))
            AADD(aDATOS2,RTRIM(CIF))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(FAX1))
            AADD(aDATOS2,"")
            AADD(aDATOS2,CTA_BAN_SUIZO(BANCTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=5
        AbrirDBF("CLIENTES",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK aDATOS1[1]-43000000
        IF EOF()
            LOCATE FOR SUBCTA=aDATOS1[1]
        ENDIF
        IF .NOT. EOF()
            AADD(aDATOS2,IF(SUBCTA<>0,SUBCTA,COD+43000000))
            AADD(aDATOS2,RTRIM(CLIENTE))
            AADD(aDATOS2,RTRIM(DIRECCION))
            AADD(aDATOS2,RTRIM(POBLACION))
            AADD(aDATOS2,RTRIM(PROVINCIA))
            AADD(aDATOS2,RTRIM(CODPOS))
            AADD(aDATOS2,RTRIM(CIF))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(FAX))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN_SUIZO(BANCTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=6
        AbrirDBF("PROVEE",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK aDATOS1[1]-40000000
        IF EOF()
            LOCATE FOR SUBCTA=aDATOS1[1]
        ENDIF
        IF .NOT. EOF()
            AADD(aDATOS2,IF(SUBCTA<>0,SUBCTA,COD+40000000))
            AADD(aDATOS2,RTRIM(PROVEE))
            AADD(aDATOS2,RTRIM(DIRECCION))
            AADD(aDATOS2,RTRIM(POBLACION))
            AADD(aDATOS2,RTRIM(PROVINCIA))
            AADD(aDATOS2,RTRIM(CODPOS))
            AADD(aDATOS2,RTRIM(CIF))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(FAX1))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN_SUIZO(BANCTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=7
        AbrirDBF("CUENTAS",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK aDATOS1[1]
        IF .NOT. EOF()
            AADD(aDATOS2,CODCTA)
            AADD(aDATOS2,HB_OEMtoANSI(RTRIM(NOMCTA)))
            AADD(aDATOS2,HB_OEMtoANSI(RTRIM(DIRCTA)))
            AADD(aDATOS2,HB_OEMtoANSI(RTRIM(POBCTA)))
            AADD(aDATOS2,HB_OEMtoANSI(RTRIM(SUBSTR(PROCTA,7,30))))
            AADD(aDATOS2,RTRIM(LEFT(PROCTA,5)))
            AADD(aDATOS2,RTRIM(CIF))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(FAX1))
            AADD(aDATOS2,"")
            AADD(aDATOS2,CTA_BAN_SUIZO(BANCTA,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=8
        AbrirDBF("ALUMNOS",,"BASE2",,RUTA2)
        DBSETORDER(1)
        SEEK aDATOS1[1]-43000000
        IF .NOT. EOF()
            AADD(aDATOS2,CODALU+43000000)
            AADD(aDATOS2,RTRIM(NOMALU))
            AADD(aDATOS2,RTRIM(DIRECCION))
            AADD(aDATOS2,RTRIM(POBLACION))
            AADD(aDATOS2,RTRIM(PROVINCIA))
            AADD(aDATOS2,RTRIM(CODPOS))
            AADD(aDATOS2,RTRIM(IF(FRAPADRE<=1,NIFPADRE,NIFMADRE)))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(TEL2))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN_SUIZO(CTABAN,2))
        ENDIF
    CASE WinCueAct.C_Programa2.Value=9
        AbrirDBF("EDUCADORAS",,"BASE2",,RUTA2)
        LOCATE FOR CODCTA=aDATOS1[1]
        IF .NOT. EOF()
            AADD(aDATOS2,CODCTA)
            AADD(aDATOS2,RTRIM(NOMEDU))
            AADD(aDATOS2,RTRIM(DIRECCION))
            AADD(aDATOS2,RTRIM(POBLACION))
            AADD(aDATOS2,RTRIM(PROVINCIA))
            AADD(aDATOS2,RTRIM(CODPOS))
            AADD(aDATOS2,RTRIM(NIFEDU))
            AADD(aDATOS2,RTRIM(TEL1))
            AADD(aDATOS2,RTRIM(TEL2))
            AADD(aDATOS2,RTRIM(EMAIL))
            AADD(aDATOS2,CTA_BAN_SUIZO(CTABAN,2))
        ENDIF
    ENDCASE

    IF LEN(aDATOS2)=0
        aDATOS2:={0,"","","","","","","","","",""}
    ENDIF
    WinCueAct.T_DatoB1.Value:=aDATOS2[1]
    WinCueAct.T_DatoB2.Value:=aDATOS2[2]
    WinCueAct.T_DatoB3.Value:=aDATOS2[3]
    WinCueAct.T_DatoB4.Value:=aDATOS2[4]
    WinCueAct.T_DatoB5.Value:=aDATOS2[5]
    WinCueAct.T_DatoB6.Value:=aDATOS2[6]
    WinCueAct.T_DatoB7.Value:=aDATOS2[7]
    WinCueAct.T_DatoB8.Value:=aDATOS2[8]
    WinCueAct.T_DatoB9.Value:=aDATOS2[9]
    WinCueAct.T_DatoB10.Value:=aDATOS2[10]
    WinCueAct.T_DatoB11.Value:=aDATOS2[11]

    WinCueAct.T_Registro2t.Value:=LASTREC()
    IF .NOT. EOF()
        WinCueAct.T_Registro2.Value:=RECNO()
    ELSE
        WinCueAct.T_Registro2.Value:=0
    ENDIF


    MAL:=0
    FOR N=1 TO 11
        DO CASE
        CASE aDATOS1[N]==aDATOS2[N]
            SetProperty("WinCueAct","C_Si"+LTRIM(STR(N)),"Value",.F.)
            SetProperty("WinCueAct","T_DatoB"+LTRIM(STR(N)),"BACKCOLOR",MICOLOR("BLANCO"))
        CASE N=11 .AND. (WinCueAct.C_Programa1.Value=1 .OR. WinCueAct.C_Programa2.Value=1) .OR. ;
            N=10 .AND. (WinCueAct.C_Programa1.Value=4 .OR. WinCueAct.C_Programa2.Value=4) .OR. ;
            N=10 .AND. (WinCueAct.C_Programa1.Value=7 .OR. WinCueAct.C_Programa2.Value=7)
            SetProperty("WinCueAct","C_Si"+LTRIM(STR(N)),"Value",.F.)
            SetProperty("WinCueAct","T_DatoB"+LTRIM(STR(N)),"BACKCOLOR",MICOLOR("BLANCO"))
        OTHERWISE
            SetProperty("WinCueAct","C_Si"+LTRIM(STR(N)),"Value",.T.)
            SetProperty("WinCueAct","T_DatoB"+LTRIM(STR(N)),"BACKCOLOR",MICOLOR("ROJOCLARO"))
            MAL:=1
        ENDCASE
    NEXT

    SELEC BASE1
    IF MAL=0 .AND. .NOT. EOF()
        CuentasActComprobar("COMPROBAR")
    ELSE
        IF WinCueAct.T_DatoB1.Value=0
            WinCueAct.C_Si1.Value:=.T.
            WinCueAct.Bt_Actualizar.Caption:="Alta"
            VerBotones("ALTA")
        ELSE
            WinCueAct.Bt_Actualizar.Caption:="Actualizar"
            VerBotones("ACTUALIZAR")
        ENDIF
    ENDIF

Return Nil



STATIC FUNCTION CuentasActFic(LLAMADA)
    DO EVENTS
***NO SE LOCALIZAN LAS BASES***
    IF SELEC("BASE1")=0 .OR. SELEC("BASE2")=0
        MSGSTOP("No se han aperturado las bases"+HB_OsNewLine()+ ;
            "Pulse el boton -Comprobar-")
        RETURN
    ENDIF
***FIN NO SE LOCALIZAN LAS BASES***

    IF LLAMADA="ALREVES"
        SELEC BASE1
        IF WinCueAct.T_Registro1.Value=0
            APPEND BLANK
        ELSE
            GO WinCueAct.T_Registro1.Value
        ENDIF
        PROGRAMA2:=WinCueAct.C_Programa1.Value
        LETRA1:="A"
        LETRA2:="B"
    ELSE
        SELEC BASE2
        IF WinCueAct.T_Registro2.Value=0
            IF WinCueAct.C_Programa1.Value=WinCueAct.C_Programa2.Value
                Duplicar_RegistroDBF(SELEC("BASE1"),SELEC("BASE2"))
            ELSE
                APPEND BLANK
            ENDIF
        ELSE
            GO WinCueAct.T_Registro2.Value
        ENDIF
        PROGRAMA2:=WinCueAct.C_Programa2.Value
        LETRA1:="B"
        LETRA2:="A"
    ENDIF


    DO CASE
    CASE PROGRAMA2=1
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE COD WITH STR(GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value"))
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE TITULO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DOMICILIO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBLACION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROVINCIA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOSTAL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE NIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TELEF01 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE FAX01 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=2
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CSUBCTA WITH STR(GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value"))
                REPLACE CCODCLI WITH STR(GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-43000000)
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE CNOMCLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE CDIRCLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE CPOBCLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE CCODPROV WITH "00"+LEFT(GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value"),2)
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CPTLCLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CDNICIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE CTFO1CLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE CFAXCLI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE CENTIDAD WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),1,4)
                REPLACE CAGENCIA WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),6,4)
                REPLACE CCUENTA  WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),14,10)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=3
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CSUBCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
                REPLACE CCODPRO WITH STR(GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-40000000)
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE CNOMPRO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE CDIRPRO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE CPOBPRO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE CCODPROV WITH "00"+LEFT(GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value"),2)
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CPTLPRO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CNIFDNI WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE CTFO1PRO WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE CFAX WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE CENTIDAD WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),1,4)
                REPLACE CAGENCIA WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),6,4)
                REPLACE CCUENTA  WITH SUBSTR(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),14,10)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=4
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CODCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE NOMCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOS WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE FAX1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
*         REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE BANCTA WITH CTA_BAN_SUIZO(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),2)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=5
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE SUBCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
                REPLACE COD    WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-43000000
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE CLIENTE WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRECCION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBLACION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROVINCIA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOS WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE FAX WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE BANCTA WITH CTA_BAN_SUIZO(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),3)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=6
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE SUBCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
                REPLACE COD    WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-40000000
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE PROVEE WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRECCION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBLACION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROVINCIA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOS WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE FAX1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE BANCTA WITH CTA_BAN_SUIZO(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),3)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=7
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CODCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE NOMCTA WITH HB_ANSItoOEM(GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value"))
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRCTA WITH HB_ANSItoOEM(GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value"))
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBCTA WITH HB_ANSItoOEM(GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value"))
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROCTA WITH LEFT(PROCTA,6)+HB_ANSItoOEM(GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value"))
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE PROCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")+"-"+SUBSTR(PROCTA,7,30)
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE CIF WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE FAX1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
*         REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE BANCTA WITH CTA_BAN_SUIZO(GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value"),3)
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=8
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CODALU WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-43000000
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE NOMALU WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRECCION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBLACION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROVINCIA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOS WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                IF FRAPADRE<=1
                    REPLACE NIFPADRE WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
                ELSE
                    REPLACE NIFMADRE WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
                ENDIF
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE TEL2 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE CTABAN WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value")
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    CASE PROGRAMA2=9
        IF RLOCK()
            IF WinCueAct.C_Si1.Value=.T.
                REPLACE CODCTA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")
                REPLACE CODEDU WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"1","Value")-46500000
            ENDIF
            IF WinCueAct.C_Si2.Value=.T.
                REPLACE NOMEDU WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"2","Value")
            ENDIF
            IF WinCueAct.C_Si3.Value=.T.
                REPLACE DIRECCION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"3","Value")
            ENDIF
            IF WinCueAct.C_Si4.Value=.T.
                REPLACE POBLACION WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"4","Value")
            ENDIF
            IF WinCueAct.C_Si5.Value=.T.
                REPLACE PROVINCIA WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"5","Value")
            ENDIF
            IF WinCueAct.C_Si6.Value=.T.
                REPLACE CODPOS WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"6","Value")
            ENDIF
            IF WinCueAct.C_Si7.Value=.T.
                REPLACE NIFEDU WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"7","Value")
            ENDIF
            IF WinCueAct.C_Si8.Value=.T.
                REPLACE TEL1 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"8","Value")
            ENDIF
            IF WinCueAct.C_Si9.Value=.T.
                REPLACE TEL2 WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"9","Value")
            ENDIF
            IF WinCueAct.C_Si10.Value=.T.
                REPLACE EMAIL WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"10","Value")
            ENDIF
            IF WinCueAct.C_Si11.Value=.T.
                REPLACE CTABAN WITH GetProperty("WinCueAct","T_Dato"+LETRA2+"11","Value")
            ENDIF
            DBCOMMIT()
            DBUNLOCK()
            MSGBOX("Programa: "+aPROG[PROGRAMA2]+HB_OsNewLine()+ ;
                "Registro: "+LTRIM(STR(RECNO())) ,"Registro actualizado")
        ENDIF
    ENDCASE

    CuentasActComprobar("COMPROBAR")
    Return Nil