#include "minigui.ch"
#include "winprint.ch"

procedure ResVenta(LLAMADA)

    IF PROGRAMA()="SUICONTA"
        aTIPO:={"Perdidas y ganancias"}
    ELSE
        aTIPO:={"Pedidos proveedor", ;
            "Albaranes proveedor", ;
            "Ofertas cliente", ;
            "Pedidos cliente", ;
            "Albaranes cliente", ;
            "Facturas cliente", ;
            "Tickets cliente"}
    ENDIF

    DO CASE
    CASE UPPER(LLAMADA)="PYG"
        nTIPO:=1
    CASE UPPER(LLAMADA)="PROPED"
        nTIPO:=1
    CASE UPPER(LLAMADA)="PROALB"
        nTIPO:=2
    CASE UPPER(LLAMADA)="CLIOFE"
        nTIPO:=3
    CASE UPPER(LLAMADA)="CLIPED"
        nTIPO:=4
    CASE UPPER(LLAMADA)="CLIALB"
        nTIPO:=5
    CASE UPPER(LLAMADA)="CLIFAC"
        nTIPO:=6
    CASE UPPER(LLAMADA)="CLITIC"
        nTIPO:=7
    OTHERWISE
        nTIPO:=6
    ENDCASE

    TituloImp:="Resumen de "+aTIPO[nTIPO]

    IF IsWindowDefined(W_ResVenta)=.T.
        DECLARE WINDOW W_ResVenta
        W_ResVenta.release
        DO EVENTS
    ENDIF

    DEFINE WINDOW W_ResVenta AT 10,10 WIDTH 800 HEIGHT 330 ;
        TITLE TituloImp ;
        CHILD NOMAXIMIZE NOSIZE ;
        ON RELEASE CloseTables()

        @010,010 GRID GR_Resumen ;
            HEIGHT 270 ;
            WIDTH 350 ;
            HEADERS {'Mes'} ;
            WIDTHS {100} ;
            ITEMS {} ;
            COLUMNCONTROLS {{'TEXTBOX','CHARACTER'}} ;
            JUSTIFY {GRID_JTFY_RIGHT}

        Menu_Grid("W_ResVenta","GR_Resumen","MENU",,{"COPREGPP","COPTABPP"})

        @ 15,410 LABEL L_Tipo VALUE 'Tipo resumen' AUTOSIZE TRANSPARENT
        @ 10,510 COMBOBOX C_Tipo WIDTH 200 ITEMS aTIPO VALUE nTIPO ;
            ON CHANGE ResVenta_Tipo() NOTABSTOP

        @ 45,410 LABEL L_Ejer1 VALUE 'Desde el ejercicio' AUTOSIZE TRANSPARENT
        @ 40,510 SPINNER Ejer1 RANGE EJERANY-100,EJERANY+100 VALUE EJERANY WIDTH 70

        @ 75,410 LABEL L_Ejer2 VALUE 'Hasta el ejercicio' AUTOSIZE TRANSPARENT
        @ 70,510 SPINNER Ejer2 RANGE EJERANY-100,EJERANY+100 VALUE EJERANY WIDTH 70

        @ 40,600 RADIOGROUP R_Acumulado OPTIONS { 'Mensual' , 'Acumulado' } VALUE 1 WIDTH 80

        @ 40,700 RADIOGROUP R_Ventas OPTIONS { 'Ventas' , 'Margen' } VALUE 1 WIDTH 80 

        @105,410 LABEL L_Serie VALUE 'Serie' AUTOSIZE TRANSPARENT
        @100,510 TEXTBOX T_Serie WIDTH 20 VALUE ''
        IF PROGRAMA()="SUICONTA"
            W_ResVenta.L_Serie.Visible:=.F.
            W_ResVenta.T_Serie.Visible:=.F.
        ENDIF

        @140,410 BUTTONEX B_Calcular CAPTION 'Calcular' ICON icobus('Buscar') WIDTH 90 HEIGHT 25 ;
            ACTION ResVentaBUS()
        @140,510 PROGRESSBAR P_Progres RANGE 0,100 WIDTH 200 HEIGHT 20 SMOOTH
        @140,715 LABEL L_VerAno VALUE '' AUTOSIZE TRANSPARENT

        LINW:=170
        COLW:=400
        draw rectangle in window W_ResVenta at LINW,COLW+10 to LINW+2,COLW+390 fillcolor{255,0,0}  //Rojo
        aIMP:=Impresoras(EMP_IMPRESORA)
        @LINW+15,COLW+10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @LINW+10,COLW+100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @LINW+45,COLW+220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @LINW+40,COLW+280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @LINW+40,COLW+10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_ResVenta.C_Impresora.Enabled:=IF(W_ResVenta.nImp.Value=.T.,.F.,.T.)

        @LINW+70,COLW+10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @LINW+100,COLW+010 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 90 HEIGHT 25 ;
            ACTION ( ResVentaBUS() , ResVentaCalc() )

        @LINW+100,COLW+110 BUTTONEX B_Grafico CAPTION 'Grafico' ICON icobus('grafico') WIDTH 90 HEIGHT 25 ;
            ACTION ( ResVentaBUS() , ResVentaGraf() )

        @LINW+100,COLW+210 BUTTONEX B_Imprimir CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION ( ResVentaBUS() , ResVentaImp() )

        @LINW+100,COLW+310 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('Salir') WIDTH 80 HEIGHT 25 ;
            ACTION W_ResVenta.release

    END WINDOW
    VentanaCentrar("W_ResVenta","Ventana1")
    CENTER WINDOW W_ResVenta
    ACTIVATE WINDOW W_ResVenta

Return Nil


STATIC FUNCTION ResVenta_Tipo()
    W_ResVenta.Title:="Resumen de "+W_ResVenta.C_Tipo.DisplayValue
    IF W_ResVenta.C_Tipo.Value=4 .OR. ;
        W_ResVenta.C_Tipo.Value=5 .OR. ;
        W_ResVenta.C_Tipo.Value=6 .OR. ;
        W_ResVenta.C_Tipo.Value=7
        W_ResVenta.R_Ventas.Visible:=.T.
    ELSE
        W_ResVenta.R_Ventas.Visible:=.F.
        W_ResVenta.R_Ventas.Value:=1
    ENDIF
Return Nil


STATIC FUNCTION ResVentaBUS()
    TotalCol2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR NCOL=2 TO TotalCol2
        W_ResVenta.GR_Resumen.DeleteColumn(2)
    NEXT
    DO EVENTS

    aRESUMEN:={}
    DO CASE
    CASE W_ResVenta.C_Tipo.Value=1 .AND. PROGRAMA()="SUICONTA"
        ResVentaPyG()
    CASE W_ResVenta.C_Tipo.Value=1
        ResVentaProPed()
    CASE W_ResVenta.C_Tipo.Value=2
        ResVentaProAlb()
    CASE W_ResVenta.C_Tipo.Value=3
        ResVentaCliOfe()
    CASE W_ResVenta.C_Tipo.Value=4
        ResVentaCliPed()
    CASE W_ResVenta.C_Tipo.Value=5
        ResVentaCliAlb()
    CASE W_ResVenta.C_Tipo.Value=6
        ResVentaCliFac()
    CASE W_ResVenta.C_Tipo.Value=7
        ResVentaCliTic()
    ENDCASE

    IF LEN(aRESUMEN)=0
        MsgExclamation("No hay datos en las fechas introducidas","Informacion")
        RETURN
    ENDIF

    ResVentaGrid(aRESUMEN)

Return Nil


STATIC FUNCTION ResVentaPyG()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\APUNTES.DBF")
            AbrirDBF("APUNTES",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                IF DELETED()=.T.
                    SKIP
                    LOOP
                ENDIF
                IF LTRIM(STR(CODCTA))="6" .OR. LTRIM(STR(CODCTA))="7"
                    IF LTRIM(STR(CODCTA))<>"72" .AND. LTRIM(STR(CODCTA))<>"78"
                        aRESUMEN[LEN(aRESUMEN),MONTH(FECHA)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FECHA)]-DEBE+HABER
                    ENDIF
                ENDIF
                SKIP
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil

STATIC FUNCTION ResVentaProPed()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\DEMANDA.DBF")
            AbrirDBF("DEMANDA",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                IF DELETED()=.T. .OR. NDEM<=0
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERDEM
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                SERDEM2:=SERDEM
                NDEM2:=NDEM
                FDEM2:=FDEM
                DESC2:=DTO1
                IMP1:=0
                DO WHILE NDEM=NDEM2 .AND. SERDEM=SERDEM2 .AND. .NOT. EOF()
                    IF DELETED()=.T.
                        W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                        SKIP
                        LOOP
                    ENDIF
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    IMP1:=IMP1+ROUND((UNIDAD*PRECIO),MDA_DEC(EJERANY))
                    SKIP
                ENDDO
                DES1:=ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                BIMP2:=IMP1-DES1
                aRESUMEN[LEN(aRESUMEN),MONTH(FDEM2)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FDEM2)]+BIMP2
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaProAlb()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\ENTRADA.DBF")
            AbrirDBF("ENTRADA",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                IF DELETED()=.T.
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERENT
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                SERENT2:=SERENT
                NENT2:=NENT
                FENT2:=FENT
                DESC2:=DTO1
                IMP1:=0
                DO WHILE NENT=NENT2 .AND. SERENT=SERENT2 .AND. .NOT. EOF()
                    IF DELETED()=.T.
                        W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                        SKIP
                        LOOP
                    ENDIF
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    IMP1:=IMP1+ROUND((UNIDAD*PRECIO),MDA_DEC(EJERANY))
                    SKIP
                ENDDO
                DES1:=ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                BIMP2:=IMP1-DES1
                aRESUMEN[LEN(aRESUMEN),MONTH(FENT2)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FENT2)]+BIMP2
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaCliOfe()
    aRESUMEN:={}
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\OFERTAS.DBF")
            AbrirDBF("OFERTAS",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                DO EVENTS
                IF DELETED()=.T.
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SEROFE
                    SKIP
                    LOOP
                ENDIF
                aRESUMEN[LEN(aRESUMEN),MONTH(FOFE)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FOFE)]+IMPORTE
                SKIP
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaCliPed()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\PEDIDOS.DBF")
            AbrirDBF("PEDIDOS",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                IF DELETED()=.T. .OR. NPED<=0
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERPED
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                SERPED2:=SERPED
                NPED2:=NPED
                FPED2:=FPED
                DESC2:=DESCUENTO
                IMP1:=0
                DO WHILE NPED=NPED2 .AND. SERPED=SERPED2 .AND. .NOT. EOF()
                    IF DELETED()=.T.
                        W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                        SKIP
                        LOOP
                    ENDIF
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    IF W_ResVenta.R_Ventas.Value=1
                        IMP1:=IMP1+ROUND((UNIDAD*PRECIO),MDA_DEC(EJERANY))
                    ELSE
                        IMP1:=IMP1+ROUND((UNIDAD*(PRECIO-PCOSTE)),MDA_DEC(EJERANY))
                    ENDIF
                    SKIP
                ENDDO
                DES1:=ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                BIMP2:=IMP1-DES1
                aRESUMEN[LEN(aRESUMEN),MONTH(FPED2)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FPED2)]+BIMP2
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaCliAlb()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\ALBARAN.DBF")
            AbrirDBF("ALBARAN",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                IF DELETED()=.T.
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERALB
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                SERALB2:=SERALB
                NALB2:=NALB
                FALB2:=FALB
                DESC2:=DESCUENTO
                IMP1:=0
                DO WHILE NALB=NALB2 .AND. SERALB=SERALB2 .AND. .NOT. EOF()
                    IF DELETED()=.T.
                        W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                        SKIP
                        LOOP
                    ENDIF
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    IF W_ResVenta.R_Ventas.Value=1
                        IMP1:=IMP1+ROUND((UNIDAD*PRECIO),MDA_DEC(EJERANY))
                    ELSE
                        IMP1:=IMP1+ROUND((UNIDAD*(PRECIO-PCOSTE)),MDA_DEC(EJERANY))
                    ENDIF
                    SKIP
                ENDDO
                DES1:=ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                BIMP2:=IMP1-DES1
                aRESUMEN[LEN(aRESUMEN),MONTH(FALB2)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FALB2)]+BIMP2
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaCliFac()
    aRESUMEN:={}
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\FACTURAS.DBF")
            AbrirDBF("FACTURAS",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                DO EVENTS
                IF DELETED()=.T.
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERFAC
                    SKIP
                    LOOP
                ENDIF
                IF W_ResVenta.R_Ventas.Value=1
                    BIMP2:=BIMP
                ELSE
                    SERFAC2:=SERFAC
                    NFAC2:=NFAC
                    DESC2:=DTO1
                    IMP1:=0
                    AbrirDBF("ALBARAN",,,,RUTA2,2)
                    SEEK STRZERO(NFAC2,6)+SERFAC2
                    DO WHILE NFAC=NFAC2 .AND. SERFAC=SERFAC2 .AND. .NOT. EOF()
                        IF DELETED()=.T.
                            SKIP
                            LOOP
                        ENDIF
                        IMP1:=IMP1+ROUND((UNIDAD*(PRECIO-PCOSTE)),MDA_DEC(EJERANY))
                        SKIP
                    ENDDO
                    DES1:=ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                    BIMP2:=IMP1-DES1
                    SELEC FACTURAS
                ENDIF
                aRESUMEN[LEN(aRESUMEN),MONTH(FFAC)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FFAC)]+BIMP2
                SKIP
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaCliTic()
*** INCLUIR EJERCICIOS ANTORIORES ***
    COL2:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR ANY2=W_ResVenta.Ejer1.value TO W_ResVenta.Ejer2.value
        W_ResVenta.L_VerAno.Value:=LTRIM(STR(ANY2))
        ADD COLUMN INDEX ++COL2 CAPTION LTRIM(STR(ANY2)) WIDTH 100 JUSTIFY 1 TO GR_Resumen OF W_ResVenta
        DO EVENTS
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        AADD(aRESUMEN,{0,0,0,0,0,0,0,0,0,0,0,0})
        IF FILE(RUTA2+"\TICKET.DBF")
            AbrirDBF("TICKET",,,,RUTA2)
            SET DELETE OFF
            TANTO:=0
            GO TOP
            DO WHILE .NOT. EOF()
                IF DELETED()=.T.
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                IF LEN(RTRIM(W_ResVenta.T_Serie.Value))<>0 .AND. W_ResVenta.T_Serie.Value<>SERTIC
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    SKIP
                    LOOP
                ENDIF
                SERTIC2:=SERTIC
                NTIC2:=NTIC
                FTIC2:=FTIC
***         DESC2:=DESCUENTO
                IMP1:=0
                DO WHILE NTIC=NTIC2 .AND. SERTIC=SERTIC2 .AND. .NOT. EOF()
                    IF DELETED()=.T.
                        W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                        SKIP
                        LOOP
                    ENDIF
                    W_ResVenta.P_Progres.Value:= ++TANTO*100/LASTREC()
                    IF W_ResVenta.R_Ventas.Value=1
                        IMP1:=IMP1+ROUND((UNIDAD*PVENTA),MDA_DEC(EJERANY))
                    ELSE
                        IMP1:=IMP1+ROUND((UNIDAD*(PVENTA-PCOSTE)),MDA_DEC(EJERANY))
                    ENDIF

                    IMP1:=IMP1+ROUND((UNIDAD*PVENTA),MDA_DEC(EJERANY))
                    SKIP
                ENDDO
                DES1:=0                             //ROUND(IMP1*(DESC2/100),MDA_DEC(EJERANY))
                BIMP2:=IMP1-DES1
                aRESUMEN[LEN(aRESUMEN),MONTH(FTIC2)]:=aRESUMEN[LEN(aRESUMEN),MONTH(FTIC2)]+BIMP2
            ENDDO
            SET DELETE ON
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***
RETURN Nil


STATIC FUNCTION ResVentaGrid(aRESUMEN)
    TOT1:=0
    TOT2:=0
    aTOT:=ARRAY(LEN(aRESUMEN)+1)
    AFill(aTOT,0)
    W_ResVenta.GR_Resumen.DeleteAllItems
    FOR N1=1 TO 12
        FOR N2=1 TO LEN(aRESUMEN)
            aTOT[N2+1]:=aTOT[N2+1]+aRESUMEN[N2,N1]
            IF W_ResVenta.R_Acumulado.Value=2 .AND. N1>1
                aRESUMEN[N2,N1]:=aRESUMEN[N2,N1]+aRESUMEN[N2,N1-1]
            ENDIF
        NEXT
    NEXT
    FOR N1=1 TO 12
        aMas:={}
        AADD(aMas,MES(N1))
        FOR N2=1 TO LEN(aRESUMEN)
            AADD(aMas,LTRIM(MIL(aRESUMEN[N2,N1],15,2)))
        NEXT
        W_ResVenta.GR_Resumen.AddItem(aMas)
        DO EVENTS
    NEXT
    FOR N=1 TO LEN(aTOT)
        aTOT[N]:=LTRIM(MIL(aTOT[N],15,2))
    NEXT
    aTOT[1]:="Total"
    W_ResVenta.GR_Resumen.AddItem(aTOT)
    W_ResVenta.GR_Resumen.Value:=1

RETURN Nil


STATIC FUNCTION ResVentaGraf()
    IF W_ResVenta.GR_Resumen.ItemCount=0
        MsgStop("No hay datos en las fechas introducidas","error")
        RETURN
    ENDIF

    aSer:={}
    aTit:=MES(,1)
    aNomSer:={}
    FOR NCOL=2 TO GetGridColumnsCount("GR_Resumen","W_ResVenta")
        AADD(aNomSer,W_ResVenta.GR_Resumen.Header(NCOL))
        aSer1:={}
        FOR N=1 TO W_ResVenta.GR_Resumen.ItemCount-1
            AADD(aSer1,MILQUITAR(W_ResVenta.GR_Resumen.Cell(N,NCOL)))
        NEXT
        AADD(aSer,aSer1)
    NEXT
    Graf1(aSer,aTit,aNomSer,W_ResVenta.Title)
RETURN Nil


PROCEDURE ResVentaImp()
    local oprint

    IF W_ResVenta.GR_Resumen.ItemCount=0
        MsgStop("No hay datos en las fechas introducidas","error")
        RETURN
    ENDIF

    nTOTCOL:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    VerHoriz:=IF(nTOTCOL>6,.T.,.F.)

    oprint:=tprint(UPPER(W_ResVenta.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",5)
    oprint:selprinter(W_ResVenta.nImp.value , W_ResVenta.nVer.value , VerHoriz , 9 , W_ResVenta.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(W_ResVenta.Title)
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    PAG:=1

    oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
    oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
    oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

    oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
    oprint:printdata(32,105,W_ResVenta.Title,"times new roman",18,.F.,,"C",)

    oprint:printdata(40,20,'Desde el ejercicio: '+LTRIM(STR(W_ResVenta.Ejer1.value)),"times new roman",12,.F.,,"L",)
    oprint:printdata(45,20,'Hasta el ejercicio: '+LTRIM(STR(W_ResVenta.Ejer2.value)),"times new roman",12,.F.,,"L",)

    LIN:=55
    FOR COL2=1 TO nTOTCOL
        IF COL2=1
            oprint:printdata(LIN,20, W_ResVenta.GR_Resumen.Header(COL2) ,"times new roman",10,.F.,,"L",)
        ELSE
            oprint:printdata(LIN,(COL2*25)+20, W_ResVenta.GR_Resumen.Header(COL2) ,"times new roman",10,.F.,,"R",)
        ENDIF
    NEXT
    oprint:printline(LIN+4,20,LIN+4,(nTOTCOL*25)+20,,0.5)
    LIN:=LIN+5
    FOR LIN2=1 TO W_ResVenta.GR_Resumen.ItemCount
        FOR COL2=1 TO nTOTCOL
            IF COL2=1
                oprint:printdata(LIN,20, W_ResVenta.GR_Resumen.Cell(LIN2,COL2) ,"times new roman",10,.F.,,"L",)
            ELSE
                oprint:printdata(LIN,(COL2*25)+20, W_ResVenta.GR_Resumen.Cell(LIN2,COL2) ,"times new roman",10,.F.,,"R",)
            ENDIF
        NEXT
        LIN:=LIN+5
    NEXT
    oprint:printline(LIN-6,20,LIN-6,(nTOTCOL*25)+20,,0.5)

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

Return Nil


STATIC FUNCTION ResVentaCalc()
    local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

    IF W_ResVenta.GR_Resumen.ItemCount=0
        MsgStop("No hay datos en las fechas introducidas","error")
        RETURN
    ENDIF

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

    // escribir en las celdas
    aMiColor:=MiColor("AMARILLOPALIDO")
    nTOTCOL:=GetGridColumnsCount("GR_Resumen","W_ResVenta")
    FOR N1=1 TO W_ResVenta.GR_Resumen.ItemCount-1
        FOR N2=1 TO nTOTCOL
            IF N1=1
                oCell:=oSheet:getCellByPosition(N2-1,1)  // columna,linea
                oCell:SetValue(VAL(W_ResVenta.GR_Resumen.Header(N2)))
                oCell:CharWeight:=150               //NEGRITA
                oCell:CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
            ENDIF
            oCell:=oSheet:getCellByPosition(N2-1,N1+1)  // columna,linea
            IF N2=1
                oCell:SetString(W_ResVenta.GR_Resumen.Cell(N1,N2))
            ELSE
                oCell:SetValue(MILQUITAR(W_ResVenta.GR_Resumen.Cell(N1,N2)))
            ENDIF
        NEXT
        oSheet:getCellRangeByPosition(1,N1+1,nTOTCOL-1,N1+1):NumberFormat:=4  //#.##0,00
    NEXT

    IF W_ResVenta.R_Acumulado.Value=1
        LIN1:=3
        LIN:=W_ResVenta.GR_Resumen.ItemCount+1
        oSheet:getCellByPosition(0,LIN):SetString("TOTAL")
        FOR N1=1 TO nTOTCOL-1
            oSheet:getCellByPosition(N1,LIN):SetFormula("=SUM("+LetraExcel(N1+1)+LTRIM(STR(LIN1))+":"+LetraExcel(N1+1)+LTRIM(STR(LIN))+")")
        NEXT
        oSheet:getCellRangeByPosition(1,LIN,nTOTCOL-1,LIN):NumberFormat:=4  //#.##0,00
        oSheet:getCellRangeByPosition(0,LIN,nTOTCOL-1,LIN):CharWeight:=150  //NEGRITA
        aMiColor:=MiColor("VERDEPALIDO")
        oSheet:getCellRangeByPosition(0,LIN,nTOTCOL-1,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
    ENDIF

    oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)

    oCell:=oSheet:getCellByPosition(0,0)
    oCell:SetString(W_ResVenta.Title)
    oSheet:getCellRangeByPosition(0,0,0,0):CharWeight:=150  //NEGRITA

    QuitarEspera()

Return nil
