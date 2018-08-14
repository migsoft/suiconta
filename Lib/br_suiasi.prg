#include "minigui.ch"

Function br_suizoasi(RUTA1,ASIENTO1,NUMEMP1)
    NUMEMP1:=IF(VALTYPE(NUMEMP1)="N",STR(NUMEMP1),NUMEMP1)
    aAPUNTE:={}
    NOMEMP1:=""
    CUADRE2:=0
    DO CASE
    CASE AT("SUICONTA",UPPER(RUTA1))<>0
        IF AT("SUIZO",UPPER(RUTA1))<>0 .AND. VAL(NUMEMP1)<>0
            RUTA2:=LEFT(RUTA1, RAT("SUIZO",UPPER(RUTA1))-2 )
        ELSE
            RUTA2:=RUTA1
        ENDIF
        IF FILE(RUTA2+"\EMPRESA.DBF") .AND. ;
            FILE(RUTA2+"\EMPRESA.CDX")
            AbrirDBF("EMPRESA",,,,RUTA2)
            SEEK VAL(NUMEMP1)
            RUTA2:=RUTA2+"\"+RTRIM(RUTA)
            NOMEMP1:=RTRIM(EMP)+" "+LTRIM(STR(EJERCICIO))
            EMPRESA->( DBCLOSEAREA() )
        ENDIF
        IF FILE(RUTA2+"\APUNTES.DBF") .AND. ;
            FILE(RUTA2+"\APUNTES.CDX")
            AbrirDBF("APUNTES",,,,RUTA2)
            SEEK ASIENTO1
            DO WHILE NASI=ASIENTO1 .AND. .NOT. EOF()
                AADD(aAPUNTE,{STR(NASI),DIA(FECHA,10),STR(CODCTA),"",RTRIM(NOMAPU),DEBE,HABER})
                CUADRE2:=CUADRE2+DEBE-HABER
                SKIP
            ENDDO
            APUNTES->( DBCLOSEAREA() )
            IF FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
                FILE(RUTA2+"\CUENTAS.CDX")
                AbrirDBF("CUENTAS",,,,RUTA2)
                FOR NA=1 TO LEN(aAPUNTE)
                    SEEK VAL(aAPUNTE[NA,3])
                    IF .NOT. EOF()
                        aAPUNTE[NA,4]:=RTRIM(NOMCTA)
                    ENDIF
                NEXT
                CUENTAS->( DBCLOSEAREA() )
            ENDIF
        ELSE
            RETURN Nil
        ENDIF
    CASE AT("GRUPOSP",UPPER(RUTA1))<>0
        RUTA2:=LEFT(RUTA1,RAT("EMP",UPPER(RUTA1))-2)
        RUTA2:=RUTA2+"\EMP"
        IF FILE(RUTA2+"\EMPRESA.DBF") .AND. ;
            FILE(RUTA2+"\EMPRESA.CDX")
            AbrirDBF("EMPRESA",,,,RUTA2)
            DBSETORDER(1)
            NUMEMP1:=PADR(NUMEMP1,LEN(COD)," ")
            SEEK NUMEMP1
            IF EOF() .AND. VAL(NUMEMP1)<>0
                LOCATE FOR VAL(COD)=VAL(NUMEMP1)
                NUMEMP1:=COD
            ENDIF
            NOMEMP1:=RTRIM(NOMBRE)+" "+LTRIM(EJERCICIO)
            EMPRESA->( DBCLOSEAREA() )
        ENDIF
        RUTA2:=RUTA2+NUMEMP1
        IF FILE(RUTA2+"\DIARIO.DBF") .AND. ;
            FILE(RUTA2+"\DIARIO.CDX")
            AbrirDBF("DIARIO",,,,RUTA2)
            DBSETORDER(3)
            SEEK ASIENTO1
            DO WHILE ASIEN=ASIENTO1 .AND. .NOT. EOF()
                IF DELETE()=.F.
                    AADD(aAPUNTE,{LTRIM(STR(ASIEN)),DIA(FECHA,10),SUBCTA,"",RTRIM(CONCEPTO),EURODEBE,EUROHABER})
                    CUADRE2:=CUADRE2+EURODEBE-EUROHABER
                ENDIF
                SKIP
            ENDDO
            DIARIO->( DBCLOSEAREA() )
            IF FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
                FILE(RUTA2+"\SUBCTA.CDX")
                AbrirDBF("SUBCTA",,,,RUTA2)
                DBSETORDER(1)
                FOR NA=1 TO LEN(aAPUNTE)
                    SEEK PADR(LTRIM(aAPUNTE[NA,3]),LEN(COD)," ")
                    IF .NOT. EOF()
                        aAPUNTE[NA,4]:=RTRIM(TITULO)
                    ENDIF
                NEXT
                SUBCTA->( DBCLOSEAREA() )
            ENDIF
        ELSE
            RETURN Nil
        ENDIF
    OTHERWISE
        RETURN Nil
    ENDCASE

    IF LEN(aAPUNTE)=0
        RETURN Nil
    ENDIF

    DEFINE WINDOW WinBRasi1 ;
        AT 0,0     ;
        WIDTH 750  ;
        HEIGHT 450 ;
        TITLE "Consulta de asientos" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("AZULCLARO")

        @ 10,10 LABEL L_Empresa VALUE NOMEMP1 AUTOSIZE TRANSPARENT
        @ 10,300 LABEL L_Ruta VALUE RUTA2 AUTOSIZE TRANSPARENT

        @ 45,10 LABEL L_Asiento VALUE "Asiento" AUTOSIZE TRANSPARENT
        @ 40,60 TEXTBOX T_Asiento WIDTH 100 VALUE VAL(aAPUNTE[1,1]) ;
            NUMERIC INPUTMASK '999,999,999,999' FORMAT 'E' RIGHTALIGN

        @ 45,250 LABEL L_Fecha VALUE "Fecha" AUTOSIZE TRANSPARENT
        @ 40,300 DATEPICKER D_Fecha WIDTH 100 VALUE CTOD(aAPUNTE[1,2]) NOTABSTOP

        @ 45,500 LABEL L_Cuadre VALUE "Cuadre" AUTOSIZE TRANSPARENT
        @ 40,550 TEXTBOX T_Cuadre WIDTH 100 VALUE CUADRE2 ;
            NUMERIC INPUTMASK '999,999,999,999.99' FORMAT 'E' RIGHTALIGN

        @ 70,10 GRID G_Asiento ;
            HEIGHT 300 ;
            WIDTH 710 ;
            HEADERS {,,'Cuenta','Descripcion cuenta','Descripcion apunte','Debe','Haber'} ;
            WIDTHS {0,0,80,150,250,100,100 } ;
            ITEMS aAPUNTE ;
            JUSTIFY {,,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            VALUE 1 ;
            EDIT INPLACE ;
            COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
            NAVIGATEBYCELL

        Menu_Grid("WinBRasi1","G_Asiento","MENU",,{"COPCELPP","COPREGPP","COPTABPP"})

        @380,430 BUTTONEX Bt_GuardarF CAPTION 'Fecha' ICON icobus('guardar') ;
            ACTION br_suizoasiG(RUTA2,ASIENTO1,NUMEMP1,"Fecha") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @380,530 BUTTONEX Bt_GuardarD CAPTION 'Descripcion' ICON icobus('guardar') ;
            ACTION br_suizoasiG(RUTA2,ASIENTO1,NUMEMP1,"Descripcion") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @380,630 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
            ACTION WinBRasi1.Release ;
            WIDTH 90 HEIGHT 25 ;
            NOTABSTOP

        WinBRasi1.T_Asiento.ReadOnly:= .T.
        WinBRasi1.T_Cuadre.ReadOnly := .T.
        WinBRasi1.G_Asiento.SetFocus

    END WINDOW
    VentanaCentrar("WinBRasi1","Ventana1")
    CENTER WINDOW WinBRasi1
    ACTIVATE WINDOW WinBRasi1

Return Nil

STATIC FUNCTION br_suizoasiG(RUTA2b,ASIENTO1,NUMEMP1,LLAMADA)
    DEFAULT LLAMADA TO " "
    RUTA2:=RUTA2b
    IF MSGYESNO("¿Desea modificar la "+LOWER(LLAMADA)+" de los asientos?")=.F.
        RETURN Nil
    ENDIF

    PonerEspera("Modificando asiento....")

    IF FILE(RUTA2+"\APUNTES.DBF") .AND. ;
        FILE(RUTA2+"\APUNTES.CDX")
        AbrirDBF("APUNTES",,,,RUTA2)
        SEEK ASIENTO1
        IF YEAR(FECHA)<>YEAR(WinBRasi1.D_Fecha.Value)
            MsgStop("La fecha del asiento esta fuera del ejercicio")
            QuitarEspera()
            RETURN Nil
        ENDIF
        LIN2:=1
        DO WHILE NASI=ASIENTO1 .AND. .NOT. EOF()
            IF WinBRasi1.G_Asiento.ItemCount>=LIN2
                IF RLOCK()
                    IF UPPER(LLAMADA)="FECHA"
                        REPLACE FECHA WITH WinBRasi1.D_Fecha.Value
                    ELSE
                        REPLACE NOMAPU WITH WinBRasi1.G_Asiento.Cell(LIN2,5)
                    ENDIF
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
            ENDIF
            LIN2++
            SKIP
        ENDDO
        APUNTES->( DBCLOSEAREA() )
    ENDIF
    IF FILE(RUTA2+"\DIARIO.DBF") .AND. ;
        FILE(RUTA2+"\DIARIO.CDX")
        AbrirDBF("DIARIO",,,,RUTA2)
        DBSETORDER(3)
        SEEK ASIENTO1
        IF YEAR(FECHA)<>YEAR(WinBRasi1.D_Fecha.Value)
            MsgStop("La fecha del asiento esta fuera del ejercicio")
            QuitarEspera()
            RETURN Nil
        ENDIF
        LIN2:=1
        DO WHILE ASIEN=ASIENTO1 .AND. .NOT. EOF()
            IF DELETE()=.F.
                IF WinBRasi1.G_Asiento.ItemCount>=LIN2
                    IF RLOCK()
                        IF UPPER(LLAMADA)="FECHA"
                            REPLACE FECHA WITH WinBRasi1.D_Fecha.Value
                        ELSE
                            REPLACE CONCEPTO WITH WinBRasi1.G_Asiento.Cell(LIN2,5)
                        ENDIF
                        DBCOMMIT()
                        DBUNLOCK()
                    ENDIF
                ENDIF
                LIN2++
            ENDIF
            SKIP
        ENDDO
        DIARIO->( DBCLOSEAREA() )
    ENDIF

    QuitarEspera()

Return Nil
