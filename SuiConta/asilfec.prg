#include "minigui.ch"
#include "winprint.ch"

procedure Asilfecha(LLAMADA)
    IF LLAMADA="LISTADO"
        TituloImp:="Listado de asientos"
    ELSE
        TituloImp:="Libro diario"
    ENDIF

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 600 HEIGHT 420 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

        @ 75,10 LABEL L_Nasi1 VALUE 'Desde asiento' AUTOSIZE TRANSPARENT
        @ 70,110 TEXTBOX T_Nasi1 WIDTH 100 VALUE 0 NUMERIC RIGHTALIGN
        @105,10 LABEL L_Nasi2 VALUE 'Hasta asiento' AUTOSIZE TRANSPARENT
        @100,110 TEXTBOX T_Nasi2 WIDTH 100 VALUE 999999999 NUMERIC RIGHTALIGN

        @130,10 CHECKBOX C_Apertura CAPTION 'Incluir asientos de apertura' ;
            WIDTH 170 VALUE .F. TRANSPARENT NOTABSTOP
        @160,10 CHECKBOX C_Cierre CAPTION 'Incluir asientos de cierre' ;
            WIDTH 160 VALUE .F. TRANSPARENT NOTABSTOP

        @195,10 LABEL L_Orden VALUE 'Ordenado por' AUTOSIZE TRANSPARENT
        @190,110 RADIOGROUP R_Orden OPTIONS { 'Fecha' , 'Asiento' } ;
            VALUE 1 WIDTH 75 HORIZONTAL


        @ 15,350 LABEL L_Datos VALUE 'Datos iniciales del listado' AUTOSIZE BOLD TRANSPARENT
        @ 45,350 LABEL L_LISASI VALUE 'Nº asiento' AUTOSIZE TRANSPARENT
        @ 40,450 TEXTBOX T_LISASI WIDTH 100 VALUE 1 NUMERIC RIGHTALIGN
        @ 75,350 LABEL L_LISPAG VALUE 'Nº de hoja' AUTOSIZE TRANSPARENT
        @ 70,450 TEXTBOX T_LISPAG WIDTH 100 VALUE 1 NUMERIC RIGHTALIGN
        @105,350 LABEL L_LISFEC VALUE 'Fecha del listado' AUTOSIZE TRANSPARENT
        @100,450 DATEPICKER D_LISFEC WIDTH 100 VALUE DATE()
        @135,350 LABEL L_LISDEB VALUE 'Sumas al debe' AUTOSIZE TRANSPARENT
        @130,450 TEXTBOX T_LISDEB WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
        @165,350 LABEL L_LISHAB VALUE 'Sumas al haber' AUTOSIZE TRANSPARENT
        @160,450 TEXTBOX T_LISHAB WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN


        @220,10 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

        draw rectangle in window W_Imp1 at 260,010 to 262,390 fillcolor{255,0,0}  //Rojo
        aIMP:=Impresoras(EMP_IMPRESORA)
        @275,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @270,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @305,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @300,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @300, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @330, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @360, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION Apulfechai(LLAMADA)

        @360,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

        IF LLAMADA="LISTADO"
            W_Imp1.L_Datos.Visible:=.F.
            W_Imp1.L_LISASI.Visible:=.F.
            W_Imp1.T_LISASI.Visible:=.F.
            W_Imp1.L_LISPAG.Visible:=.F.
            W_Imp1.T_LISPAG.Visible:=.F.
            W_Imp1.L_LISFEC.Visible:=.F.
            W_Imp1.D_LISFEC.Visible:=.F.
            W_Imp1.L_LISDEB.Visible:=.F.
            W_Imp1.T_LISDEB.Visible:=.F.
            W_Imp1.L_LISHAB.Visible:=.F.
            W_Imp1.T_LISHAB.Visible:=.F.
        ELSE
            W_Imp1.L_Orden.Visible:=.F.
            W_Imp1.R_Orden.Visible:=.F.
        ENDIF

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Apulfechai(LLAMADA)
    local oprint

    IF FILE("FIN.DBF")
        AbrirDBF("Fin","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    AbrirDBF("APUNTES")
    COPY STRUC TO FIN.DBF
    USE
    AbrirDBF("FIN","SIN_INDICE",,"Exclusive",,)

   ***INCLUIR ASIENTO DE APERTURA***
    IF W_Imp1.C_Apertura.value=.T.
        APPEND FROM APUNTES FOR FECHA>=W_Imp1.D_Fec1.value .AND. FECHA<=W_Imp1.D_Fec2.value ;
            .AND. NASI>=W_Imp1.T_Nasi1.value .AND. NASI<=W_Imp1.T_Nasi2.value
    ELSE
        APPEND FROM APUNTES FOR FECHA>=W_Imp1.D_Fec1.value .AND. FECHA<=W_Imp1.D_Fec2.value ;
            .AND. NASI>=W_Imp1.T_Nasi1.value .AND. NASI<=W_Imp1.T_Nasi2.value .AND. NASI<>1
    ENDIF

    IF W_Imp1.C_Cierre.value=.T.
        APPEND FROM CIERRE FOR FECHA>=W_Imp1.D_Fec1.value .AND. FECHA<=W_Imp1.D_Fec2.value ;
            .AND. NASI>=W_Imp1.T_Nasi1.value .AND. NASI<=W_Imp1.T_Nasi2.value
    ENDIF

    IF W_Imp1.R_Orden.value=1
        INDEX ON DTOS(FECHA)+STR(NASI,6)+STR(APU,3) TO FIN
    ELSE
        INDEX ON STR(NASI,6)+STR(APU,3)+DTOS(FECHA) TO FIN
    ENDIF

    GO TOP
    IF LASTREC()=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",5)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(TituloImp)
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    W_Imp1.P_Progres.RangeMax:=LASTREC()
    W_Imp1.P_Progres.Value:=0
    PAG:=0
    LIN:=0
    NASI2:=NASI
    LISASI:=W_Imp1.T_LISASI.value
    LISDEB:=W_Imp1.T_LISDEB.value
    LISHAB:=W_Imp1.T_LISHAB.value
    SAL2:=0
    SALTA2:=0
    LISMES:=MONTH(FECHA)
    aCOL:={}
    IF LLAMADA="LISTADO"
        AADD(aCOL,  8)                              //1-tamaño letra
        AADD(aCOL, 28)                              //2-num. asiento
        AADD(aCOL, 36)                              //3-fecha
        AADD(aCOL, 49)                              //4-codigo cuenta
        AADD(aCOL, 62)                              //5-nombre cuenta
        AADD(aCOL, 90)                              //6-descripcion apunte
        AADD(aCOL,155)                              //7-debe
        AADD(aCOL,172)                              //8-haber
        AADD(aCOL,190)                              //9-saldo
        AADD(aCOL,190)                              //10-referencia num.asiento
    ELSE
        AADD(aCOL,  8)                              //1-tamaño letra
        AADD(aCOL, 23)                              //2-num. asiento
        AADD(aCOL, 31)                              //3-fecha
        AADD(aCOL, 44)                              //4-codigo cuenta
        AADD(aCOL, 57)                              //5-nombre cuenta
        AADD(aCOL, 85)                              //6-descripcion apunte
        AADD(aCOL,150)                              //7-debe
        AADD(aCOL,167)                              //8-haber
        AADD(aCOL,185)                              //9-saldo
        AADD(aCOL,195)                              //10-referencia num.asiento
    ENDIF
    DO WHILE .NOT. EOF()
        W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
        DO EVENTS

        IF LIN>=265 .OR. PAG=0 .OR. SALTA2=1
            SALTA2:=0
            IF PAG<>0
                oprint:printdata(LIN,aCOL[7]-30,"Sumas","times new roman",aCOL[1],.F.,,"R",)
                oprint:printdata(LIN,aCOL[7],MIL(LISDEB,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                NUM2:=IF(LISHAB>999999,5,0)
                oprint:printdata(LIN,aCOL[8]+NUM2,MIL(LISHAB,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                LIN:=LIN+5
                oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
            oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

            oprint:printdata(40,20,'Desde fecha: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,20,'Hasta fecha: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(40,130,'Desde asiento: '+LTRIM(MIL(W_Imp1.T_Nasi1.value,15,0)),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,130,'Hasta asiento: '+LTRIM(MIL(W_Imp1.T_Nasi2.value,15,0)),"times new roman",12,.F.,,"L",)

            LIN:=55
            oprint:printdata(LIN,aCOL[2],"Asi/Apu","times new roman",aCOL[1],.F.,,"C",)
            oprint:printdata(LIN,aCOL[3],"Fecha","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[4],"Cuenta","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[5],"Descripcion","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[6],"Concepto","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[7],"Debe","times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[8],"Haber","times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[9],"Saldo","times new roman",aCOL[1],.F.,,"R",)
            IF LLAMADA="LIBRO"
                oprint:printdata(LIN,aCOL[10],"Ref.","times new roman",aCOL[1],.F.,,"R",)
            ENDIF
            oprint:printline(LIN+4,15,LIN+4,aCOL[10],,0.5)

            LIN:=LIN+5
        ENDIF

        IF LLAMADA="LISTADO"
            oprint:printdata(LIN,aCOL[2],LTRIM(STR(NASI))+"/"+LTRIM(STR(APU)),"times new roman",aCOL[1],.F.,,"C",)
        ELSE
            oprint:printdata(LIN,aCOL[2],LTRIM(STR(LISASI))+"/"+LTRIM(STR(APU)),"times new roman",aCOL[1],.F.,,"C",)
        ENDIF
        oprint:printdata(LIN,aCOL[3],DIA(FECHA,8),"times new roman",aCOL[1],.F.,,"L",)
        oprint:printdata(LIN,aCOL[4],CODCTA,"times new roman",aCOL[1],.F.,,"L",)
        CODCTA2:=CODCTA
        AbrirDBF("CUENTAS",,,,,1)
        SEEK CODCTA2
        oprint:printdata(LIN,aCOL[5],LEFT(NOMCTA,15),"times new roman",aCOL[1],.F.,,"L",)
        AbrirDBF("FIN")
        oprint:printdata(LIN,aCOL[6],NOMAPU,"times new roman",aCOL[1],.F.,,"L",)

        LISDEB:=LISDEB+DEBE
        LISHAB:=LISHAB+HABER
        SAL2:=SAL2+DEBE-HABER

        oprint:printdata(LIN,aCOL[7],MIL(DEBE ,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[8],MIL(HABER,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[9],MIL(SAL2 ,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        IF LLAMADA="LIBRO"
            oprint:printdata(LIN,aCOL[10],NASI,"times new roman",aCOL[1],.F.,,"R",)
        ENDIF

        LIN:=LIN+5
        SKIP

        IF NASI2<>NASI
            IF NASI2=0                              //ASIENTO DE APERTURA
                SALTA2:=1
            ENDIF
            NASI2:=NASI
            LISASI++
            SAL2:=0
            LIN:=LIN+5
            IF W_Imp1.R_Orden.value=1 .AND. LISMES<>MONTH(FECHA)
                LISMES:=MONTH(FECHA)
                SALTA2:=1
            ENDIF
            IF NASI2>=999997
                SALTA2:=1
            ENDIF
        ENDIF

    ENDDO

    oprint:printdata(LIN,aCOL[7]-30,"Total","times new roman",aCOL[1],.F.,,"R",)
    oprint:printdata(LIN,aCOL[7],MIL(LISDEB,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
    NUM2:=IF(LISHAB>999999,5,0)
    oprint:printdata(LIN,aCOL[8]+NUM2,MIL(LISHAB,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

Return Nil
