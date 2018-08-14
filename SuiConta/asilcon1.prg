#include "minigui.ch"
#include "winprint.ch"

procedure Asilconcep1()
    TituloImp:="Listado de asiento por concepto/importe"
    FEC1:=CTOD( "01-"+STR(MONTH(DATE()),2)+"-"+STR(YEAR(DATE()),4) )

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 420 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_Fec1 VALUE 'Desde la fecha' AUTOSIZE TRANSPARENT
        @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 15,215 LABEL L_Fec1b VALUE 'Año = ejercicios anteriores' AUTOSIZE TRANSPARENT

        @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2
        @ 45,215 LABEL L_Fec2b VALUE 'Año = ejercicios posteriores' AUTOSIZE TRANSPARENT

        @ 75,10 LABEL L_TipoLis VALUE 'Buscar por' AUTOSIZE TRANSPARENT
        @ 70,110 RADIOGROUP R_TipoLis ;
            OPTIONS { 'Concepto' , 'Importe' } ;
            VALUE 1 ;
            WIDTH 100 ;
            ON CHANGE Asilconcep1_ActLis() ;
            HORIZONTAL

        @105, 10 LABEL L_Concepto VALUE 'Concepto' AUTOSIZE TRANSPARENT
        @100,110 TEXTBOX T_Concepto WIDTH 270 MAXLENGTH 30

        @135,10 LABEL L_Imp1 VALUE 'Desde importe' AUTOSIZE TRANSPARENT
        @130,110 TEXTBOX T_Imp1 WIDTH 120 HEIGHT 25 VALUE -9999999999 ;
            NUMERIC INPUTMASK '99,999,999,999.99 €' FORMAT 'E' RIGHTALIGN

        @165,10 LABEL L_Imp2 VALUE 'Hasta importe' AUTOSIZE TRANSPARENT
        @160,110 TEXTBOX T_Imp2 WIDTH 120 HEIGHT 25 VALUE 9999999999 ;
            NUMERIC INPUTMASK '99,999,999,999.99 €' FORMAT 'E' RIGHTALIGN
        W_Imp1.T_Imp1.Enabled:=.F.
        W_Imp1.T_Imp2.Enabled:=.F.

        @195,10 LABEL L_Orden VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
        @190,110 RADIOGROUP R_Orden ;
            OPTIONS { 'Fecha' , 'Concepto' , 'Importe' } ;
            VALUE 1 ;
            WIDTH 100 ;
            HORIZONTAL


        draw rectangle in window W_Imp1 at 230,010 to 232,390 fillcolor{255,0,0}  //Rojo
        aIMP:=Impresoras(EMP_IMPRESORA)
        @245,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @240,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @275,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @270,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @270, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @300, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @330,10 LABEL L_Progress_1 VALUE 'Imprimiendo' AUTOSIZE TRANSPARENT
        @330,100 PROGRESSBAR Progress_1 WIDTH 280 HEIGHT 15 SMOOTH

        @355, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION Asilconcep1i("IMPRESORA")

        @355,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil

STATIC FUNCTION Asilconcep1_ActLis()
    IF W_Imp1.R_TipoLis.Value=1
        W_Imp1.T_Concepto.Enabled:=.T.
        W_Imp1.T_Imp1.Enabled:=.F.
        W_Imp1.T_Imp2.Enabled:=.F.
    ELSE
        W_Imp1.T_Concepto.Enabled:=.F.
        W_Imp1.T_Imp1.Enabled:=.T.
        W_Imp1.T_Imp2.Enabled:=.T.
    ENDIF
Return Nil

procedure Asilconcep1i(LLAMADA)
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

    FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
        RUTA2:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
        RUTA2:=RUTA2[1]+"\APUNTES.DBF"
        IF .NOT. FILE(RUTA2)
            LOOP
        ENDIF
        IF W_Imp1.R_TipoLis.Value=1
            APPEND FROM &RUTA2 FOR FECHA>=W_Imp1.D_Fec1.value .AND. FECHA<=W_Imp1.D_Fec2.value ;
                .AND. AT(UPPER(W_Imp1.T_Concepto.Value),UPPER(NOMAPU))<>0
        ELSE
            APPEND FROM &RUTA2 FOR FECHA>=W_Imp1.D_Fec1.value .AND. FECHA<=W_Imp1.D_Fec2.value ;
                .AND. DEBE+HABER>=W_Imp1.T_Imp1.value .AND. DEBE+HABER<=W_Imp1.T_Imp2.value
        ENDIF
    NEXT


    DO CASE
    CASE W_Imp1.R_Orden.value=1                     //Fecha
        INDEX ON DTOS(FECHA) TAG ORDEN1 TO FIN
    CASE W_Imp1.R_Orden.value=2                     //Concepto
        INDEX ON UPPER(NOMAPU)+DTOS(FECHA) TAG ORDEN1 TO FIN
    CASE W_Imp1.R_Orden.value=3                     //Importe
        INDEX ON STR(DEBE+HABER,15,5)+DTOS(FECHA) TAG ORDEN1 TO FIN
    ENDCASE

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

    W_Imp1.Progress_1.RangeMax:=LASTREC()
    W_Imp1.Progress_1.Value:=0
    PAG:=0
    LIN:=0
    aCOL:={}

    AADD(aCOL, 08)                                  //1-tamaño letra
    AADD(aCOL, 28)                                  //2-num. asiento
    AADD(aCOL, 36)                                  //3-fecha
    AADD(aCOL, 49)                                  //4-codigo cuenta
    AADD(aCOL, 62)                                  //5-nombre cuenta
    AADD(aCOL, 90)                                  //6-descripcion apunte
    AADD(aCOL,155)                                  //7-debe
    AADD(aCOL,172)                                  //8-haber
    AADD(aCOL,190)                                  //9-saldo
    AADD(aCOL,190)                                  //10-referencia num.asiento

    DO WHILE .NOT. EOF()
        W_Imp1.Progress_1.Value:=W_Imp1.Progress_1.Value+1
        DO EVENTS

        IF LIN>=265 .OR. PAG=0 .OR. SALTA2=1
            SALTA2:=0
            IF PAG<>0
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
            IF W_Imp1.R_TipoLis.Value=1
                oprint:printdata(40,120,'Concepto: ',"times new roman",12,.F.,,"L",)
                oprint:printdata(45,120,W_Imp1.T_Concepto.Value,"times new roman",12,.F.,,"L",)
            ELSE
                oprint:printdata(40,130,'Desde importe: '+LTRIM(MIL(W_Imp1.T_Imp1.value,15,0)),"times new roman",12,.F.,,"L",)
                oprint:printdata(45,130,'Hasta importe: '+LTRIM(MIL(W_Imp1.T_Imp2.value,15,0)),"times new roman",12,.F.,,"L",)
            ENDIF
            LIN:=55
            oprint:printdata(LIN,aCOL[2],"Asi/Apu","times new roman",aCOL[1],.F.,,"C",)
            oprint:printdata(LIN,aCOL[3],"Fecha","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[4],"Cuenta","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[5],"Descripcion","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[6],"Concepto","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[7],"Debe","times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[8],"Haber","times new roman",aCOL[1],.F.,,"R",)
            oprint:printline(LIN+4,15,LIN+4,aCOL[8],,0.5)

            LIN:=LIN+5
        ENDIF

        oprint:printdata(LIN,aCOL[2],LTRIM(STR(NASI))+"/"+LTRIM(STR(APU)),"times new roman",aCOL[1],.F.,,"C",)
        oprint:printdata(LIN,aCOL[3],DIA(FECHA,8),"times new roman",aCOL[1],.F.,,"L",)
        oprint:printdata(LIN,aCOL[4],CODCTA,"times new roman",aCOL[1],.F.,,"L",)
        CODCTA2:=CODCTA
        AbrirDBF("CUENTAS",,,,,1)
        SEEK CODCTA2
        oprint:printdata(LIN,aCOL[5],LEFT(NOMCTA,15),"times new roman",aCOL[1],.F.,,"L",)
        AbrirDBF("FIN")
        oprint:printdata(LIN,aCOL[6],NOMAPU,"times new roman",aCOL[1],.F.,,"L",)

        oprint:printdata(LIN,aCOL[7],MIL(DEBE ,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[8],MIL(HABER,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)

        LIN:=LIN+5
        SKIP

    ENDDO

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return


