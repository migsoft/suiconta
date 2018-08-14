#include "minigui.ch"

procedure Remlcon()
    TituloImp:="Listado remesa por concepto/importe"

    RUTAREMESA:=RUTAEMPRESA
    RUTAREMEMP:=RUTAPROGRAMA

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 400 ;
        TITLE TituloImp ;
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
            OPTIONS {'Concepto','Importe','Factura'} ;
            VALUE 1 ;
            WIDTH 100 ;
            ON CHANGE Remlcon_ActLis() ;
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

        @195,010 LABEL L_Orden VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
        @190,110 RADIOGROUP R_Orden ;
            OPTIONS {'Codigo','Cliente','Factura','Vto.' } ;
            VALUE 1 WIDTH 75 ;
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

        @330, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION RemlconImp()

        @330,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Remlcon_ActLis()
    IF W_Imp1.R_TipoLis.Value=1 .OR. W_Imp1.R_TipoLis.Value=3
        W_Imp1.T_Concepto.Enabled:=.T.
        W_Imp1.T_Imp1.Enabled:=.F.
        W_Imp1.T_Imp2.Enabled:=.F.
    ELSE
        W_Imp1.T_Concepto.Enabled:=.F.
        W_Imp1.T_Imp1.Enabled:=.T.
        W_Imp1.T_Imp2.Enabled:=.T.
    ENDIF

Return Nil

STATIC FUNCTION RemlconImp()

    IF FILE("FIN.DBF")
        AbrirDBF("Fin","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    AbrirDBF("REMESA")
    COPY STRUC TO FIN
    REMESA->( DBCLOSEAREA() )
    AbrirDBF("Fin","SIN_INDICE")

    FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
        RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
        RUTA2:=RUTA1[1]+"\REMESA.DBF"
        IF .NOT. FILE(RUTA2)
            LOOP
        ENDIF
        DO CASE
        CASE W_Imp1.R_TipoLis.Value=1
            APPEND FROM &RUTA2 FOR FREM>=W_Imp1.D_Fec1.value .AND. FREM<=W_Imp1.D_Fec2.value .AND. ;
                AT(UPPER(W_Imp1.T_Concepto.Value),UPPER(NOMCTA))<>0
        CASE W_Imp1.R_TipoLis.Value=2
            APPEND FROM &RUTA2 FOR FREM>=W_Imp1.D_Fec1.value .AND. FREM<=W_Imp1.D_Fec2.value .AND. ;
                IMPORTE>=W_Imp1.T_Imp1.value .AND. IMPORTE<=W_Imp1.T_Imp2.value
        CASE W_Imp1.R_TipoLis.Value=3
            APPEND FROM &RUTA2 FOR FREM>=W_Imp1.D_Fec1.value .AND. FREM<=W_Imp1.D_Fec2.value .AND. ;
                AT(UPPER(W_Imp1.T_Concepto.Value),UPPER(NFRA))<>0
        ENDCASE
    NEXT

    DO CASE
    CASE W_Imp1.R_Orden.Value=1
        INDEX ON STR(CODCTA,8)+DTOS(FREM)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=2
        INDEX ON UPPER(NOMCTA)+DTOS(FREM)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=3
        INDEX ON NFRA+STR(CODCTA,8) TO FIN
    CASE W_Imp1.R_Orden.Value=4
        INDEX ON DTOS(FVTO)+STR(CODCTA,8)+NFRA TO FIN
    ENDCASE

    GO TOP
    IF LASTREC()=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF

    RemlconImp2()

Return Nil

STATIC FUNCTION RemlconImp2()
    local oprint

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

    TOT1:=0
    TOTREC2:=0
    PAG:=0
    LIN:=0
    SELEC FIN
    GO TOP
    DO WHILE .NOT. EOF()
        DO EVENTS

        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN,160,'Suma',"times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,190,MIL(TOT1,15,2),"times new roman",8,.F.,,"R",)
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
            oprint:printdata(32,105,TituloImp,"times new roman",12,.F.,,"C",)

            LIN:=40
            oprint:printdata(LIN, 32,'Factura',"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN, 34,'Codigo',"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 47,'Descripcion',"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,102,'Cuenta',"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,117,'Talon',"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,150,'Vencimiento',"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,190,'Importe',"times new roman",8,.F.,,"R",)
            oprint:printline(LIN+4,20,LIN+4,190,,0.5)

            LIN:=LIN+4
        ENDIF

        oprint:printdata(LIN, 32,RTRIM(NFRA),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN, 34,CODCTA,"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN, 47,RTRIM(NOMCTA),"times new roman",8,.F.,,"L",)
        IF SERIE=3 .OR. SERIE=4
            oprint:printdata(LIN,102,LEFT(BANCTA,9),"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,117,TALON,"times new roman",8,.F.,,"L",)
        ELSE
            oprint:printdata(LIN,102,CTA_BAN_SUIZO(BANCTA,24),"times new roman",8,.F.,,"L",)
        ENDIF
        oprint:printdata(LIN,150,DIA(FVTO,10),"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,190,MIL(IMPORTE,15,2),"times new roman",8,.F.,,"R",)

        TOT1:=TOT1+IMPORTE
        TOTREC2++

        LIN:=LIN+4
        SKIP

    ENDDO

    oprint:printdata(LIN,40,MIL(TOTREC2,15,0)+' Recibos',"times new roman",8,.F.,,"L",)
    oprint:printdata(LIN,160,'Total',"times new roman",8,.F.,,"L",)
    oprint:printdata(LIN,190,MIL(TOT1,15,2),"times new roman",8,.F.,,"R",)

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return Nil