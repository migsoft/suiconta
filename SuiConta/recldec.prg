#include "minigui.ch"
#include "winprint.ch"

procedure Recldec()
    TituloImp:="Listado declaracion modelo 347"

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 420 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 15,215 LABEL L_Fec1b VALUE 'Año = ejercicios anteriores' AUTOSIZE TRANSPARENT

        @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2
        @ 45,215 LABEL L_Fec2b VALUE 'Año = ejercicios posteriores' AUTOSIZE TRANSPARENT

        @ 75,010 LABEL L_Importe VALUE 'Volumen de operaciones superior a' AUTOSIZE TRANSPARENT
        @ 70,220 TEXTBOX T_Importe  WIDTH 110 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN

        @105,010 LABEL L_TipLis VALUE 'Tipo de listado' AUTOSIZE TRANSPARENT
        @100,100 RADIOGROUP R_TipLis ;
            OPTIONS { 'Compras y ventas' , 'Solo compras' , 'Solo ventas' } ;
            VALUE 1 WIDTH 100  ;
            HORIZONTAL NOTABSTOP

        @135,010 LABEL L_MasRet VALUE 'Incrementar las retenciones' AUTOSIZE TRANSPARENT
        @130,175 RADIOGROUP R_MasRet ;
            OPTIONS { 'Si' , 'No' } ;
            VALUE 1 WIDTH 75 ;
            HORIZONTAL NOTABSTOP

        @165,10 LABEL L_OrdLis VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
        @160,100 RADIOGROUP R_OrdLis ;
            OPTIONS {'Codigo cuenta','Nombre cuenta'} ;
            VALUE 1 WIDTH 150 ;
            HORIZONTAL NOTABSTOP

        @195,10 LABEL L_FtoLis VALUE 'Formato del listado' AUTOSIZE TRANSPARENT
        @190,120 RADIOGROUP R_FtoLis ;
            OPTIONS {"Impresora","Columnas","Fichero AEAT"} ;
            VALUE 1 WIDTH 90 ;
            HORIZONTAL NOTABSTOP
        @225,10 LABEL L_FtoLis2 VALUE RUTAPROGRAMA+"\MOD_347.TXT" AUTOSIZE TRANSPARENT

        LINW:=260
        COLW:=0
        draw rectangle in window W_Imp1 at LINW,COLW+10 to LINW+2,COLW+390 fillcolor{255,0,0}  //Rojo
        aIMP:=Impresoras(EMP_IMPRESORA)
        @LINW+15,COLW+10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @LINW+10,COLW+100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @LINW+45,COLW+220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @LINW+40,COLW+280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @LINW+40,COLW+10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @LINW+70,COLW+10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @LINW+100,COLW+10 BUTTONEX B_Imprimir CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION ( Recldeci() , W_Imp1.release )

        @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Recldeci()
    AbrirDBF1("FACREB",,,"Exclusive",,1)
    AbrirDBF1("CUENTAS",,,"Exclusive",,2)
    COPY STRUC TO FIN
    AbrirDBF1("FIN","SIN_INDICE",,"Exclusive",,3)
    INDEX ON STR(CODCTA,8)+NOMCTA TO FIN

    IF W_Imp1.R_TipLis.Value=1 .OR. W_Imp1.R_TipLis.Value=2
*** INCLUIR EJERCICIOS ANTORIORES ***
        FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
            RUTA2:=EMPRUTANY(NUMEMP,ANY2)
            AbrirDBF1("FACREB",,,"Exclusive",RUTA2,1)
            SELEC 1
            DO WHILE .NOT. EOF()
                IF FREG<W_Imp1.D_Fec1.value .OR. FREG>W_Imp1.D_Fec2.value
                    SKIP
                    LOOP
                ENDIF
                CODPRO2:=CODIGO
                NOMPRO2:=NOMCTA
                TOT1=TFAC
                IF W_Imp1.R_MasRet.Value=1
                    TOT1=TOT1+IMPRET
                    RET2:=IF(IMPRET=0," ","*")
                ENDIF
                SELEC 2
                SEEK CODPRO2
                IF .NOT. EOF()
                    DIRCTA2:=DIRCTA
                    POBCTA2:=POBCTA
                    PROCTA2:=PROCTA
                    CPCTA2:=CODPOS
                    CIF2:=CIF
                    NTEL2:=TEL1
                    NFAX2:=FAX1
                ELSE
                    DIRCTA2:=SPACE(30)
                    POBCTA2:=SPACE(30)
                    PROCTA2:=SPACE(30)
                    CPCTA2:=SPACE(15)
                    CIF2:=SPACE(15)
                    NTEL2:=SPACE(15)
                    NFAX2:=SPACE(15)
                ENDIF
                SELEC 3
                IF RIGHT(STR(CODPRO2),5)="00000"
                    SEEK STR(CODPRO2,8)+NOMPRO2
                ELSE
                    SEEK STR(CODPRO2,8)
                ENDIF
                IF EOF()
                    APPEND BLANK
                    REPLACE CODCTA WITH CODPRO2
                    REPLACE NOMCTA WITH NOMPRO2
                    REPLACE DIRCTA WITH DIRCTA2
                    REPLACE POBCTA WITH POBCTA2
                    REPLACE PROCTA WITH PROCTA2
                    REPLACE CODPOS WITH CPCTA2
                    REPLACE CIF WITH CIF2
                    REPLACE TEL1 WITH NTEL2
                    REPLACE FAX1 WITH NFAX2
                ENDIF
                REPLACE DEBE   WITH DEBE+TOT1
                REPLACE SALDO  WITH SALDO+TOT1
                REPLACE BANCTA WITH RET2
                SELEC 1
                SKIP
            ENDDO
        NEXT
    ENDIF

    AbrirDBF1("FAC92",,,"Exclusive",,1)

    IF W_Imp1.R_TipLis.Value=1 .OR. W_Imp1.R_TipLis.Value=3
*** INCLUIR EJERCICIOS ANTORIORES ***
        FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
            RUTA2:=EMPRUTANY(NUMEMP,ANY2)
            AbrirDBF1("FAC92",,,"Exclusive",RUTA2,1)
            SELEC 1
            DO WHILE .NOT. EOF()
                IF FFAC<W_Imp1.D_Fec1.value .OR. FFAC>W_Imp1.D_Fec2.value
                    SKIP
                    LOOP
                ENDIF
                CODCTA2:=CODCTA
                CODCLI2:=COD
                NOMCLI2:=CLIENTE
                TOT1=TFAC
                RET2:=" "
                SELEC 2
                SEEK CODCTA2
                IF .NOT. EOF()
                    DIRCTA2:=DIRCTA
                    POBCTA2:=POBCTA
                    PROCTA2:=PROCTA
                    CPCTA2:=CODPOS
                    CIF2:=CIF
                    NTEL2:=TEL1
                    NFAX2:=FAX1
                ELSE
                    DIRCTA2:=SPACE(30)
                    POBCTA2:=SPACE(30)
                    PROCTA2:=SPACE(30)
                    CPCTA2:=SPACE(15)
                    CIF2:=SPACE(15)
                    NTEL2:=SPACE(15)
                    NFAX2:=SPACE(15)
                ENDIF
                SELEC 3
                IF RIGHT(STR(CODCTA2),5)="00000"
                    SEEK STR(CODCTA2,8)+NOMCLI2
                ELSE
                    SEEK STR(CODCTA2,8)
                ENDIF
                IF EOF()
                    APPEND BLANK
                    REPLACE CODCTA WITH CODCTA2
                    REPLACE NOMCTA WITH NOMCLI2
                    REPLACE DIRCTA WITH DIRCTA2
                    REPLACE POBCTA WITH POBCTA2
                    REPLACE PROCTA WITH PROCTA2
                    REPLACE CODPOS WITH CPCTA2
                    REPLACE CIF WITH CIF2
                    REPLACE TEL1 WITH NTEL2
                    REPLACE FAX1 WITH NFAX2
                ENDIF
                REPLACE HABER WITH HABER+TOT1
                REPLACE SALDO WITH SALDO+TOT1
                SELEC 1
                SKIP
            ENDDO
        NEXT
    ENDIF
    CLOSE DATABASES

    AbrirDBF1("FIN","SIN_INDICE",,"Exclusive",,3)
    DELETE FOR SALDO<W_Imp1.T_Importe.Value
    PACK
    IF W_Imp1.R_OrdLis.Value=1
        INDEX ON STR(CODCTA,8)+NOMCTA TO FIN
    ELSE
        INDEX ON NOMCTA+STR(CODCTA,8) TO FIN
    ENDIF

    IF FIN->(LASTREC())=0
        MsgExclamation("No hay datos en las fechas introducidas","Informacion")
        FIN->( DBCLOSEAREA() )
        RETURN
    ENDIF

    DO CASE
    CASE W_Imp1.R_FtoLis.Value=1
        Recldeci2()
    CASE W_Imp1.R_FtoLis.Value=2
        RecldeciC()
    CASE W_Imp1.R_FtoLis.Value=3
        RecldeciA()
    ENDCASE

Return Nil



STATIC FUNCTION Recldeci2()
    local oprint
    STORE 0 TO PAG,TOT1,TOT2

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",5)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(TituloqImp(W_Imp1.Title))
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    LIN:=0
    REGISTROS:=LASTREC()
    GO TOP
    DO WHILE .NOT. EOF()
        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN,130,"Suma","times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,170,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,190,MIL(TOT2,12,2),"times new roman",8,.F.,,"R",)
                LIN:=LIN+5
                oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",8,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
            oprint:printdata(32,105,TituloqImp(W_Imp1.Title),"times new roman",18,.F.,,"C",)

            oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,20,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

            oprint:printdata(40,90,'Superior a: '+LTRIM(MIL(W_Imp1.T_Importe.Value,12,MDA_DEC(EJERANY))),"times new roman",12,.F.,,"L",)

            LIN:=50
            oprint:printdata(LIN, 20,"Cuenta","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 35,"Descripcion","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 90,"Poblacion","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,130,"CIF/NIF","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,170,"Compras","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,190,"Ventas","times new roman",8,.F.,,"R",)
            oprint:printline(LIN+4,20,LIN+4,190,,0.5)

            LIN:=LIN+5
        ENDIF

        oprint:printdata(LIN, 20,CODCTA,"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN, 35,NOMCTA,"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN, 90,CODPOS+"-"+POBCTA,"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,130,CIF,"times new roman",8,.F.,,"L",)
        IF DEBE<>0
            oprint:printdata(LIN,170,MIL(SALDO,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,172,LEFT(BANCTA,1),"times new roman",8,.F.,,"L",)
            TOT1=TOT1+DEBE
        ELSE
            oprint:printdata(LIN,190,MIL(SALDO,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,192,LEFT(BANCTA,1),"times new roman",8,.F.,,"L",)
            TOT2=TOT2+HABER
        ENDIF
        LIN:=LIN+5
        SKIP
    ENDDO
    oprint:printdata(LIN,130,"Total","times new roman",8,.F.,,"L",)
    oprint:printdata(LIN,170,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
    oprint:printdata(LIN,190,MIL(TOT2,12,2),"times new roman",8,.F.,,"R",)

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

Return Nil



STATIC FUNCTION RecldeciC()
    STORE 0 TO PAG,TOT1,TOT2
    SET PRINTER TO &RUTAPROGRAMA\MOD_347.TXT
    NUM2:=LIM_CAM("WINDOWS")
    GO TOP
    DO WHILE .NOT. EOF()
        IF .NOT. EOF() .AND. PAG=0
            PAG=PAG+1
            SET PRINT ON
            ?? PADR("CUENTA",8," ")
            ?? CHR(NUM2)
            ?? PADR("DESCRIPCION",30," ")
            ?? CHR(NUM2)
            ?? PADR("DIRECCION",30," ")
            ?? CHR(NUM2)
            ?? PADR("CP",5," ")
            ?? CHR(NUM2)
            ?? PADR("POBLACION",30," ")
            ?? CHR(NUM2)
            ?? PADR("PROVINCIA",30," ")
            ?? CHR(NUM2)
            ?? PADR("CIF/NIF",15," ")
            ?? CHR(NUM2)
            ?? PADR("TELEFONO",15," ")
            ?? CHR(NUM2)
            ?? PADR("FAX",15," ")
            ?? CHR(NUM2)
            ?? PADL("IMPORTE",12," ")
            ?? CHR(NUM2)
            ?? PADL("AÑO",4," ")
            ?? CHR(NUM2)
            ?? PADL("RET.",4," ")
            ?? CHR(NUM2)
            ?? PADL("DOC.",6," ")
            SET PRINT OFF
        ENDIF
        SET PRINT ON
        ? CODCTA
        ?? CHR(NUM2)
        ?? LEFT(NOMCTA,30)
        ?? CHR(NUM2)
        ?? DIRCTA
        ?? CHR(NUM2)
        ?? CODPOS
        ?? CHR(NUM2)
        ?? POBCTA
        ?? CHR(NUM2)
        ?? PROCTA
        ?? CHR(NUM2)
        ?? CIF
        ?? CHR(NUM2)
        ?? TEL1
        ?? CHR(NUM2)
        ?? FAX1
        ?? CHR(NUM2)
        ?? MIL(SALDO,12,MDA_DEC(EJERANY))
        ?? CHR(NUM2)
        ?? EJERANY
        ?? CHR(NUM2)
        ?? IF(LEFT(BANCTA,1)="*","RET.","    ")
        ?? IF(DEBE<>0,"COMPRA","VENTA ")
        SET PRINT OFF
        TOT1=TOT1+DEBE
        TOT2=TOT2+HABER
        SKIP
    ENDDO
    SET PRINTER TO
    VerFicRtf(RUTAPROGRAMA+"\MOD_347.TXT")

Return Nil

STATIC FUNCTION RecldeciA()
    STORE 0 TO PAG,TOT1,TOT2
    SET PRINTER TO &RUTAPROGRAMA\MOD_347.TXT
    SUM SALDO TO SUMTOT
    REGISTROS:=LASTREC()
    GO TOP
    DO WHILE .NOT. EOF()
        IF .NOT. EOF() .AND. PAG=0
            PAG=PAG+1
            AbrirDBF("EMPRESA",,,"Exclusive",RUTAPROGRAMA)
            SET PRINT ON
            ?? "1"
            ?? "347"
            ?? PADL(ALLTRIM(STR(YEAR(W_Imp1.D_Fec1.value))) ,4,"0")
            CIFEMP:=CIF
            ?? PADL(DNI_LET(CIFEMP,1) ,9,"0")
            ?? PADR(EMP,40," ")
            ?? "D"
            ?? PADR("0",9,"0")
            ?? SPACE(40)
            ?? PADL("1",13,"0")
            ?? SPACE(2)
            ?? PADL("0",13,"0")
            ?? PADL(ALLTRIM(STR(REGISTROS)) ,9,"0")
            SUMTOT:=MIL(SUMTOT,20,2)
            SUMTOT:=STRTRAN(SUMTOT, ",", "")
            SUMTOT:=STRTRAN(SUMTOT, ".", "")
            ?? PADL(ALLTRIM(SUMTOT),15,"0")
            ?? SPACE(9)
            ?? SPACE(15)
            ?? SPACE(54)
            ?? SPACE(13)
            ?
            SET PRINT OFF
            USE
            SELEC FIN
        ENDIF
        SET PRINT ON
        ?? "2"
        ?? "347"
        ?? PADL(ALLTRIM(STR(YEAR(W_Imp1.D_Fec1.value))) ,4,"0")
        ?? PADL(DNI_LET(CIFEMP,1) ,9,"0")
        ?? PADL(DNI_LET(CIF,1) ,9,"0")
        ?? SPACE(9)
        ?? PADR(NOMCTA,40," ")
        ?? "D"
        ?? LEFT(CODPOS,2)+PADL("0",3,"0")
        IF DEBE<>0
            ?? "A"
        ELSE
            ?? "B"
        ENDIF
        SALDO2:=MIL(SALDO,20,2)
        SALDO2:=STRTRAN(SALDO2, ",", "")
        SALDO2:=STRTRAN(SALDO2, ".", "")
        ?? PADL(ALLTRIM(SALDO2),15,"0")
        ?? " "
        ?? " "
        ?? SPACE(151)
        ?
        SET PRINT OFF
        TOT2=TOT2+SALDO
        SKIP
    ENDDO
    SET PRINTER TO
    VerFicRtf(RUTAPROGRAMA+"\MOD_347.TXT")

    Return Nil