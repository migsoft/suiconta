#include "minigui.ch"
#include "winprint.ch"

procedure Reclret()
    TituloImp:="Listado de facturas con retencion"

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 300 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

        @ 70,10 CHECKBOX SiFacEmi CAPTION 'Incluir facturas emitidas' WIDTH 155 VALUE .T.
        @100,10 CHECKBOX SiFacRec CAPTION 'Incluir facturas recibidas' WIDTH 155 VALUE .T.


        LINW:=130
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
            ACTION Reclreti()

        @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION Reclreti()
    aFacRet:={}
    AbrirDBF("FAC92")
    GO TOP
    DO WHILE .NOT. EOF()
        IF RET<>0 .AND. FFAC>=W_Imp1.D_Fec1.value .AND. FFAC<=W_Imp1.D_Fec2.value
            AADD(aFacRet,{"EMI",STR(ASIENTO,6),FFAC,CODCTA,CLIENTE, PADR(STR(NFAC,6)+"-"+SERFAC,15) ,BIMP,RET,IMPRET})
        ENDIF
        SKIP
    ENDDO
    AbrirDBF("FACREB")
    GO TOP
    DO WHILE .NOT. EOF()
        IF RET<>0 .AND. FREG>=W_Imp1.D_Fec1.value .AND. FREG<=W_Imp1.D_Fec2.value
            AADD(aFacRet,{"REC",STR(NREG,6),FREG,CODIGO,NOMCTA,REF,BIMP,RET,IMPRET})
        ENDIF
        SKIP
    ENDDO

**   INDEX ON NOMCTA+STRZERO(CODIGO,8)+DTOS(FREG) TO FIN
    ASORT(aFacRet,,,{|x,y| x[1]+x[5]+STR(x[4],8)+DTOS(x[3]) < y[1]+y[5]+STR(y[4],8)+DTOS(y[3]) })

    IF LEN(aFacRet)=0
        MsgExclamation("No hay datos en las fechas introducidas","Informacion")
        RETURN
    ENDIF

    Reclreti2()

RETURN Nil




STATIC FUNCTION Reclreti2()
    local oprint

    LOCAL TOT1:=0,TOT2:=0,TOTC1:=0,TOTC2:=0,NOMCTA2:=" ",CODCTA2:=0

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

    STORE 0 TO PAG,NUNE
    AbrirDBF("CUENTAS")
    LIN:=0
    FOR N=1 TO LEN(aFacRet)
        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN,125,"Suma","times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,165,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
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
            oprint:printdata(40,70,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

            LIN:=45
            oprint:printdata(LIN, 20,"Tipo Fac.","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 43,"Registro","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN, 44,"Fecha","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 70,"Codigo","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN, 72,"Nombre","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,125,"Nº factura","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,165,"Base imp.","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,175,"%","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,190,"Retencion","times new roman",8,.F.,,"R",)
            oprint:printline(LIN+4,20,LIN+4,190,,0.5)

            LIN:=LIN+5
        ENDIF
        IF NOMCTA2<>aFacRet[N,5] .OR. CODCTA2<>aFacRet[N,4]
            NOMCTA2:=aFacRet[N,5]
            CODCTA2:=aFacRet[N,4]
            AbrirDBF("CUENTAS")
            SEEK CODCTA2
            oprint:printdata(LIN, 20,CIF,"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 40,NOMCTA,"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,100,DIRCTA,"times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,150,CODPOS+"-"+LEFT(POBCTA,20),"times new roman",8,.F.,,"L",)
            LIN:=LIN+5
        ENDIF
        IF aFacRet[N,1]="EMI"
            oprint:printdata(LIN, 20,"Emitida","times new roman",8,.F.,,"L",)
        ELSE
            oprint:printdata(LIN, 20,"Recibida","times new roman",8,.F.,,"L",)
        ENDIF
        oprint:printdata(LIN, 43,aFacRet[N,2],"times new roman",8,.F.,,"R",)  //NREG
        oprint:printdata(LIN, 44,aFacRet[N,3],"times new roman",8,.F.,,"L",)  //FREG
        oprint:printdata(LIN, 70,aFacRet[N,4],"times new roman",8,.F.,,"R",)  //CODIGO
        oprint:printdata(LIN, 72,LEFT(aFacRet[N,5],30),"times new roman",8,.F.,,"L",)  //NOMCTA
        oprint:printdata(LIN,125,aFacRet[N,6],"times new roman",8,.F.,,"L",)  //REF
        oprint:printdata(LIN,165,MIL(aFacRet[N,7],12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN,175,MIL(aFacRet[N,8], 5,2),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN,190,MIL(aFacRet[N,9],12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
        TOT1 =TOT1 +aFacRet[N,7]                    //BIMP
        TOT2 =TOT2 +aFacRet[N,9]                    //IMPRET
        TOTC1=TOTC1+aFacRet[N,7]                    //BIMP
        TOTC2=TOTC2+aFacRet[N,9]                    //IMPRET
        SiTotal:=0
        IF N=LEN(aFacRet)
            SiTotal:=1
        ELSE
            IF NOMCTA2<>aFacRet[N+1,5] .OR. CODCTA2<>aFacRet[N+1,4]
                SiTotal:=1
            ENDIF
        ENDIF
        IF SiTotal=1
            LIN:=LIN+5
            oprint:printdata(LIN,125,"Total Cuenta","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,165,MIL(TOTC1,12,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,190,MIL(TOTC2,12,2),"times new roman",8,.F.,,"R",)
            TOTC1=0
            TOTC2=0
            LIN:=LIN+5
        ENDIF
        LIN:=LIN+5
    NEXT
    oprint:printdata(LIN,125,"Total","times new roman",8,.F.,,"L",)
    oprint:printdata(LIN,165,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
    oprint:printdata(LIN,190,MIL(TOT2,12,2),"times new roman",8,.F.,,"R",)


    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return Nil