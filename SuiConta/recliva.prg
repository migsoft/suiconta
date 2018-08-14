#include "minigui.ch"
#include "winprint.ch"

procedure Recliva()
    TituloImp:="Listado de facturas recibidas IVA"

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 390 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

        aREGIMEN:={"99-TODOS LOS TIPOS DE REGIMEN"}
        FOR NN=0 TO 4
            AADD(aREGIMEN,LTRIM(STR(NN))+"-"+REGIVAREC(NN))
        NEXT
        @ 75,10 LABEL L_Regimen VALUE 'Regimen' AUTOSIZE TRANSPARENT
        @ 70,110 COMBOBOX C_Regimen WIDTH 250 ITEMS aREGIMEN VALUE 1

        @105,10 LABEL L_TipLis VALUE 'Tipo de listado' AUTOSIZE TRANSPARENT
        @100,110 RADIOGROUP R_TipLis OPTIONS { 'Listado' , 'Libro IVA' } ;
            VALUE 1 WIDTH 75 HORIZONTAL

        @135,10 LABEL L_FecLis VALUE 'Fecha listado' AUTOSIZE TRANSPARENT
        @130,110 DATEPICKER D_FecLis WIDTH 100 VALUE DATE()

        @165,10 LABEL L_PagIni VALUE 'Numero de pagina' AUTOSIZE TRANSPARENT
        @160,110 TEXTBOX T_PagIni WIDTH 100 HEIGHT 25 VALUE 1 ;
            NUMERIC INPUTMASK '99,999,999,999' FORMAT 'E' RIGHTALIGN

        @195,10 LABEL L_NumIni VALUE 'Registro inicial' AUTOSIZE TRANSPARENT
        @190,110 TEXTBOX T_NumIni WIDTH 100 HEIGHT 25 VALUE 1 ;
            NUMERIC INPUTMASK '99,999,999,999' FORMAT 'E' RIGHTALIGN

        LINW:=220
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
            ACTION Reclivai()

        @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil

STATIC FUNCTION Reclivai()
    IF W_Imp1.C_Regimen.value<=1
        RESPREG2:=99
    ELSE
        RESPREG2:=W_Imp1.C_Regimen.value-2
    ENDIF

    IF FILE("FIN.DBF")
        AbrirDBF("FIN","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    AbrirDBF("FACREB")
    IF RESPREG2=99
        COPY TO FIN FOR FREG>=W_Imp1.D_Fec1.value .AND. FREG<=W_Imp1.D_Fec2.value
    ELSE
        COPY TO FIN FOR FREG>=W_Imp1.D_Fec1.value .AND. FREG<=W_Imp1.D_Fec2.value .AND. REGIMEN=RESPREG2
    ENDIF
    FACREB->( DBCLOSEAREA() )
    AbrirDBF("FIN","SIN_INDICE")
    INDEX ON STR(REGIMEN)+DTOS(FREG)+STR(NREG) TO FIN

    IF FIN->(LASTREC())=0
        MsgExclamation("No hay datos en las fechas introducidas","Informacion")
        FIN->( DBCLOSEAREA() )
        RETURN
    ENDIF

    Reclivai2()

RETURN Nil


STATIC FUNCTION Reclivai2()
    local oprint

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",4)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(TituloImp)
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    MIVA:={}
    PAG=0
    LIN:=0
    REGIMEN3:=-1
    TOTR:=0
    TOTB:=0
    TOTC:=0
    TOTD:=0
    NUMINI:=W_Imp1.T_NumIni.Value

    DO WHILE .NOT. EOF()
        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN,110,"Sumas","times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,140,MIL(TOTB,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,170,MIL(TOTC,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
                IF TOTR<>0
                    LIN:=LIN+4
                    oprint:printdata(LIN,140,"Retencion","times new roman",8,.F.,,"R",)
                    oprint:printdata(LIN,170,MIL(TOTR,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
                ENDIF
                oprint:printdata(LIN,190,MIL(TOTD,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
                LIN:=LIN+4
                oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1+W_Imp1.T_PagIni.Value-1)),"times new roman",8,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG+W_Imp1.T_PagIni.value-1)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
            oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

            oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(40,60,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

            IF RESPREG2<>99
                oprint:printdata(45,20,'Regimen:',"times new roman",12,.F.,,"L",)
                oprint:printdata(45,50,MIL(RESPREG2,3,0)+"-"+REGIVAREC(RESPREG2),"times new roman",12,.F.,,"L",)
            ENDIF

            LIN:=50
            oprint:printdata(LIN, 25,"Factura","times new roman",8,.T.,,"R",)
            oprint:printdata(LIN, 27,"Fecha","times new roman",8,.T.,,"L",)
            oprint:printdata(LIN, 50,"Codigo","times new roman",8,.T.,,"R",)
            oprint:printdata(LIN, 52,"Cliente","times new roman",8,.T.,,"L",)
            oprint:printdata(LIN,100,"Referencia","times new roman",8,.T.,,"L",)
            oprint:printdata(LIN,140,"Base imp.","times new roman",8,.T.,,"R",)
            oprint:printdata(LIN,150,"%IVA","times new roman",8,.T.,,"R",)
            oprint:printdata(LIN,170,"Cuota","times new roman",8,.T.,,"R",)
            oprint:printdata(LIN,190,"Total fac.","times new roman",8,.T.,,"R",)
            IF W_Imp1.R_TipLis.Value=2
                oprint:printdata(LIN,200,"Reg.","times new roman",8,.F.,,"R",)
            ENDIF
            oprint:printline(LIN+4,15,LIN+4,200,,0.5)

            LIN:=LIN+4
        ENDIF

        IF REGIMEN3<>REGIMEN
            REGIMEN3:=REGIMEN
            IF REGIMEN3<2000
                oprint:printdata(LIN,105,MIL(REGIMEN3,3,0)+"-"+REGIVAREC(REGIMEN3),"times new roman",12,.F.,,"C",)
            ELSE
                oprint:printdata(LIN,105,MIL(REGIMEN3-2000,3,0)+"-"+REGIVAREC(REGIMEN3-2000),"times new roman",12,.F.,,"C",)
            ENDIF
            LIN:=LIN+5
        ENDIF


        IF W_Imp1.R_TipLis.Value=2
            oprint:printdata(LIN, 25,STR(NUMINI++,5),"times new roman",8,.F.,,"R",)
        ELSE
            oprint:printdata(LIN, 25,NREG,"times new roman",8,.F.,,"R",)
        ENDIF
        oprint:printdata(LIN, 27,DIA(FREG,8),"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN, 50,CODIGO,"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN, 52,LEFT(NOMCTA,30),"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,100,REF,"times new roman",8,.F.,,"L",)

        oprint:printdata(LIN,140,MIL(BIMP,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN,150,MIL(IVA,5,2),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN,170,MIL(CUOTA,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
        TIPIVA2:=(REGIMEN*1000)+IVA
        NUM2:=ASCAN(MIVA,{|AVAL| AVAL[1]=TIPIVA2})
        IF NUM2=0
            AADD(MIVA,{TIPIVA2,REGIMEN,IVA,BIMP,CUOTA,"IVA"})
        ELSE
            MIVA[NUM2,4]=MIVA[NUM2,4]+BIMP
            MIVA[NUM2,5]=MIVA[NUM2,5]+CUOTA
        ENDIF

        IF BIMPT2<>0
            LIN:=LIN+4
            oprint:printdata(LIN,140,MIL(BIMPT2,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,150,MIL(IVAT2,5,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,170,MIL(CUOTAT2,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            TIPIVA2:=(REGIMEN*1000)+IVAT2
            NUM2:=ASCAN(MIVA,{|AVAL| AVAL[1]=TIPIVA2})
            IF NUM2=0
                AADD(MIVA,{TIPIVA2,REGIMEN,IVAT2,BIMPT2,CUOTAT2,"IVA"})
            ELSE
                MIVA[NUM2,4]=MIVA[NUM2,4]+BIMPT2
                MIVA[NUM2,5]=MIVA[NUM2,5]+CUOTAT2
            ENDIF
        ENDIF

        IF BIMPT3<>0
            LIN:=LIN+4
            oprint:printdata(LIN,140,MIL(BIMPT3,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,150,MIL(IVAT3,5,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,170,MIL(CUOTAT3,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            TIPIVA2:=(REGIMEN*1000)+IVAT3
            NUM2:=ASCAN(MIVA,{|AVAL| AVAL[1]=TIPIVA2})
            IF NUM2=0
                AADD(MIVA,{TIPIVA2,REGIMEN,IVAT3,BIMPT3,CUOTAT3,"IVA"})
            ELSE
                MIVA[NUM2,4]=MIVA[NUM2,4]+BIMPT3
                MIVA[NUM2,5]=MIVA[NUM2,5]+CUOTAT3
            ENDIF
        ENDIF

        IF RET<>0
            LIN:=LIN+4
            oprint:printdata(LIN,140,"Retencion","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,150,MIL(RET,5,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,170,MIL(IMPRET,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            TOTR:=TOTR+IMPRET
        ENDIF

        IF REQ<>0
            LIN:=LIN+4
            oprint:printdata(LIN,140,"Recargo","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,150,MIL(REQ,5,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,170,MIL(IMPREQ,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
            TIPIVA2:=(REGIMEN*1000)+REQ+800
            NUM2:=ASCAN(MIVA,{|AVAL| AVAL[1]=TIPIVA2})
            IF NUM2=0
                AADD(MIVA,{TIPIVA2,REGIMEN,REQ,BIMP+BIMPT2+BIMPT3,IMPREQ,"REQ"})
            ELSE
                MIVA[NUM2,4]=MIVA[NUM2,4]+BIMP+BIMPT2+BIMPT3
                MIVA[NUM2,5]=MIVA[NUM2,5]+IMPREQ
            ENDIF
        ENDIF

        oprint:printdata(LIN,190,MIL(TFAC,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)

        IF W_Imp1.R_TipLis.Value=2
            oprint:printdata(LIN,200,NREG,"times new roman",8,.F.,,"R",)
        ENDIF

        TOTB=TOTB+BIMP+BIMPT2+BIMPT3
        TOTC=TOTC+CUOTA+CUOTAT2+CUOTAT3+IMPREQ
        TOTD=TOTD+TFAC

        LIN:=LIN+4
        SKIP
    ENDDO

    oprint:printdata(LIN,110,"Total","times new roman",8,.F.,,"R",)
    oprint:printdata(LIN,140,MIL(TOTB,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
    oprint:printdata(LIN,170,MIL(TOTC,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
    IF TOTR<>0
        LIN:=LIN+4
        oprint:printdata(LIN,140,"Retencion","times new roman",8,.F.,,"R",)
        oprint:printdata(LIN,170,MIL(TOTR,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)
    ENDIF
    oprint:printdata(LIN,190,MIL(TOTD,12,MDA_DEC(EJERANY)),"times new roman",8,.F.,,"R",)

********* IVAS TOTALES
    IF LIN+LEN(MIVA)>260
        LIN:=LIN+4
        oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1+W_Imp1.T_PagIni.Value-1)),"times new roman",8,.F.,,"C",)
        oprint:endpage()
        oprint:beginpage()

        PAG=PAG+1
        oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
        oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG+W_Imp1.T_PagIni.value-1)),"times new roman",12,.F.,,"R",)
        oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

        oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
        oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

        oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
        oprint:printdata(40,60,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

        IF RESPREG2<>99
            oprint:printdata(45,20,'Regimen:',"times new roman",12,.F.,,"L",)
            oprint:printdata(45,50,MIL(RESPREG2,3,0)+"-"+REGIVAREC(RESPREG2),"times new roman",12,.F.,,"L",)
        ENDIF

        LIN:=50
    ENDIF

    LIN:=LIN+15

    oprint:printdata(LIN,30,"Regimen","times new roman",10,.T.,,"L",)
    oprint:printdata(LIN,75,"Tipo IVA","times new roman",10,.T.,,"R",)
    oprint:printdata(LIN,110,"Base Imp.","times new roman",10,.T.,,"R",)
    oprint:printdata(LIN,150,"Cuota","times new roman",10,.T.,,"R",)
    oprint:printline(LIN+4,30,LIN+4,150,,0.5)

    LIN:=LIN+5

    ASORT(MIVA,,,{|X,Y| X[1]<Y[1]})
    TOT1:=0
    REGIMEN3:=-1
    FOR N=1 TO LEN(MIVA)
        IF REGIMEN3<>MIVA[N,2]
            REGIMEN3:=MIVA[N,2]
            IF REGIMEN3<2000
                oprint:printdata(LIN,30,MIL(REGIMEN3,3,0)+"-"+REGIVAREC(REGIMEN3),"times new roman",12,.F.,,"L",)
            ELSE
                oprint:printdata(LIN,30,MIL(REGIMEN3-2000,3,0)+"-"+REGIVAREC(REGIMEN3-2000),"times new roman",12,.F.,,"L",)
            ENDIF
            LIN:=LIN+5
        ENDIF
        IF MIVA[N,6]="REQ"
            oprint:printdata(LIN,50,"Recargo","times new roman",10,.F.,,"L",)
        ENDIF
        oprint:printdata(LIN, 75,MIL(MIVA[N,3],5,2)+"%","times new roman",10,.F.,,"R",)
        oprint:printdata(LIN,110,MIL(MIVA[N,4],12,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
        oprint:printdata(LIN,150,MIL(MIVA[N,5],12,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)
        TOT1=TOT1+MIVA[N,5]
        LIN:=LIN+5
    NEXT

    oprint:printdata(LIN,110,"Total","times new roman",10,.F.,,"L",)
    oprint:printdata(LIN,150,MIL(TOT1,12,MDA_DEC(EJERANY)),"times new roman",10,.F.,,"R",)

    SELEC FIN
    FIN->( DBCLOSEAREA() )

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return Nil