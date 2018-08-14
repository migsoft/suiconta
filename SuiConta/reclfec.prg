#include "minigui.ch"
#include "winprint.ch"

procedure Reclfec()
    TituloImp:="Listado de facturas recibidas por fecha"

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 380 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @ 15,10 LABEL L_CodCta1 VALUE 'Desde el codigo' AUTOSIZE TRANSPARENT
        @ 10,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
            ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
        @ 10,225 BUTTONEX Bt_BuscarCue1 ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") , ;
            W_Imp1.T_CodCta2.Value:=W_Imp1.T_CodCta1.Value )  ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @ 45,10 LABEL L_CodCta2 VALUE 'Hasta el codigo' AUTOSIZE TRANSPARENT
        @ 40,110 TEXTBOX T_CodCta2 WIDTH 100 VALUE "99999999" MAXLENGTH 8 ;
            ON LOSTFOCUS W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta2.Value)
        @ 40,225 BUTTONEX Bt_BuscarCue2 ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION br_cue1(VAL(W_Imp1.T_CodCta2.Value),"W_Imp1","T_CodCta2") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP


        @ 75,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
        @ 70,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
        @ 75,215 LABEL L_Fec1b VALUE 'Año = ejercicios anteriores' AUTOSIZE TRANSPARENT

        @105,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @100,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2
        @105,215 LABEL L_Fec2b VALUE 'Año = ejercicios posteriores' AUTOSIZE TRANSPARENT

        @135,10 LABEL L_OrdLis ;
            VALUE 'Ordenar por:' ;
            WIDTH 90 HEIGHT 25
        @130,110 COMBOBOX C_OrdLis ;
            WIDTH 150 ;
            ITEMS {'Registro entrada','Fecha entrada','Codigo cuenta','Nombre cuenta'} ;
            VALUE 1

        @165,10 LABEL L_Mostrar VALUE 'Mostrar' AUTOSIZE TRANSPARENT
        @160,110 RADIOGROUP R_Mostrar OPTIONS {"Base imponible y Total factura","Total factura y Pendiente"} ;
            VALUE 1 WIDTH 200

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
            ACTION Reclfeci()

        @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Reclfeci()
    local oprint

    IF FILE("FIN1.DBF")
        AbrirDBF("Fin1","SIN_INDICE")
        FIN1->( DBCLOSEAREA() )
        ERASE FIN1.DBF
        ERASE FIN1.CDX
    ENDIF

    AbrirDBF("FACREB")
    COPY STRUC TO FIN1
    FACREB->( DBCLOSEAREA() )
    AbrirDBF("Fin1","SIN_INDICE")

    FOR ANY2=YEAR(W_Imp1.D_Fec1.value) TO YEAR(W_Imp1.D_Fec2.value)
        RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
        RUTA2:=RUTA1[1]+"\FACREB.DBF"
        IF .NOT. FILE(RUTA2)
            LOOP
        ENDIF
        APPEND FROM &RUTA2 FOR ;
            CODIGO>=VAL(W_Imp1.T_CodCta1.Value) .AND. CODIGO<=VAL(W_Imp1.T_CodCta2.Value) .AND. ;
            FREG>=W_Imp1.D_Fec1.value .AND. FREG<=W_Imp1.D_Fec2.value
    NEXT

    DO CASE
    CASE W_Imp1.C_OrdLis.value=1
        INDEX ON STR(YEAR(FREG))+STR(NREG) TO FIN1
    CASE W_Imp1.C_OrdLis.value=2
        INDEX ON DTOS(FREG)+STR(YEAR(FREG))+STR(NREG) TO FIN1
    CASE W_Imp1.C_OrdLis.value=3
        INDEX ON STR(CODIGO,10)+STR(YEAR(FREG))+STR(NREG) TO FIN1
    CASE W_Imp1.C_OrdLis.value=4
        INDEX ON NOMCTA+STR(YEAR(FREG))+STR(NREG) TO FIN1
    ENDCASE

    SELEC FIN1
    GO TOP
    IF LASTREC()=0
        MsgExclamation("No hay datos en las fechas introducidas","Informacion")
        FIN1->( DBCLOSEAREA() )
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

    PAG:=0
    LIN:=0
    TOT1:=0
    TOT2:=0
    DO WHILE .NOT. EOF()
        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                oprint:printdata(LIN,100,"Suma","times new roman",8,.F.,,"L",)
                oprint:printdata(LIN,145,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,165,MIL(TOT2,12,2),"times new roman",8,.F.,,"R",)
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

            oprint:printdata(40,20,'Desde: '+DIA(W_Imp1.D_Fec1.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,20,'Hasta: '+DIA(W_Imp1.D_Fec2.value,10),"times new roman",12,.F.,,"L",)

            oprint:printdata(40,80,'Desde: '+W_Imp1.T_CodCta1.Value,"times new roman",12,.F.,,"L",)
            oprint:printdata(45,80,'Hasta: '+W_Imp1.T_CodCta2.Value,"times new roman",12,.F.,,"L",)

            oprint:printdata(40,150,'Ordenado por:',"times new roman",12,.F.,,"L",)
            oprint:printdata(45,150,W_Imp1.C_OrdLis.DisplayValue,"times new roman",12,.F.,,"L",)

            LIN:=55
            oprint:printdata(LIN, 30,"Registro","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN, 32,"Fecha","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN, 58,"Codigo","times new roman",8,.F.,,"R",)
            oprint:printdata(LIN, 60,"Nombre","times new roman",8,.F.,,"L",)
            oprint:printdata(LIN,110,"Factura","times new roman",8,.F.,,"L",)
            IF W_Imp1.R_Mostrar.Value=1
                oprint:printdata(LIN,145,"Base imp.","times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,165,"Total fac.","times new roman",8,.F.,,"R",)
            ELSE
                oprint:printdata(LIN,145,"Total fac.","times new roman",8,.F.,,"R",)
                oprint:printdata(LIN,165,"Pendiente","times new roman",8,.F.,,"R",)
            ENDIF
            oprint:printdata(LIN,167,"Forma pago","times new roman",8,.F.,,"L",)
            oprint:printline(LIN+4,20,LIN+4,180,,0.5)

            LIN:=LIN+5
        ENDIF

        oprint:printdata(LIN, 30,MIL(NREG,10,0),"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN, 32,DIA(FREG,10),"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN, 58,CODIGO,"times new roman",8,.F.,,"R",)
        oprint:printdata(LIN, 60,LEFT(NOMCTA,30),"times new roman",8,.F.,,"L",)
        oprint:printdata(LIN,110,REF,"times new roman",8,.F.,,"L",)
        IF W_Imp1.R_Mostrar.Value=1
            oprint:printdata(LIN,145,MIL(BIMP+BIMPT2+BIMPT3,12,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,165,MIL(TFAC,12,2),"times new roman",8,.F.,,"R",)
            TOT1:=TOT1+BIMP+BIMPT2+BIMPT3
            TOT2:=TOT2+TFAC
        ELSE
            oprint:printdata(LIN,145,MIL(TFAC,12,2),"times new roman",8,.F.,,"R",)
            oprint:printdata(LIN,165,MIL(PEND,12,2),"times new roman",8,.F.,,"R",)
            TOT1:=TOT1+TFAC
            TOT2:=TOT2+PEND
        ENDIF
        oprint:printdata(LIN,167,VENCI_NC(VENCIR(NEMP,NREG,FREG),18),"times new roman",8,.F.,,"L",)

        LIN:=LIN+5
        SKIP

    ENDDO

    oprint:printdata(LIN,100,"Total","times new roman",8,.F.,,"L",)
    oprint:printdata(LIN,145,MIL(TOT1,12,2),"times new roman",8,.F.,,"R",)
    oprint:printdata(LIN,165,MIL(TOT2,12,2),"times new roman",8,.F.,,"R",)

    SELEC FIN1
    FIN1->( DBCLOSEAREA() )

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return Nil