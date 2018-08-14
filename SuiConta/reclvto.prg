#include "minigui.ch"
#include "winprint.ch"

procedure Reclvto()
    TituloImp:="Cuadro prevision de pagos"

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 510 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()


        @ 15,10 LABEL L_Tipo VALUE 'Tipo de listado' AUTOSIZE TRANSPARENT
        @ 10,110 RADIOGROUP R_Tipo OPTIONS { 'Cuadro de vencimientos' , 'Listado de vencimientos' } ;
            VALUE 1 WIDTH 160 ON CHANGE Reclvto_act()

        @ 75,10 LABEL L_Vto1 VALUE 'Desde vencimiento' AUTOSIZE TRANSPARENT
        @ 70,120 DATEPICKER D_Vto1 WIDTH 100 VALUE DIAINIMES(DATE())
        @105,10 LABEL L_Vto2 VALUE 'Hasta vencimiento' AUTOSIZE TRANSPARENT
        @100,120 DATEPICKER D_Vto2 WIDTH 100 VALUE DIAINIMES(DATE())+365

        @130,10 BUTTONEX Bt_BuscarCue1 CAPTION 'Banco' ICON icobus('buscar') ;
            ACTION br_cue1(VAL(W_Imp1.T_CodBan1.Value),"W_Imp1","T_CodBan1","T_NomBan1") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @130,110 TEXTBOX T_CodBan1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
            ON LOSTFOCUS W_Imp1.T_CodBan1.Value:=PCODCTA3(W_Imp1.T_CodBan1.Value)
        @130,220 TEXTBOX T_NomBan1 WIDTH 150 VALUE "" READONLY

        @160,10 BUTTONEX Bt_BuscarCta1 CAPTION 'Cuenta' ICON icobus('buscar') ;
            ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1","T_NomCta1") ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @160,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
            ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
        @160,220 TEXTBOX T_NomCta1 WIDTH 150 VALUE "" READONLY


        @195,10 LABEL L_SiVto VALUE 'Listar' AUTOSIZE TRANSPARENT
        @190,110 RADIOGROUP R_SiVto OPTIONS { 'Todas las facturas' , 'Solo con vencimiento' } ;
            VALUE 1 WIDTH 160

        @255,10 LABEL L_Orden VALUE 'Ordenado por' AUTOSIZE TRANSPARENT
        @250,110 RADIOGROUP R_Orden OPTIONS { 'Fecha vencimiento' , 'Nombre del proveedor' } ;
            VALUE 1 WIDTH 160

        @310,10 LABEL L_Progres VALUE 'Progreso' AUTOSIZE TRANSPARENT
        @310,110 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 250 HEIGHT 20 SMOOTH

        draw rectangle in window W_Imp1 at 350,010 to 352,390 fillcolor{255,0,0}  //Rojo
        aIMP:=Impresoras(EMP_IMPRESORA)
        @365,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @360,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @395,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @390,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @390, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 155 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @420, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 155 value .f.

        @450, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION Reclvtoi1()

        @450,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION Reclvto_act()
    IF W_Imp1.R_Tipo.value=1
        W_Imp1.D_Vto1.value:=DIAINIMES(DATE())
        W_Imp1.D_Vto2.value:=DIAFINMES(DATE())+365
    ELSE
        W_Imp1.D_Vto1.value:=DIAINIMES(DIAMESMAS(DATE(),1))
        W_Imp1.D_Vto2.value:=DIAFINMES(DIAMESMAS(DATE(),1))
    ENDIF
Return Nil

STATIC FUNCTION Reclvtoi1()

    IF FILE("FIN.DBF")
        AbrirDBF("Fin","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    FICBASE:={{'NREG        ','N',         5,         0},;
        {'FREG        ','D',        10,         0},;
        {'CODCTA      ','N',         8,         0},;
        {'NOMCTA      ','C',        30,         0},;
        {'REF         ','C',        15,         0},;
        {'TFAC        ','N',        14,         3},;
        {'TVTO        ','N',        14,         3},;
        {'FVTO        ','D',        10,         0},;
        {'BANCO       ','N',         8,         0},;
        {'PEND        ','L',         1,         0},;
        {'ANYFAC      ','N',         4,         0}}
    DBCREATE("FIN.DBF",FICBASE)
    AbrirDBF("FIN","SIN_INDICE",,"Exclusive")


*** INCLUIR EJERCICIOS ANTORIORES ***
    FOR ANY2=EJERANY-1 TO EJERANY+1
        W_Imp1.L_Progres.Value:=STR(ANY2,4)
        RUTA2:=EMPRUTANY(NUMEMP,ANY2)
        IF FILE(RUTA2+"\FACREB.DBF") .AND. FILE(RUTA2+"\FACREB.CDX") .AND. ;
            FILE(RUTA2+"\FACVTO.DBF") .AND. FILE(RUTA2+"\FACVTO.CDX")
            W_Imp1.P_Progres.RangeMax:=LASTREC()
            W_Imp1.P_Progres.Value:=0
            AbrirDBF("FACVTO",,,,RUTA2)
            AbrirDBF("FACREB",,,,RUTA2)
            GO TOP
            DO WHILE .NOT. EOF()
                W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
                IF VAL(W_Imp1.T_CodCta1.Value)<>0 .AND. CODIGO<>VAL(W_Imp1.T_CodCta1.Value)
                    SKIP
                    LOOP
                ENDIF
                NREG2:=NREG
                FREG2:=FREG
                CODCTA2:=CODIGO
                NOMCTA2:=NOMCTA
                REF2:=REF
                TFAC2:=TFAC
                PEND2:=PEND
                IMPVTO2:=0

                SELEC FACVTO
                SEEK NREG2

                IF EOF()
                    IF W_Imp1.R_SiVto.value=1
                        SELEC FIN
                        APPEND BLANK
                        REPLACE NREG WITH NREG2
                        REPLACE FREG WITH FREG2
                        REPLACE CODCTA WITH CODCTA2
                        REPLACE NOMCTA WITH NOMCTA2
                        REPLACE REF WITH REF2
                        REPLACE TFAC WITH TFAC2
                        REPLACE TVTO WITH TFAC2
                        REPLACE FVTO WITH FREG2
                        REPLACE BANCO WITH 0
                        REPLACE ANYFAC WITH ANY2
                        IF TFAC2>0
                            REPLACE PEND WITH IF(PEND2>0, .T. , .F. )
                        ELSE
                            REPLACE PEND WITH IF(PEND2<0, .T. , .F. )
                        ENDIF
                        SELEC FACREB
                    ENDIF
                ELSE
                    DO WHILE NREG=NREG2 .AND. .NOT. EOF()
                        FVTO2:=FVTO
                        TVTO2:=IMPORTE
                        BANCO2:=BANCO
                        IMPVTO2:=IMPVTO2+TVTO2
                        IF FVTO2>=W_Imp1.D_Vto1.value .AND. FVTO2<=W_Imp1.D_Vto2.value
                            SELEC FIN
                            APPEND BLANK
                            REPLACE NREG WITH NREG2
                            REPLACE FREG WITH FREG2
                            REPLACE CODCTA WITH CODCTA2
                            REPLACE NOMCTA WITH NOMCTA2
                            REPLACE REF WITH REF2
                            REPLACE TFAC WITH TFAC2
                            REPLACE TVTO WITH TVTO2
                            REPLACE FVTO WITH FVTO2
                            REPLACE BANCO WITH BANCO2
                            REPLACE ANYFAC WITH ANY2
                            IF TFAC2>0
                                REPLACE PEND WITH IF(PEND2>0, .T. , .F. )
                            ELSE
                                REPLACE PEND WITH IF(PEND2<0, .T. , .F. )
                            ENDIF
                        ENDIF
                        SELEC FACVTO
                        SKIP
                    ENDDO
                ENDIF
                SELEC FACREB
                SKIP
            ENDDO
        ENDIF
    NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***

    SELEC FIN

***BORRAR BANCOS***
    IF VAL(W_Imp1.T_CodBan1.Value)<>0
        DELETE FOR BANCO<>VAL(W_Imp1.T_CodBan1.Value)
        PACK
    ENDIF
***FIN BORRAR BANCOS***

    IF W_Imp1.R_Orden.value=1
        INDEX ON DTOS(FVTO)+UPPER(NOMCTA)+STR(NREG) TO FIN
    ELSE
        INDEX ON UPPER(NOMCTA)+DTOS(FVTO)+STR(NREG) TO FIN
    ENDIF

    GO TOP
    IF LASTREC()=0
        MsgExclamation("No hay datos en los parametros introducidos","Informacion")
        RETURN
    ENDIF

    Reclvtoi2()

Return Nil



STATIC FUNCTION Reclvtoi2()
    local oprint

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",5)
    HOJAHORZ2:=IF(W_Imp1.R_Tipo.value=1,.T.,.F.)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , HOJAHORZ2 , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(TituloImp)
    oprint:setpreviewsize(W_Imp1.R_Tipo.value)      // tamaño del preview
    oprint:beginpage()

    W_Imp1.P_Progres.RangeMax:=LASTREC()
    W_Imp1.P_Progres.Value:=0
    aFECVTO:={DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,0)) , ;
        DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,1)) , ;
        DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,2)) , ;
        DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,3)) , ;
        DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,4)) }
    PAG:=0
    LIN:=0
    TOT1:=0
    TOT2:=0
    TOT3:=0
    TOT4:=0
    TOT5:=0
    aCOL:={}
    IF W_Imp1.R_Tipo.value=1
        AADD(aCOL, 10)                              //tamaño letra
        AADD(aCOL,297)                              //ancho pagina
        AADD(aCOL,210)                              //alto pagina
        AADD(aCOL, 35)
        AADD(aCOL, 37)
        AADD(aCOL, 65)
        AADD(aCOL, 67)
        AADD(aCOL,130)
        AADD(aCOL,170)
        AADD(aCOL,190)
        AADD(aCOL,210)
        AADD(aCOL,230)
        AADD(aCOL,250)
        AADD(aCOL,252)
        AADD(aCOL,270)
    ELSE
        AADD(aCOL, 10)                              //tamaño letra
        AADD(aCOL,210)                              //ancho pagina
        AADD(aCOL,297)                              //alto pagina
        AADD(aCOL, 35)
        AADD(aCOL, 37)
        AADD(aCOL, 65)
        AADD(aCOL, 67)
        AADD(aCOL,130)
        AADD(aCOL,170)
        AADD(aCOL, 10)
        AADD(aCOL, 10)
        AADD(aCOL, 10)
        AADD(aCOL, 10)
        AADD(aCOL,172)
        AADD(aCOL,190)
    ENDIF
    DO WHILE .NOT. EOF()
        W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
        DO EVENTS

        IF LIN>=aCOL[3]-35 .OR. PAG=0 .OR. SALTA2=1
            SALTA2:=0
            IF PAG<>0
                oprint:printline(LIN-1,aCOL[9]-40,LIN-1,aCOL[15],,0.5)
                oprint:printdata(LIN,aCOL[9]-40,"Sumas","times new roman",aCOL[1],.F.,,"L",)
                oprint:printdata(LIN,aCOL[9],MIL(TOT1,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                IF W_Imp1.R_Tipo.value=1
                    oprint:printdata(LIN,aCOL[10],MIL(TOT2,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                    oprint:printdata(LIN,aCOL[11],MIL(TOT3,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                    oprint:printdata(LIN,aCOL[12],MIL(TOT4,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                    oprint:printdata(LIN,aCOL[13],MIL(TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                    oprint:printdata(LIN,aCOL[15],MIL(TOT1+TOT2+TOT3+TOT4+TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
                ENDIF
                LIN:=LIN+10
                oprint:printdata(LIN,aCOL[2]/2,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,aCOL[2]-20,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,aCOL[2]/2,NOMEMPRESA,"times new roman",12,.F.,,"C",)
            oprint:printdata(32,aCOL[2]/2,TituloImp,"times new roman",18,.F.,,"C",)

            oprint:printdata(40,20,'Desde fecha: '+DIA(W_Imp1.D_Vto1.value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,20,'Hasta fecha: '+DIA(W_Imp1.D_Vto2.value,10),"times new roman",12,.F.,,"L",)

            LIN:=55
            oprint:printdata(LIN,aCOL[4],"Registro","times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[5],"Fecha","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[6],"Factura","times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[7],"Proveedor","times new roman",aCOL[1],.F.,,"L",)
            oprint:printdata(LIN,aCOL[8],"Fecha vto.","times new roman",aCOL[1],.F.,,"L",)

            IF W_Imp1.R_Tipo.value=1
                oprint:printdata(LIN,aCOL[09],MES(MONTH(aFECVTO[1])),"times new roman",aCOL[1],.F.,,"R",)
                oprint:printdata(LIN,aCOL[10],MES(MONTH(aFECVTO[2])),"times new roman",aCOL[1],.F.,,"R",)
                oprint:printdata(LIN,aCOL[11],MES(MONTH(aFECVTO[3])),"times new roman",aCOL[1],.F.,,"R",)
                oprint:printdata(LIN,aCOL[12],MES(MONTH(aFECVTO[4])),"times new roman",aCOL[1],.F.,,"R",)
                oprint:printdata(LIN,aCOL[13],"Posterior","times new roman",aCOL[1],.F.,,"R",)
            ELSE
                oprint:printdata(LIN,aCOL[9],"Importe","times new roman",aCOL[1],.F.,,"R",)
            ENDIF

            oprint:printdata(LIN,aCOL[14],"Banco","times new roman",aCOL[1],.F.,,"L",)
            oprint:printline(LIN+4,15,LIN+4,aCOL[15],,0.5)

            LIN:=LIN+5
        ENDIF

        oprint:printdata(LIN,aCOL[4],NREG,"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[5],DIA(FREG,8),"times new roman",aCOL[1],.F.,,"L",)
        oprint:printdata(LIN,aCOL[6],RTRIM(LEFT(REF,10)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[7],NOMCTA,"times new roman",aCOL[1],.F.,,"L",)
        oprint:printdata(LIN,aCOL[8],DIA(FVTO,8),"times new roman",aCOL[1],.F.,,"L",)
        DO CASE
        CASE FVTO<aFECVTO[2] .OR. W_Imp1.R_Tipo.value=2
            oprint:printdata(LIN,aCOL[9],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[9],IF(PEND=.T.,"","p"),"times new roman",aCOL[1],.F.,,"L",)
            TOT1=TOT1+TVTO
        CASE FVTO>=aFECVTO[2] .AND. FVTO<aFECVTO[3]
            oprint:printdata(LIN,aCOL[10],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[10],IF(PEND=.T.,"","p"),"times new roman",aCOL[1],.F.,,"L",)
            TOT2=TOT2+TVTO
        CASE FVTO>=aFECVTO[3] .AND. FVTO<aFECVTO[4]
            oprint:printdata(LIN,aCOL[11],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[11],IF(PEND=.T.,"","p"),"times new roman",aCOL[1],.F.,,"L",)
            TOT3=TOT3+TVTO
        CASE FVTO>=aFECVTO[4] .AND. FVTO<aFECVTO[5]
            oprint:printdata(LIN,aCOL[12],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[12],IF(PEND=.T.,"","p"),"times new roman",aCOL[1],.F.,,"L",)
            TOT4=TOT4+TVTO
        OTHERWISE
            oprint:printdata(LIN,aCOL[13],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[13],IF(PEND=.T.,"","p"),"times new roman",aCOL[1],.F.,,"L",)
            TOT5=TOT5+TVTO
        ENDCASE

        oprint:printdata(LIN,aCOL[14],BANCO,"times new roman",aCOL[1],.F.,,"L",)

        LIN:=LIN+5
        SKIP

    ENDDO

    oprint:printline(LIN-1,aCOL[9]-40,LIN-1,aCOL[15],,0.5)
    oprint:printdata(LIN,aCOL[9]-40,"Total","times new roman",aCOL[1],.F.,,"L",)
    oprint:printdata(LIN,aCOL[9],MIL(TOT1,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
    IF W_Imp1.R_Tipo.value=1
        oprint:printdata(LIN,aCOL[10],MIL(TOT2,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[11],MIL(TOT3,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[12],MIL(TOT4,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[13],MIL(TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
        oprint:printdata(LIN,aCOL[15],MIL(TOT1+TOT2+TOT3+TOT4+TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
    ENDIF

    SELEC FIN
    FIN->( DBCLOSEAREA() )

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    W_Imp1.release

    Return Nil