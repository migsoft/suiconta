#include "minigui.ch"
#include "winprint.ch"

Function br_suizoext(M1RUTA,M1CODCTA,M1FEC1,M1FEC2,M1NUMEMP,M1LLAMADA)
    M1RUTA:=IF(M1RUTA=NIL,"",M1RUTA)
    M1CODCTA:=IF(M1CODCTA=NIL,0,M1CODCTA)
    M1FEC1:=IF(M1FEC1=NIL,DMA1,M1FEC1)
    M1FEC2:=IF(M1FEC2=NIL,IF(DATE()<DMA2,DATE(),DMA2),M1FEC2)
    M1NUMEMP:=IF(M1NUMEMP=NIL," ",M1NUMEMP)
    M1LLAMADA:=IF(M1LLAMADA=NIL," ",M1LLAMADA)
/*
    IF LEN(M1RUTA)=0
        M1RUTA2:=LEFT(RUTAPROGRAMA,AT("\SUIZOWIN",UPPER(RUTAPROGRAMA))-1)+"\GRUPOSP"
        aM1DIR:=Mi_Dir(M1RUTA2,'empresa.dbf')
        ASORT( aM1DIR,,, {| x, y | x[3] > y[3] } )
        nM1DIR:=0
        FOR M1=1 TO LEN(aM1DIR)
            IF AT("\EMP\EMPRESA.DBF",UPPER(aM1DIR[M1,1]))<>0
                M1RUTA:=M1RUTA2+"\"+LEFT(aM1DIR[M1,1],RAT("\EMP\EMPRESA.DBF",UPPER(aM1DIR[M1,1]))-1)
                EXIT
            ENDIF
        NEXT
    ENDIF

    IF LEN(M1RUTA)=0
        DO CASE
        CASE FILE("\SUIZOWIN.GAT\SUICONTA\SUICONTA.EXE")=.T.
            M1RUTA:="\SUIZOWIN.GAT\SUICONTA"
        CASE FILE("\SUIZOWIN.TEY\SUICONTA\SUICONTA.EXE")=.T.
            M1RUTA:="\SUIZOWIN.TEY\SUICONTA"
        ENDCASE
    ENDIF
*/
    IF M1LLAMADA="SUIWIN2"
        DEFINE WINDOW WinBRext1 ;
            AT 0,0 WIDTH 430 HEIGHT 250 ;
            TITLE "Mayor de subcuentas" ;
            MAIN ;
            NOSIZE BACKCOLOR MiColor("AZULCLARO")
        ELSE
            DEFINE WINDOW WinBRext1 ;
                AT 0,0 WIDTH 430 HEIGHT 250 ;
                TITLE "Mayor de subcuentas" ;
                MODAL ;
                NOSIZE BACKCOLOR MiColor("AZULCLARO")
            ENDIF

            @010,010 BUTTONEX Bt_Buscardir CAPTION 'Ruta' ICON icobus('buscar') WIDTH 90 HEIGHT 25 ;
                ACTION WinBRext1.T_Ruta.Value:=GetFolder(,WinBRext1.T_Ruta.Value) ;
                NOTABSTOP
            @010,110 TEXTBOX T_Ruta WIDTH 300 VALUE '' MAXLENGTH 150 ;
                ON CHANGE br_suizoextemp(M1NUMEMP)

            @045,010 LABEL L_Empresa VALUE 'Empresa' AUTOSIZE TRANSPARENT
            @040,110 COMBOBOX C_Empresa WIDTH 300 ON CHANGE (WinBRext1.C_RutaEmp.Value:=WinBRext1.C_Empresa.Value, ;
                WinBRext1.C_CodEmp.Value:=WinBRext1.C_Empresa.Value)
            @040,410 COMBOBOX C_Rutaemp WIDTH 300 INVISIBLE
            @040,410 COMBOBOX C_CodEmp  WIDTH 300 INVISIBLE

            @070,010 BUTTONEX Bt_BuscarCue CAPTION 'Cuenta' ICON icobus('buscar') WIDTH 90 HEIGHT 25 ;
                ACTION (br_suizocue(WinBRext1.C_RutaEmp.DisplayValue,VAL(WinBRext1.T_CodCta.Value), ;
                "WinBRext1","T_CodCta"), ;
                BusCtaSuizo(WinBRext1.C_RutaEmp.DisplayValue,WinBRext1.T_CodCta.Value,"WinBRext1","T_NomCta") ) ;
                NOTABSTOP
            @070,110 TEXTBOX T_CodCta WIDTH 90 MAXLENGTH 8 ;
                ON LOSTFOCUS (WinBRext1.T_CodCta.Value:=PCODCTA3(WinBRext1.T_CodCta.Value) , ;
                BusCtaSuizo(WinBRext1.C_RutaEmp.DisplayValue,WinBRext1.T_CodCta.Value,"WinBRext1","T_NomCta") )
            @070,210 TEXTBOX T_NomCta WIDTH 200 READONLY NOTABSTOP

            @105,10 LABEL L_Fec1 VALUE 'Desde la fecha' AUTOSIZE TRANSPARENT
            @100,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DATE()
            @105,215 LABEL L_Fec1b VALUE 'Año = ejercicios anteriores' AUTOSIZE TRANSPARENT

            @135,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
            @130,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DATE()
            @135,215 LABEL L_Fec2b VALUE 'Año = ejercicios posteriores' AUTOSIZE TRANSPARENT

            @160,10 LABEL L_Progres VALUE 'Progreso' AUTOSIZE TRANSPARENT
            @160,110 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH


            @190,010 BUTTONEX Bt_Mayor CAPTION 'Ver mayor' ICON icobus('consultar') ;
                ACTION br_suizoextpan() ;
                WIDTH 90 HEIGHT 25

            @190,110 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
                ACTION WinBRext1.Release ;
                WIDTH 90 HEIGHT 25 ;
                NOTABSTOP

        END WINDOW

        WinBRext1.T_Ruta.Value:=M1RUTA
        WinBRext1.T_CodCta.Value:=STR(M1CODCTA)
        IF WinBRext1.D_Fec1.Value=DATE()
            WinBRext1.D_Fec1.Value:=M1FEC1
        ENDIF
        WinBRext1.D_Fec2.Value:=M1FEC2
        BusCtaSuizo(WinBRext1.C_RutaEmp.DisplayValue,WinBRext1.T_CodCta.Value,"WinBRext1","T_NomCta")

        WinBRext1.T_CodCta.SetFocus

        VentanaCentrar("WinBRext1","Ventana1")
        ACTIVATE WINDOW WinBRext1

Return Nil


STATIC FUNCTION br_suizoextemp(M1NUMEMP)
    M1RUTA2:=WinBRext1.T_Ruta.Value
    DO CASE
    CASE AT("SUICONTA",UPPER(WinBRext1.T_Ruta.Value))<>0
        M1RUTA2:=LEFT(WinBRext1.T_Ruta.Value,RAT("SUICONTA",UPPER(WinBRext1.T_Ruta.Value))-2)
        IF AT("SUICONTA",UPPER(M1RUTA2))=0
            M1RUTA2:=WinBRext1.T_Ruta.Value
        ENDIF
    CASE AT("GRUPOSP",UPPER(WinBRext1.T_Ruta.Value))<>0
        IF AT("EMP",UPPER(WinBRext1.T_Ruta.Value))<>0
            M1RUTA2:=LEFT(WinBRext1.T_Ruta.Value,RAT("EMP",UPPER(WinBRext1.T_Ruta.Value))-2)
        ENDIF
    ENDCASE
    aRUTAEMP:=aRUTAEMP(WinBRext1.T_Ruta.Value)
    WinBRext1.C_Empresa.DeleteAllItems
    WinBRext1.C_RutaEmp.DeleteAllItems
    WinBRext1.C_CodEmp.DeleteAllItems
    IF LEN(aRUTAEMP)>0
        IRNUMEMP:=0
        FOR N=1 TO LEN(aRUTAEMP)
            WinBRext1.C_Empresa.AddItem(aRUTAEMP[N,1])
            WinBRext1.C_RutaEmp.AddItem(M1RUTA2+"\"+aRUTAEMP[N,2])
            WinBRext1.C_CodEmp.AddItem(aRUTAEMP[N,4])
            IF STRTRAN(aRUTAEMP[N,4]," ","")==STRTRAN(M1NUMEMP," ","")
                IRNUMEMP:=N
                WinBRext1.D_Fec1.Value:=CTOD("01-01-"+STR(aRUTAEMP[N,3],4))
            ENDIF
        NEXT
        IF IRNUMEMP<>0
            WinBRext1.C_Empresa.Value:=IRNUMEMP
            WinBRext1.C_RutaEmp.Value:=IRNUMEMP
            WinBRext1.C_CodEmp.Value:=IRNUMEMP
        ELSE
            WinBRext1.C_Empresa.Value:=LEN(aRUTAEMP)
            WinBRext1.C_RutaEmp.Value:=LEN(aRUTAEMP)
            WinBRext1.C_CodEmp.Value:=LEN(aRUTAEMP)
        ENDIF
    ENDIF
Return Nil

FUNCTION br_suizoextpan()
    LOCAL aArq:={}
    Aadd( aArq , { 'NASI'      , 'N' , 10  , 0 } )
    Aadd( aArq , { 'FECHA'     , 'D' , 10  , 0 } )
    Aadd( aArq , { 'CODCTA'    , 'N' ,  8  , 0 } )
    Aadd( aArq , { 'CONCEPTO'  , 'C' , 40  , 0 } )
    Aadd( aArq , { 'DEBE'      , 'N' , 14  , 3 } )
    Aadd( aArq , { 'HABER'     , 'N' , 14  , 3 } )
    Aadd( aArq , { 'SALDO'     , 'N' , 14  , 3 } )
    Aadd( aArq , { 'NUMEMP'    , 'C' ,  5  , 0 } )

    aMayor:={}
    AADD(aMayor,{0,CTOD("  -  -  "),0,"Saldo inicial",0,0,0,''})

    FEC1:=WinBRext1.D_Fec1.Value
    FEC2:=WinBRext1.D_Fec2.Value
    CODCTA2:=WinBRext1.T_CodCta.Value

    DO CASE
    CASE AT("SUICONTA",UPPER(WinBRext1.C_RutaEmp.DisplayValue))<>0
        FOR ANY2=YEAR(FEC1) TO YEAR(FEC2)
            WinBRext1.L_Progres.Value:=STR(ANY2,4)
            aRUTA2:=BUSRUTAEMP(WinBRext1.C_RutaEmp.DisplayValue,WinBRext1.C_CodEmp.DisplayValue,ANY2,"SUICONTA")
            RUTA2:=aRUTA2[1]

            IF FILE(RUTA2+"\APUNTES.DBF") .AND. ;
                FILE(RUTA2+"\APUNTES.CDX")
                AbrirDBF("APUNTES",,,,RUTA2)
                GO TOP
                SET DELETE OFF
                TANTO:=0
                DO WHILE .NOT. EOF()
                    WinBRext1.P_Progres.Value:= ++TANTO*100/LASTREC()
                    DO EVENTS
                    IF CODCTA=VAL(CODCTA2) .AND. FECHA<=FEC2 .AND. DELETE()=.F.

               ***SALDO INICIAL***
                        IF APUNTES->NASI=1
                            IF ANY2=YEAR(FEC1)
                                aMayor[1,5]:=aMayor[1,5]+MDA_PASAR(APUNTES->DEBE ,ANY2 )
                                aMayor[1,6]:=aMayor[1,6]+MDA_PASAR(APUNTES->HABER,ANY2 )
                            ENDIF
                            SKIP
                            LOOP
                        ENDIF
               ***FIN SALDO INICIAL***

                        IF APUNTES->FECHA<FEC1
                            aMayor[1,5]:=aMayor[1,5]+MDA_PASAR(APUNTES->DEBE ,ANY2 )
                            aMayor[1,6]:=aMayor[1,6]+MDA_PASAR(APUNTES->HABER,ANY2 )
                        ELSE
                            AADD(aMayor,{APUNTES->NASI,APUNTES->FECHA,APUNTES->CODCTA,RTRIM(APUNTES->NOMAPU), ;
                                MDA_PASAR(APUNTES->DEBE ,ANY2 ) , ;
                                MDA_PASAR(APUNTES->HABER,ANY2 ) , 0 , aRUTA2[2] } )
                        ENDIF
                    ENDIF
                    SKIP
                ENDDO
                SET DELETE ON
                APUNTES->( DBCLOSEAREA() )
            ENDIF
        NEXT
    CASE AT("GRUPOSP",UPPER(WinBRext1.C_RutaEmp.DisplayValue))<>0
        RUTAEMP2:=WinBRext1.C_RutaEmp.DisplayValue
        IF AT("EMP",UPPER(RUTAEMP2))<>0
            RUTAEMP2:=LEFT(RUTAEMP2,RAT("EMP",UPPER(RUTAEMP2))-2)+"\EMP"
        ELSE
            RUTAEMP2:=RUTAEMP2+"\EMP"
        ENDIF
        ASIAPE3:=0
        FOR ANY2=YEAR(FEC1) TO YEAR(FEC2)
            WinBRext1.L_Progres.Value:=STR(ANY2,4)
            aRUTA2:=BUSRUTAEMP(WinBRext1.C_RutaEmp.DisplayValue,WinBRext1.C_CodEmp.DisplayValue,ANY2,"CONTAELI")
            RUTA2:=aRUTA2[1]
            ASIAPE2:=0
            ASICIE2:=0
            IF FILE(RUTAEMP2+"\EMPRESA.DBF") .AND. ;
                FILE(RUTAEMP2+"\EMPRESA.CDX")
                AbrirDBF("EMPRESA",,,,RUTAEMP2)
                SEEK aRUTA2[2]
                ASIAPE2:=APERTURA
                ASICIE2:=CIERRE
                EMPRESA->( DBCLOSEAREA() )
            ENDIF
            IF FILE(RUTA2+"\DIARIO.DBF") .AND. ;
                FILE(RUTA2+"\DIARIO.CDX")
                AbrirDBF("DIARIO",,,,RUTA2)
                IF ASIAPE2<>0 .AND. ASIAPE3=0
                    ASIAPE3:=1
                    DBSETORDER(3)
                    SEEK ASIAPE2
                    DO WHILE ASIEN=ASIAPE2 .AND. .NOT. EOF()
                        IF VAL(SUBCTA)=VAL(CODCTA2)
                            aMayor[1,1]:=DIARIO->ASIEN
                            aMayor[1,5]:=aMayor[1,5]+DIARIO->EURODEBE
                            aMayor[1,6]:=aMayor[1,6]+DIARIO->EUROHABER
                            EXIT
                        ENDIF
                        SKIP
                    ENDDO
                ENDIF
                AbrirDBF("DIARIO",,,,RUTA2)
                GO TOP
                SET DELETE OFF
                TANTO:=0
                DO WHILE .NOT. EOF()
                    WinBRext1.P_Progres.Value:= ++TANTO*100/LASTREC()
                    DO EVENTS
                    IF ASIEN=ASICIE2 .OR. ASIEN=ASIAPE2
                        SKIP
                        LOOP
                    ENDIF
                    IF VAL(SUBCTA)=VAL(CODCTA2) .AND. FECHA<=FEC2 .AND. DELETE()=.F.
                        IF DIARIO->FECHA<FEC1
                            aMayor[1,5]:=aMayor[1,5]+DIARIO->EURODEBE
                            aMayor[1,6]:=aMayor[1,6]+DIARIO->EUROHABER
                        ELSE
                            AADD(aMayor,{DIARIO->ASIEN,DIARIO->FECHA,VAL(DIARIO->SUBCTA),RTRIM(DIARIO->CONCEPTO), ;
                                DIARIO->EURODEBE  , ;
                                DIARIO->EUROHABER , 0 , aRUTA2[2] } )
                        ENDIF
                    ENDIF
                    SELEC DIARIO
                    SKIP
                ENDDO
                SET DELETE ON
                DIARIO->( DBCLOSEAREA() )
            ENDIF
        NEXT
    ENDCASE


    ASORT(aMayor,,, { |x, y| DTOS(x[2])+STR(x[1],10) < DTOS(y[2])+STR(y[1],10) })
    SALDO2:=0
    FOR N=1 TO LEN(aMayor)
        SALDO2:=SALDO2+aMayor[N,5]-aMayor[N,6]
        aMayor[N,7]:=SALDO2
    NEXT

    DECLARE WINDOW WinBRext1b

    DEFINE WINDOW WinBRext1b ;
        AT 0,0     ;
        WIDTH 750  ;
        HEIGHT 600 ;
        TITLE "Mayor de subcuentas" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("AZULCLARO") ;
        ON INIT WinBRext1b.GR_Fin1.SetFocus

        @015,010 LABEL L_NomEmp VALUE 'Empresa' AUTOSIZE TRANSPARENT
        @010,100 TEXTBOX T_NomEmp WIDTH 300 VALUE WinBRext1.C_Empresa.DisplayValue READONLY

        @045,010 LABEL L_RutaEmp VALUE 'Ruta' AUTOSIZE TRANSPARENT
        @040,100 TEXTBOX T_RutaEmp WIDTH 300 VALUE WinBRext1.C_RutaEmp.DisplayValue READONLY

        @075,10 LABEL L_CodCta VALUE 'Subcuenta' AUTOSIZE TRANSPARENT
        @070,100 TEXTBOX T_CodCta WIDTH 100 VALUE WinBRext1.T_CodCta.Value READONLY

        @075,250 LABEL L_NomCta VALUE 'Nombre' AUTOSIZE TRANSPARENT
        @070,310 TEXTBOX T_NomCta WIDTH 300 VALUE WinBRext1.T_NomCta.Value READONLY

        @015,510 LABEL L_Fec1 VALUE 'Desde la fecha' AUTOSIZE TRANSPARENT
        @010,600 TEXTBOX D_Fec1 WIDTH 100 VALUE WinBRext1.D_Fec1.Value READONLY DATE

        @045,510 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
        @040,600 TEXTBOX D_Fec2 WIDTH 100 VALUE WinBRext1.D_Fec2.Value READONLY DATE

        @100,10 GRID GR_Fin1 ;
            HEIGHT 420 ;
            WIDTH 700 ;
            HEADERS {'Asiento','Fecha','Cuenta','Descripcion','Debe','Haber','Saldo','Empresa'} ;
            WIDTHS { 75,90,0,200,100,100,100,0 } ;
            ITEMS aMayor ;
            VALUE IF(LEN(aMayor)>=1,1,0) ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, {'TEXTBOX','CHARACTER'}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT, ;
            BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            ON DBLCLICK br_suizoextpanb()

        Menu_Grid("WinBRext1b","GR_Fin1","MENU",,{"COPCELPP","COPREGPP","COPTABPP"})

        @530,330 BUTTONEX Bt_Duplicar CAPTION 'Duplicar' ICON icobus('guardar') WIDTH 90 HEIGHT 25 ;
            ACTION Menu_WinBRext1b("Duplicarasiento") NOTABSTOP

        @530,430 BUTTONEX Bt_Imprimir CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION br_suizoextimp() NOTABSTOP

        @530,530 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') WIDTH 90 HEIGHT 25 ;
            ACTION br_suizoextCalc() NOTABSTOP

        @530,630 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION WinBRext1b.Release NOTABSTOP
        /*
        DEFINE CONTEXT MENU CONTROL GR_Fin1 OF WinBRext1b
        ITEM "Copiar registro"         ACTION Menu_WinBRext1b("Copiarregistro")
        ITEM "Duplicar asiento"        ACTION Menu_WinBRext1b("Duplicarasiento")
    END MENU
*/

END WINDOW
VentanaCentrar("WinBRext1b","Ventana1")
ACTIVATE WINDOW WinBRext1b

Return Nil


STATIC FUNCTION Menu_WinBRext1b(LLAMADA)
    DO CASE
    CASE LLAMADA="Copiarregistro"
        aMayor2:=WinBRext1b.GR_Fin1.Item(WinBRext1b.GR_Fin1.Value)
        CopyToClipboard( LTRIM(STR(aMayor2[1]))+HB_OsNewLine()+ ;
            DIA(aMayor2[2])+HB_OsNewLine()+ ;
            LTRIM(STR(aMayor2[3]))+HB_OsNewLine()+ ;
            RTRIM(aMayor2[4])+HB_OsNewLine()+ ;
            LTRIM(MIL(aMayor2[5],15,2))+HB_OsNewLine()+ ;
            LTRIM(MIL(aMayor2[6],15,2))+HB_OsNewLine()+ ;
            LTRIM(MIL(aMayor2[7],15,2))+HB_OsNewLine()+ ;
            RTRIM(aMayor2[8]))
        MSGBOX(RetrieveTextFromClipboard(),"Texto del portapapeles")
    CASE LLAMADA="Duplicarasiento"
        Asi_Duplicar(WinBRext1b.T_RutaEmp.Value,WinBRext1b.GR_Fin1.Cell(WinBRext1b.GR_Fin1.Value,1),WinBRext1b.GR_Fin1.Cell(WinBRext1b.GR_Fin1.Value,8))
    ENDCASE
Return Nil

STATIC FUNCTION br_suizoextpanb()
    br_suizoasi(WinBRext1b.T_RutaEmp.Value,WinBRext1b.GR_Fin1.Cell(WinBRext1b.GR_Fin1.Value,1),WinBRext1b.GR_Fin1.Cell(WinBRext1b.GR_Fin1.Value,8))
Return Nil

STATIC FUNCTION br_suizoextimp()

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 400 HEIGHT 160 ;
        TITLE 'Imprimir: Mayor de subcuenta' ;
        MODAL NOSIZE

        aIMP:=Impresoras(0)
        @015,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
        @010,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

        @045,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
        @040,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

        @040, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
            width 150 value .f. ;
            ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

        @070, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
            width 150 value .f.

        @100, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
            ACTION br_suizoextimpI()

        @100,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    ACTIVATE WINDOW W_Imp1

Return Nil

STATIC FUNCTION br_suizoextimpI()
    dirimp:=GetCurrentFolder()

    oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
    oprint:init()
    oprint:setunits("MM",5)
    oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
    if oprint:lprerror
        oprint:release()
        return nil
    endif
    oprint:begindoc(WinBRext1b.Title)
    oprint:setpreviewsize(2)                        // tamaño del preview
    oprint:beginpage()

    PonerEspera("Procesando datos del listado")

    PAG:=0
    LIN:=0
    FOR N=1 TO WinBRext1b.GR_Fin1.ItemCount
        DO EVENTS

        IF LIN>=260 .OR. PAG=0
            IF PAG<>0
                LIN:=LIN+5
                oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
                oprint:endpage()
                oprint:beginpage()
            ENDIF
            PAG=PAG+1

            oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
            oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
            oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

            oprint:printdata(25,105,WinBRext1.C_Empresa.DisplayValue,"times new roman",12,.F.,,"C",)
            oprint:printdata(32,105,WinBRext1b.Title,"times new roman",18,.F.,,"C",)

            oprint:printdata(40,20,"Desde fecha "+DIA(WinBRext1.D_Fec1.Value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,20,"Hasta fecha "+DIA(WinBRext1.D_Fec2.Value,10),"times new roman",12,.F.,,"L",)
            oprint:printdata(40,70,"Subcuenta "+LTRIM(WinBRext1.T_CodCta.Value),"times new roman",12,.F.,,"L",)
            oprint:printdata(45,70,WinBRext1.T_NomCta.Value,"times new roman",12,.F.,,"L",)

            LIN:=55
            oprint:printdata(LIN, 30,"Asiento","times new roman",10,.F.,,"R",)
            oprint:printdata(LIN, 50,"Fecha","times new roman",10,.F.,,"R",)
            oprint:printdata(LIN, 52,"Descripcion","times new roman",10,.F.,,"L",)
            oprint:printdata(LIN,130,"Debe","times new roman",10,.F.,,"R",)
            oprint:printdata(LIN,150,"Haber","times new roman",10,.F.,,"R",)
            oprint:printdata(LIN,170,"Saldo","times new roman",10,.F.,,"R",)
            oprint:printline(LIN+4,20,LIN+4,170,,0.5)

            LIN:=LIN+5
        ENDIF

        IF WinBRext1b.GR_Fin1.Cell(N,1)<>0          //NASI
            oprint:printdata(LIN, 30,WinBRext1b.GR_Fin1.Cell(N,1),"times new roman",10,.F.,,"R",)
            oprint:printdata(LIN, 50,WinBRext1b.GR_Fin1.Cell(N,2),"times new roman",10,.F.,,"R",)
        ENDIF
        oprint:printdata(LIN, 52,WinBRext1b.GR_Fin1.Cell(N,4),"times new roman",10,.F.,,"L",)
        oprint:printdata(LIN,130,MIL(WinBRext1b.GR_Fin1.Cell(N,5),14,2),"times new roman",10,.F.,,"R",)
        oprint:printdata(LIN,150,MIL(WinBRext1b.GR_Fin1.Cell(N,6),14,2),"times new roman",10,.F.,,"R",)
        oprint:printdata(LIN,170,MIL(WinBRext1b.GR_Fin1.Cell(N,7),14,2),"times new roman",10,.F.,,"R",)

        LIN:=LIN+5

    NEXT

    oprint:endpage()
    oprint:enddoc()
    oprint:RELEASE()

    QuitarEspera()

    W_Imp1.release

Return Nil



STATIC FUNCTION br_suizoextCalc()
    nOption:=InputWindow("Elija el tipo de salida",{"Programa"},{1},{{"Hoja Calculo","Hoja Excel"}},300,300,.T.,{"Aceptar","Cancelar"})
    DO CASE
    CASE nOption[1] == Nil
        RETURN
    CASE nOption[1]=1
        br_suizoextCalc2()
    CASE nOption[1]=2
        br_suizoextexcel()
    ENDCASE
Return Nil


STATIC FUNCTION br_suizoextCalc2()
    local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

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

    oSheet:getCellByPosition(0,0):SetString(NOMPROGRAMA)
    oSheet:getCellByPosition(0,1):SetString(DIA(DATE(),10))
    oSheet:getCellByPosition(0,2):SetString(NOMEMPRESA)
    oSheet:getCellByPosition(0,3):SetString(WinBRext1b.Title)
    oSheet:getCellByPosition(0,3):CharHeight:=16
    oSheet:getCellByPosition(0,4):SetString('Desde fecha:')
    oSheet:getCellByPosition(1,4):SetString(DIA(WinBRext1.D_Fec1.value,10))
    oSheet:getCellByPosition(0,5):SetString('Hasta fecha:')
    oSheet:getCellByPosition(1,5):SetString(DIA(WinBRext1.D_Fec2.value,10))
    oSheet:getCellByPosition(0,6):SetString('Empresa:')
    oSheet:getCellByPosition(1,6):SetString(WinBRext1.C_Empresa.DisplayValue)
    oSheet:getCellByPosition(0,7):SetString('Cuenta:')
    oSheet:getCellByPosition(1,7):SetString(WinBRext1.T_CodCta.Value+" "+WinBRext1.T_NomCta.Value)

    LIN:=9

    oSheet:getCellByPosition(0,LIN):SetString("Asiento")
    oSheet:getCellByPosition(1,LIN):SetString("Fecha")
    oSheet:getCellByPosition(2,LIN):SetString("Descripcion")
    oSheet:getCellByPosition(3,LIN):SetString("Debe")
    oSheet:getCellByPosition(4,LIN):SetString("Haber")
    oSheet:getCellByPosition(5,LIN):SetString("Saldo")
    oSheet:getCellRangeByPosition(0,LIN,5,LIN):HoriJustify:=2
    oSheet:getCellRangeByPosition(0,LIN,5,LIN):CharWeight:=150  //NEGRITA
    aMiColor:=MiColor("AMARILLOPALIDO")
    oSheet:getCellRangeByPosition(0,LIN,5,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

    LIN++


    PAG:=0
    LIN2:=LIN
    FOR N=1 TO WinBRext1b.GR_Fin1.ItemCount
        DO EVENTS

        IF WinBRext1b.GR_Fin1.Cell(N,1)<>0
            oSheet:getCellByPosition(0,LIN):SetValue(WinBRext1b.GR_Fin1.Cell(N,1))
            oSheet:getCellByPosition(1,LIN):SetString(DIA(WinBRext1b.GR_Fin1.Cell(N,2),8))
            oSheet:getCellByPosition(1,LIN):HoriJustify:=2
        ENDIF
        oSheet:getCellByPosition(2,LIN):SetString(RTRIM(WinBRext1b.GR_Fin1.Cell(N,4)))
        oSheet:getCellByPosition(3,LIN):SetValue(WinBRext1b.GR_Fin1.Cell(N,5))
        oSheet:getCellByPosition(4,LIN):SetValue(WinBRext1b.GR_Fin1.Cell(N,6))
        IF LIN2=LIN
            TEXTO2:="=+"+CHR(64+4)+LTRIM(STR(LIN+1))+"-"+CHR(64+5)+LTRIM(STR(LIN+1))
        ELSE
            TEXTO2:="=+"+CHR(64+6)+LTRIM(STR(LIN))+"+"+CHR(64+4)+LTRIM(STR(LIN+1))+"-"+CHR(64+5)+LTRIM(STR(LIN+1))
        ENDIF
        oSheet:getCellByPosition(5,LIN):SetFormula(TEXTO2)
        oSheet:getCellRangeByPosition(3,LIN,5,LIN):NumberFormat:=4  //#.##0,00

        LIN++

    NEXT

    oColumns:=oSheet:getColumns()
    oColumns:getByIndex(2):setPropertyValue("OptimalWidth", .T.)

    QuitarEspera()

*   W_Imp1.release

Return Nil




STATIC FUNCTION br_suizoextexcel()
    LOCAL oExcel, oHoja

    PonerEspera("Procesando datos del listado")

    oExcel := TOleAuto():New( "Excel.Application" )
    IF Ole2TxtError() != 'S_OK'
        QuitarEspera()
        MsgStop('Excel no esta disponible','error')
        RETURN Nil
    ENDIF
    oExcel:WorkBooks:Add()
    oExcel:Sheets("Hoja1"):Name := "Listado"
*   oExcel:Sheets("Hoja2"):Name := "Resumen"
    oHoja := oExcel:Get( "ActiveSheet" )
    oHoja:Cells:Font:Name := "Arial"
    oHoja:Cells:Font:Size := 10

    LIN:=9

    oHoja:Cells( LIN, 1 ):Value := "Asiento"
    oHoja:Cells( LIN, 2 ):Value := "Fecha"
    oHoja:Cells( LIN, 3 ):Value := "Descripcion"
    oHoja:Cells( LIN, 4 ):Value := "Debe"
    oHoja:Cells( LIN, 5 ):Value := "Haber"
    oHoja:Cells( LIN, 6 ):Value := "Saldo"

    oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Font:Bold := .T.
    oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Interior:ColorIndex := 36  //sombrear celdas
    oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):Borders(4):LineStyle:= 1  //linea inferior
    oHoja:Range(LetraExcel(1)+LTRIM(STR(LIN))+":"+LetraExcel(6)+LTRIM(STR(LIN))):HorizontalAlignment := -4108  //Centrar

    LIN++


    PAG:=0
    LIN2:=LIN
    FOR N=1 TO WinBRext1b.GR_Fin1.ItemCount
        DO EVENTS

        IF WinBRext1b.GR_Fin1.Cell(N,1)<>0
            oHoja:Cells( LIN, 1 ):Value := WinBRext1b.GR_Fin1.Cell(N,1)
            oHoja:Cells( LIN, 2 ):Value := DIA(WinBRext1b.GR_Fin1.Cell(N,2),8)
        ENDIF
        oHoja:Cells( LIN, 3 ):Value := RTRIM(WinBRext1b.GR_Fin1.Cell(N,4))
        oHoja:Cells( LIN, 4 ):Value := WinBRext1b.GR_Fin1.Cell(N,5)
        oHoja:Cells( LIN, 4 ):Set( "NumberFormat", "#.##0,00" )
        oHoja:Cells( LIN, 5 ):Value := WinBRext1b.GR_Fin1.Cell(N,6)
        oHoja:Cells( LIN, 5 ):Set( "NumberFormat", "#.##0,00" )
        IF LIN2=LIN
            TEXTO2:="=+"+CHR(64+4)+LTRIM(STR(LIN))+"-"+CHR(64+5)+LTRIM(STR(LIN))
        ELSE
            TEXTO2:="=+"+CHR(64+6)+LTRIM(STR(LIN-1))+"+"+CHR(64+4)+LTRIM(STR(LIN))+"-"+CHR(64+5)+LTRIM(STR(LIN))
        ENDIF
        oHoja:Cells( LIN, 6 ):Value := TEXTO2
        oHoja:Cells( LIN, 6 ):Set( "NumberFormat", "#.##0,00" )

        LIN++

    NEXT

    oHoja:Cells( 1, 1 ):Value := NOMPROGRAMA
    oHoja:Cells( 2, 1 ):Value := DIA(DATE(),10)
    oHoja:Cells( 4, 1 ):Value := 'desde:'
    oHoja:Cells( 4, 2 ):Value := DIA(WinBRext1.D_Fec1.value,10)
    oHoja:Cells( 5, 1 ):Value := 'hasta:'
    oHoja:Cells( 5, 2 ):Value := DIA(WinBRext1.D_Fec2.value,10)

    oHoja:Range("A1:B7"):HorizontalAlignment:= -4131  //Izquierda

    FOR nCol:=1 TO FCOUNT()
        oHoja:Columns( nCol ):AutoFit()
    NEXT

    oHoja:Cells( 3, 1 ):Value := WinBRext1b.Title
    oHoja:Cells( 3, 1 ):Font:Size := 16
    oHoja:Cells( 6, 1 ):Value := 'Empresa:'
    oHoja:Cells( 6, 2 ):Value := WinBRext1.C_Empresa.DisplayValue
    oHoja:Cells( 7, 1 ):Value := 'Cuenta:'
    oHoja:Cells( 7, 2 ):Value := WinBRext1.T_CodCta.Value+" "+WinBRext1.T_NomCta.Value

   *Guardar como
   *oHoja:SaveAs( WinBRext1b.Title )

    oHoja:Cells( 1, 1 ):Select()
    oExcel:Visible := .T.

    IF AT("XHARBOUR",UPPER(VERSION()))=1
        oHoja:end()
        oExcel:end()
    ENDIF

    QuitarEspera()

*   W_Imp1.release

    Return Nil