#include "minigui.ch"

Function BR_EmpSel1()
    IF IsWindowDefined(WinEmpSel1)=.T.
        DECLARE WINDOW WinEmpSel1
        WinEmpSel1.restore
        WinEmpSel1.Show
        WinEmpSel1.SetFocus
        RETURN
    ENDIF

    AbrirDBF("empresa",,,,RUTAPROGRAMA)

    DEFINE WINDOW WinEmpSel1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 300 ;
        TITLE "Seleccionar empresa" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("TURQUESA") ;
        ON RELEASE CloseTables()

        @ 10,10 BROWSE BR_emp1 ;
            HEIGHT 200 ;
            WIDTH 770 ;
            HEADERS {'Codigo','Empresa','Ejercicio','cierre','Ruta'} ;
            WIDTHS { 50,300,50,50,200 } ;
            WORKAREA empresa ;
            FIELDS {'empresa->NEMP','empresa->EMP','empresa->ejercicio','empresa->cierrejer','empresa->ruta'} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT} ;
            VALUE RECNOEMPRESA ;
            ON DBLCLICK Selec_empresa("EMPRESA",WinEmpSel1.BR_emp1.Value) ;
            ON HEADCLICK { {|| AbrirDBF("empresa",,,,RUTAPROGRAMA,1), WinEmpSel1.BR_emp1.Refresh} , ;
            {|| AbrirDBF("empresa",,,,RUTAPROGRAMA,2), WinEmpSel1.BR_emp1.Refresh} }


        @225,510 BUTTONEX Bt_Selec ;
            CAPTION 'Seleccionar' ;
            WIDTH 90 HEIGHT 25 ;
            TOOLTIP 'Seleccionar empresa' ;
            ACTION Selec_empresa("EMPRESA",WinEmpSel1.BR_emp1.Value)

        @225,610 BUTTONEX Bt_Salir ;
            CAPTION 'Salir' ;
            ICON icobus('salir') ;
            WIDTH 90 HEIGHT 25 ;
            TOOLTIP 'Salir' ;
            ACTION WinEmpSel1.Release ;

    END WINDOW
    VentanaCentrar("WinEmpSel1","Ventana1","Alinear")

    CENTER WINDOW WinEmpSel1
    ACTIVATE WINDOW WinEmpSel1

Return nil

FUNCTION Selec_empresa(LLAMADA,CODEMP2)
    AbrirDBF("empresa",,,,RUTAPROGRAMA)
    IF CODEMP2=0
        GO BOTT
    ELSE
        GO CODEMP2
    ENDIF
***ATENCION CAMBIAR TAMBIEN EN APERTURA Y CAMBIO DE EJECICIO***
    EMPRESA:=RTRIM(EMP)+" "+LTRIM(STR(EJERCICIO))
    NOMEMP:=RTRIM(EMP)
    NOMEMPRESC:=RTRIM(EMP)
    NOMEMPRESA:=RTRIM(EMP)+" ("+STR(EJERCICIO,4)+")"
    EJERANY:=EJERCICIO
    EJERCERRADO:=IF(CAMPO("CIERREJER"),CIERREJER,0)  //PARA QUE NO DE ERROR AL ENTRAR Y PODER ACTUALIZAR
    EMP_IMPRESORA:=IF(CAMPO("IMPRESORA"),IMPRESORA,"")
    NUMEMP:=NEMP
    RECNOEMPRESA:=RECNO()
    DIREMP:=RTRIM(RUTA)
    RUTAEMP    :=RUTAPROGRAMA+"\"+RTRIM(RUTA)
    RUTAEMPRESA:=RUTAPROGRAMA+"\"+RTRIM(RUTA)
    PUBLIC DMA:=CTOD("01-01-"+LTRIM(STR(EJERANY)))
    PUBLIC DMA1:=CTOD("01-01-"+LTRIM(STR(EJERANY)))
    PUBLIC DMA2:=CTOD("31-12-"+LTRIM(STR(EJERANY)))
    IF AT(".TEY",UPPER(RUTAPROGRAMA))<>0 .AND. EJERANY>=2006
        DMA:=CTOD("01-11-"+LTRIM(STR(EJERANY-1)))
        DMA1:=CTOD("01-11-"+LTRIM(STR(EJERANY-1)))
        DMA2:=CTOD("31-10-"+LTRIM(STR(EJERANY)))
    ENDIF
    IF CAMPO("NEMPBUS")
        PUBLIC NUMEMPBUS:=NEMPBUS
    ENDIF
    CLOSE DATABASES
    SetCurrentFolder(RUTAEMPRESA)
    set default to &RUTAEMPRESA
    CLOSE DATABASES
***FIN ATENCION CAMBIAR TAMBIEN EN APERTURA Y CAMBIO DE EJECICIO***


    IF LLAMADA="EMPRESA"
        SetProperty("Ventana1","TITLE",Version_Programa("SUICONTA")+" - "+NOMEMPRESA)
        /*
   ***CERRAR TODAS LAS VENTANAS***
        nTOTVEN:=LEN(_HMG_aFormHandles)             //total de form's abiertos
        aTOTVEN:={}
        FOR NN=2 TO nTOTVEN                         //1-LA PRINCIPAL
            AADD(aTOTVEN, _HMG_aFormNames[NN] )
        NEXT
        IF LEN(aTOTVEN)>0
            FOR NN=1 TO LEN(aTOTVEN)
                NOMVEN2:=aTOTVEN[NN]
                IF _IsWIndowDefined(NOMVEN2)=.T.
                    RELEASE WINDOW &NOMVEN2
                ENDIF
            NEXT
        ENDIF
   ***FIN CERRAR TODAS LAS VENTANAS***
        DO EVENTS
*/
        MSGINFO("Se ha seleccionado la empresa"+HB_OsNewLine()+NOMEMPRESA,"Atención")

        Release window WinEmpSel1                   //***
    ENDIF

Return Nil
