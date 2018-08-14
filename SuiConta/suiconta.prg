#include "minigui.ch"
#include "hbextcdp.ch"

#ifdef __XHARBOUR__
#  xtranslate GetExeFilename() => ExeName()
#else
#  xtranslate GetExeFilename() => hb_ProgName()
#endif

PROCEDURE main()

   ***CODIGO DE PAGINA español***
    REQUEST HB_LANG_ES
    HB_LANGSELECT( "ES" )
    HB_CDPSELECT( "ES850" )

   ***Inicializacion RDD DBFCDX Nativo***
    REQUEST DBFCDX , DBFFPT
    RDDSETDEFAULT( "DBFCDX" )
    SET AUTOPEN OFF                                 //no abrir los indices automaticamente

    SUICONTA()

RETURN

PROCEDURE SUICONTA()
    Set Navigation Extended                         //TAB y ENTER
    SET DATE FORMAT "dd-mm-yyyy"
    SET EPOCH TO YEAR(DATE())-50
    SET DELETE ON
    SET MULTIPLE OFF WARNING

    STORE " " TO RUTASUIZO,RUTAEMP,RUTAEMPRESA,DIREMP,EMP_IMPRESORA
    STORE SPACE(20) TO EMPRESA
    STORE 0 TO EJERANY,EJERCERRADO,NOMEMP,NOMEMPRESA,NOMEMPRESC,NOMPROGRAMA,NUMEMP,RECNOEMPRESA,MENUPLINB,LINPAN2,PAG,ESPERA,CONTADOR,NUMPBUS

    NOMFICINI:=Fichero_Programa("RUTAFICHEROEXE")
    NOMFICINI:=LEFT(NOMFICINI,LEN(NOMFICINI)-4)+".ini"

    RUTAINI:=GetCurrentFolder()
    IF .NOT. FILE(RUTAINI+"\SuiConta.exe")
        RUTAINI:=GetExeFilename()
        IF RAT("\",RUTAINI)<>0
            RUTAINI:=LEFT(RUTAINI,RAT("\",RUTAINI)-1)
        ENDIF
    ENDIF
    IF FILE(NOMFICINI)
        BEGIN INI FILE NOMFICINI
            GET RUTAINI SECTION "Rutas" ENTRY "RutaPrograma" DEFAULT RUTAINI  //LEER
        END INI
    ENDIF

    RUTASUIZO:=RUTAINI
    RUTA2SUIZO:=RUTAINI
    IF FILE("\SUIZO.TEY\CONTA.EXE")
        RUTA2CONTA:="\SUIZO.TEY"
    ELSE
        RUTA2CONTA:=RUTAINI
    ENDIF
    RUTAPROGRAMA:=RUTAINI
    SET DEFAULT TO &RUTAPROGRAMA
    RUTACARPERAEXE:=Ruta_Programa()
    NOMPROGRAMA:=Nombre_Programa()

    SetAppHotKey( VK_F10, 0, { || _OOHG_CallDump() } )


/*
    SET KEY -4 TO F2_PGC1                           //F5
*/

*IF AT("\SUIZOWIN",UPPER(RUTAPROGRAMA))=0
*   MSGSTOP("El programa se debe instalar en la carpeta:"+HB_OsNewLine()+"\Suizowin.red\SuiConta","Error")
*   RETURN
*ENDIF

**** CREAR FICHERO DE EMPRESAS
    IF .NOT. FILE(RUTAPROGRAMA+"\EMPRESA.DBF")
        IF MsgYesNo("No existe el fichero de empresas"+HB_OsNewLine()+ ;
            "Desea crear una nueva empresa en:"+HB_OsNewLine()+RUTAPROGRAMA)=.F.
            RETURN
        ENDIF
        CREAR(RUTAPROGRAMA+"\EMPRESA.DBF")
        USE &RUTAPROGRAMA\EMPRESA
        APPEND BLANK
        REPLACE NEMP WITH 1
        REPLACE EMP WITH "EMPRESA EN PRUEBAS"
        REPLACE EJERCICIO WITH YEAR(DATE())
        REPLACE RUTA WITH "Suizo001"
        INDEX ON NEMP TAG ORDEN1 TO EMPRESA
        CLOSE EMPRESA
        CreateFolder(RUTAPROGRAMA+"\Suizo001")
        Selec_empresa("MENU",1)
        CLOSE DATABASES
        SET DEFAULT TO &RUTAPROGRAMA
    ENDIF
    IF .NOT. FILE(RUTAPROGRAMA+"\EMPRESA.CDX")
        ERASE &RUTAPROGRAMA\EMPRESA.CDX
        USE &RUTAPROGRAMA\EMPRESA
        INDEX ON NEMP TAG ORDEN1 TO &RUTAPROGRAMA\EMPRESA
        CLOSE EMPRESA
    ENDIF
**** FIN CREAR FICHERO DE EMPRESAS

/*
    SET CONFIRM ON
    SET STATUS OFF
    SET SCOREBOARD OFF
    SET CONSOLE OFF
    SET HEADING OFF
    SET WRAP ON
    PUBLIC TERMINA:=0
    CLS
*/
    Selec_empresa("MENU",0)

    //*** SET DEFAULT ICON TO icobus('SuiConta')

    DEFINE WINDOW Ventana1 ;
        OBJ Ventana1 AT 0,0 ;
        WIDTH 1024 ;
        HEIGHT 768-30 ;
        TITLE Version_Programa("SUICONTA")+" - "+NOMEMPRESA ;
        ICON 'SUICONTA' ;
        MAIN BACKCOLOR MiColor("AZULCLARO") ;
        ON RELEASE CloseTables()

        DEFINE MAIN MENU

            POPUP 'Empresas'
                MENUITEM 'Alta/Modif.empresas' ACTION BR_Empresa1()
                MENUITEM 'Seleccionar empresa' ACTION BR_EmpSel1()
                MENUITEM 'Aperturar Nuevo Ejercicio' ACTION EmpApe()
                MENUITEM 'Traspasar Saldos Iniciales' ACTION EmpSini()
                MENUITEM 'Establecer ruta de red' ACTION RutaProgRed()
                SEPARATOR
                MENUITEM 'Salir'       ACTION SalirPrograma()
            END POPUP
            POPUP 'Cuentas'
                MENUITEM 'Altas/Modificacion de Cuentas' ACTION CueDat()
                MENUITEM 'Listado de Cuentas' ACTION Cuelcodigo()
                MENUITEM 'Imprimir Etiquetas Adhesivas' ACTION Cueleti()
**         MENUITEM 'Exportar Datos Cuentas'        ACTION (INIPROGDOS(),EXPCUE(),FINPROGDOS())
                MENUITEM 'Actualizar Datos Cuentas' ACTION Cuentas_Actualizar("SUICONTA")
                MENUITEM 'Cambiar Codigo de cuenta' ACTION CueCamb()
            END POPUP
            POPUP 'Asientos'
                MENUITEM 'Altas/Modif. de Asientos' ACTION AsiDat()
                MENUITEM 'Asiento de cierre automatico' ACTION AsiCie()
                MENUITEM 'Buscar Asientos Descuadrados' ACTION AsiDesc()
                MENUITEM 'Listado de Asientos por concepto/importe' ACTION Asilconcep1()
                MENUITEM 'Listado de Asientos' ACTION Asilfecha("LISTADO")
                MENUITEM 'Libro diario' ACTION Asilfecha("LIBRO")
                MENUITEM 'Revisar Asiento Apertura/Cierre' ACTION AsilApeCie1()
                MENUITEM 'Contabilizar nominas' ACTION Nomina_Conta()
            END POPUP
            POPUP 'Balances'
                MENUITEM 'Mayor de subcuentas' ACTION BR_SuizoExt(RUTAPROGRAMA,,,,STR(NUMEMP))
                MENUITEM 'Libro mayor de Subcuenta' ACTION balmay2()
                MENUITEM 'Balance de Sumas y Saldos' ACTION balsys2()
                MENUITEM 'Balance de Sumas/Saldos Cuenta' ACTION balsys1()
                MENUITEM 'Balance de Situacion' ACTION balsit()
                MENUITEM 'Balance de Perdidas y Ganancias' ACTION balpyg()
                MENUITEM 'Resumen de Perdidas y Ganancias' ACTION ResVenta("PYG")
            END POPUP
            POPUP 'Emitidas'
                MENUITEM 'Altas/Modif. Facturas Emitidas' ACTION FacEmi()
                POPUP 'Cobros de Facturas'
                    MENUITEM 'Repasar Cobros' ACTION CobRep("COBROS")
                    MENUITEM 'Agrupar Saldos Pendientes' ACTION CobAgr()
                    MENUITEM 'Listado de Cobros' ACTION Coblfec()
/*
                    MENUITEM 'Listado Cobros con Riesgo' ACTION (INIPROGDOS(),LCOBRIE(),FINPROGDOS())
                    MENUITEM 'Resumen de Cobros Pendientes' ACTION (INIPROGDOS(),LCOBRES(),FINPROGDOS())
*/
                END POPUP
                POPUP 'Resumen de Facturacion'
/*
                    MENUITEM 'De Prevision de Facturas' ACTION (INIPROGDOS(),RES_PRE(),FINPROGDOS())
                    MENUITEM 'De Cobros Facturas Emitidas' ACTION (INIPROGDOS(),RES_COB(),FINPROGDOS())
                    MENUITEM 'De Pagos Facturas Recibidas' ACTION (INIPROGDOS(),RES_PAG(),FINPROGDOS())
                    MENUITEM 'Modificar Prevision Fact.' ACTION (INIPROGDOS(),PREFAC(),FINPROGDOS())
                    MENUITEM 'De Varios Ejercicios' ACTION (INIPROGDOS(),RES_EJE(),FINPROGDOS())
*/
                END POPUP
**         MENUITEM 'Contabilizar Facturas Emitidas'  ACTION (INIPROGDOS(),FACCONT(),FINPROGDOS())
                POPUP 'Listados Facturas Emitidas'
                    MENUITEM 'Listado Facturas Fecha' ACTION Faclfec()
                    MENUITEM 'Listado Facturas I.V.A.' ACTION Facliva()
                    MENUITEM 'Listado Facturas Pendientes' ACTION FaclPen()
                    MENUITEM 'Listado Vencimiento Facturas' ACTION Faclvto()
                    MENUITEM 'Listado Facturas con Retencion' ACTION Reclret()
                END POPUP
            END POPUP
            POPUP 'Recibidas'
                MENUITEM 'Altas/Modif.Facturas Recibidas' ACTION RecDat()
                POPUP 'Pagos de Facturas Recibidas'
                    MENUITEM 'Repasar Pagos' ACTION CobRep("PAGOS")
                    MENUITEM 'Agrupar Saldos Pendientes' ACTION Pagagr()
                    MENUITEM 'Listado de Pagos' ACTION Paglfec()
/*
                    MENUITEM 'Listado Pagos por Fecha' ACTION (INIPROGDOS(),LPAGFEC(),FINPROGDOS())
                    MENUITEM 'Listado Pagos por Proveedor' ACTION (INIPROGDOS(),LPAGPRO(),FINPROGDOS())
                    MENUITEM 'Listado Pagos con Riesgo' ACTION (INIPROGDOS(),LPAGRIE(),FINPROGDOS())
                    MENUITEM 'Resumen de Pagos Pendientes' ACTION (INIPROGDOS(),LPAGRES(),FINPROGDOS())
*/
                END POPUP
                MENUITEM 'Resumen de IVA' ACTION Ivalres()
**         MENUITEM 'Altas/Modif.Bancos' ACTION (INIPROGDOS(),RECBAN(),FINPROGDOS())
                POPUP 'Listados Facturas Recibidas'
                    MENUITEM 'Listado Facturas Recibidas' ACTION Reclfec()
                    MENUITEM 'Listado Facturas I.V.A.' ACTION Recliva()
                    MENUITEM 'Listado Facturas Pendientes' ACTION Reclpen()
                    MENUITEM 'Listado Vencimiento Facturas' ACTION Reclvto()
                    MENUITEM 'Listado Declaracion Mod.347' ACTION Recldec()
                    MENUITEM 'Listado Facturas con Retencion' ACTION Reclret()
                END POPUP
**         MENUITEM 'Altas facturas de alquiler' ACTION Rec_Alquiler1()
            END POPUP
            POPUP 'Remesas'
                MENUITEM 'Altas/Modificacion de Remesas' ACTION RemDat()
**         MENUITEM 'Importar datos fichero remesa' ACTION RemImpBase()
                MENUITEM 'Eliminar remesa' ACTION RemElim()
                MENUITEM 'Contabilizar remesa' ACTION RemCont()
                MENUITEM 'Listado remesa por cuenta' ACTION Remlcta()
                MENUITEM 'Listado remesa por concepto/importe' ACTION Remlcon()
                MENUITEM 'Imprimir remesa' ACTION RemIrem()
                MENUITEM 'Grabar fichero normas' ACTION RemInor()
            END POPUP
            POPUP 'Utilidades'
                MENUITEM 'Regenerar ficheros' ACTION W_Regfic("SUICONTA")
                MENUITEM 'Propiedades de ficheros' ACTION L_Browse()
                MENUITEM 'Abrir fichero TXT, RTF' ACTION VerFicRtf()
                POPUP 'Operaciones con Empresas'
**            MENUITEM 'Importar Datos de ContaPlus' ACTION (INIPROGDOS(),EMPCPI(),FINPROGDOS())
**            MENUITEM 'Exportar Datos a  ContaPlus' ACTION (INIPROGDOS(),EMPCPE(),FINPROGDOS())
                    MENUITEM 'Importar Asientos' ACTION Asimpor()
                END POPUP
                MENUITEM 'Crear acceso directo' ACTION CrearAccesoDir()
                MENUITEM 'Acerca de '+NOMPROGRAMA ACTION VerDatosProg()
            END POPUP

        END MENU
/*
        @005,010 BUTTONEX Bt_AccesoA1 CAPTION 'Cuentas' ACTION CueDat() WIDTH 120 HEIGHT 40 ICON icobus('cuenta')

        @005,140 BUTTONEX Bt_AccesoB1 CAPTION 'Asientos' ACTION AsiDat() WIDTH 120 HEIGHT 40 ICON icobus('asiento')
        @055,140 BUTTONEX Bt_AccesoB2 CAPTION 'Mayor' ACTION BR_SuizoExt(RUTAPROGRAMA,,,,STR(NUMEMP)) WIDTH 120 HEIGHT 40 ICON icobus('mayor')

        @005,270 BUTTONEX Bt_AccesoC2 CAPTION 'Recibidas' ACTION RecDat() WIDTH 120 HEIGHT 40 ICON icobus('proveedor')
        @055,270 BUTTONEX Bt_AccesoC1 CAPTION 'Emitidas' ACTION FacEmi() WIDTH 120 HEIGHT 40 ICON icobus('cliente')

        @005,400 BUTTONEX Bt_AccesoD2 CAPTION 'Remesas' ACTION RemDat() WIDTH 120 HEIGHT 40 ICON icobus('remesa')

        @400,10 LABEL L_Estado VALUE 'Esta ventana esta inactiva' WIDTH 1000 HEIGHT 65 FONT 'Arial' SIZE 48 ;
            BACKCOLOR MICOLOR('AMARILLO') FONTCOLOR MICOLOR('ROJO') CENTERALIGN INVISIBLE
*/
        DEFINE TOOLBAR ToolBar_1 BUTTONSIZE 40,40 FONT "Arial" SIZE 9 FLAT
        BUTTON Bt_AccesoA1   TOOLTIP "Cuentas"   PICTURE "zcuenta"    ACTION CueDat()
        BUTTON Bt_AccesoB1   TOOLTIP "Asientos"  PICTURE "zasiento"   ACTION AsiDat()
        BUTTON Bt_AccesoB2   TOOLTIP "Mayor"     PICTURE "zmayor"     ACTION BR_SuizoExt(RUTAPROGRAMA,,,,STR(NUMEMP))
        BUTTON Bt_AccesoC2   TOOLTIP "Recibidas" PICTURE "zproveedor" ACTION RecDat()
        BUTTON Bt_AccesoC1   TOOLTIP "Emitidas"  PICTURE "zcliente"   ACTION FacEmi()
        BUTTON Bt_AccesoD2   TOOLTIP "Remesas"   PICTURE "zremesa"    ACTION RemDat()
        BUTTON Bt_Salir      TOOLTIP "Salir"     PICTURE "zsalir2"    ACTION SalirPrograma()
        END TOOLBAR

        DEFINE STATUSBAR FONT "Arial" SIZE 9
        STATUSITEM "Empresa: "+NOMEMPRESA  WIDTH 300
        STATUSITEM "Ruta: "+RUTAEMPRESA    WIDTH 300
        STATUSITEM "Usuario: "     WIDTH 300
        CLOCK
        DATE
        END STATUSBAR

        SET INTERACTIVECLOSE QUERY MAIN
    END WINDOW

    CENTER WINDOW Ventana1
    ACTIVATE WINDOW Ventana1

*** ELIMINAR DIRECTORIO FUSION
    DIR1:=DIRECTORY("*.*","D")
    IF ASCAN(DIR1,{|AVAL| AVAL[1]="SUIZOF"})<>0
        DIRREMOVE("SUIZOF")
    ENDIF
*** FIN ELIMINAR DIRECTORIO FUSION

Return

Function SalirPrograma()
    IF MsgYesNo("¿Desea salir del programa?")=.T.
        RELEASE WINDOW ALL
    ENDIF
Return NIL

Function CloseTables()
    DBCOMMITALL()
    DBUNLOCKALL()
    CLOSE DATABASES
Return NIL


/*
FUNCTION INIPROGDOS()
    Ventana1.Minimize
    DO EVENTS
    Ventana1.L_Estado.Visible:=.T.
    DO EVENTS
    SET CONFIRM ON
    SETMODE(25,80)
    SETCOLOR(MICOLOR)
    CLOSE DATABASES
    Hb_GtInfo( 27 , 'D:\suizowin\Ico\SuiConta.ico' )  //27-HB_GTI_ICONFILE
Return Nil

FUNCTION FINPROGDOS()
    CLOSE DATABASES
    CLEAR
    @ 10,20 SAY Version_Programa("SUICONTA") COLOR COLORTI2
    @ 11,20 SAY 'Esta ventana esta inactiva' COLOR COLORTI
    Ventana1.Restore
    Ventana1.L_Estado.Visible:=.F.
    DO EVENTS
Return Nil
*/


#pragma BEGINDUMP

#include <windows.h>
#include "hbapi.h"
#include "hbapiitm.h"

HB_FUNC( ISCLIPBOARDFORMATAVAILABLE )
    {
    hb_retl( IsClipboardFormatAvailable( hb_parni(1) ) );
    }

HB_FUNC( OPENCLIPBOARD )
    {
    hb_retl( OpenClipboard( (HWND) hb_parnl(1) ) ) ;
    }

HB_FUNC( CLOSECLIPBOARD )
    {
    hb_retl( CloseClipboard() );
    }

HB_FUNC( EMPTYCLIPBOARD )
    {
    hb_retl( EmptyClipboard() ) ;
    }

HB_FUNC( COPYTOCLIPBOARD )                          // CopyToClipboard(cText) store cText in Windows clipboard
    {
    HGLOBAL hglbCopy;
        char *  lptstrCopy;

    char * cStr = HB_ISCHAR( 1 ) ? ( char * ) hb_parc( 1 ) : "";
        int    nLen = strlen( cStr );

    if( ( nLen == 0 ) || ! OpenClipboard( GetActiveWindow() ) )
    return;

    EmptyClipboard();

    hglbCopy = GlobalAlloc( GMEM_DDESHARE, ( nLen + 1 ) * sizeof( TCHAR ) );
        if( hglbCopy == NULL )
    {
    CloseClipboard();
    return;
    }

    lptstrCopy = ( char * ) GlobalLock( hglbCopy );
    memcpy( lptstrCopy, cStr, nLen * sizeof( TCHAR ) );
    lptstrCopy[ nLen ] = ( TCHAR ) 0;               // null character
    GlobalUnlock( hglbCopy );

    SetClipboardData( HB_ISNUM( 2 ) ? ( UINT ) hb_parni( 2 ) : CF_TEXT, hglbCopy );
    CloseClipboard();
    }

HB_FUNC( RETRIEVETEXTFROMCLIPBOARD )
    {
    HGLOBAL hClipMem;
        LPSTR   lpClip;

    if( IsClipboardFormatAvailable( CF_TEXT ) && OpenClipboard( GetActiveWindow() ) )
    {
    hClipMem = GetClipboardData( CF_TEXT );
        if( hClipMem )
    {
    lpClip = ( LPSTR ) GlobalLock( hClipMem );
        if( lpClip )
    {
    hb_retc( lpClip );
        GlobalUnlock( hClipMem );
    }
    else
    hb_retc( "" );
    }
    else
    hb_retc( NULL );
        CloseClipboard();
    }
    else
    hb_retc( NULL );
    }

HB_FUNC( CLEARCLIPBOARD )
    {
    if( OpenClipboard( ( HWND ) hb_parnl( 1 ) ) )
    {
    EmptyClipboard();
        CloseClipboard();
        hb_retl( TRUE );
    }
    else
    hb_retl( FALSE );
    }

#pragma ENDDUMP