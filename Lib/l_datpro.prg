#include "minigui.ch"

Function VerDatosProg(nomprog,libimage2)
    LOCAL RUTALOGOPROG:=IF(RUTALOGOPROG=NIL,'',RUTALOGOPROG)
    nomprog:=UPPER( IF(nomprog=NIL,programa(),nomprog) )
    libimage2:=IF(libimage2=NIL,"Libreria no soportada",libimage2)
    aWindowsVersion:=WindowsVersion()               //OsVersionInfo()
    DATOSFIC1:=Fichero_Programa("RUTAFICHEROEXE")
    IF FILE(DATOSFIC1)
        DATOSFIC2:=DIRECTORY(DATOSFIC1)
        DATOSFIC1:=DATOSFIC1+" - "+DIA(DATOSFIC2[1,3],11)+" "+DATOSFIC2[1,4]
    ENDIF

    PRIVATE aDatosProg:={}
    AADD(aDatosProg,{"Datos del programa:",Version_Programa(nomprog) })
    AADD(aDatosProg,{"",DATOSFIC1 })
    AADD(aDatosProg,{""," " })
    AADD(aDatosProg,{"Programador:","Jose Miguel Ingles" })
    AADD(aDatosProg,{"","josemisu@yahoo.com.ar" })
    AADD(aDatosProg,{""," " })
    AADD(aDatosProg,{"Compilador:",hb_compiler() })
    AADD(aDatosProg,{"Lenguaje Programa:",Version() })
    AADD(aDatosProg,{"",MiniGuiVersion() })
    AADD(aDatosProg,{"Libreria grafica:",libimage2 })
    AADD(aDatosProg,{""," " })
    AADD(aDatosProg,{"Sistema Operativo:",aWindowsVersion[1] })
    AADD(aDatosProg,{"",aWindowsVersion[2] })
    AADD(aDatosProg,{"",aWindowsVersion[3] })
    AADD(aDatosProg,{"Plataforma:",GetE("OS") })

    DEFINE WINDOW W_Datos_Programa ;
        AT 10,10 ;
        WIDTH 600 HEIGHT 330 ;
        TITLE 'Datos del programa' ;
        MODAL   ;
        NOSIZE BACKCOLOR {255,225,150}              //Naranja claro
**      ICON icobus('datpro') ;

        LIN:=15
        FOR N1=1 TO LEN(aDatosProg)
            FOR N2=1 TO LEN(aDatosProg[N1])
                DO CASE
                CASE N1>=1 .AND. N1<=3 .AND. N2>=2
                    COLOR2:={255,0,0}               //rojo
                CASE N1>=4 .AND. N1<=6 .AND. N2>=2
                    COLOR2:={0,0,255}               //azul
                CASE N1>=7 .AND. N1<=11 .AND. N2>=2
                    COLOR2:={0,153,51}              //Verde Hierba
                CASE N1>=12 .AND. N2>=2
                    COLOR2:={204,51,0}              //Rojo Ladrillo
                OTHERWISE
                    COLOR2:={0,0,0}                 //NEGRO
                ENDCASE
                DO CASE
                CASE N2=1
                    COL2:=10
                CASE N2=2
                    COL2:=150
                CASE N2=3
                    COL2:=300
                OTHERWISE
                    LIN:=LIN+15
                    COL2:=10
                ENDCASE
                NOMLABEL:="L_DatPro"+LTRIM(STR(N1))+LTRIM(STR(N2))
                @LIN,COL2 LABEL &NOMLABEL VALUE aDatosProg[N1,N2] AUTOSIZE TRANSPARENT FONTCOLOR COLOR2
            NEXT
            LIN:=LIN+15
        NEXT

        @265,10 BUTTON Bt_Rutas ;
            CAPTION 'Rutas' ;
            WIDTH 90 HEIGHT 25 ;
            TOOLTIP 'Rutas' ;
            ACTION VerRutasProg()

        @265,110 BUTTONEX Bt_Copiar CAPTION 'Copiar' ICON icobus('portapapeles') WIDTH 90 HEIGHT 25 ;
            ACTION CopiarDatosProg() ;
            TOOLTIP 'Copiar datos al portapapeles' ;
            NOTABSTOP

        @265,210 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Datos_Programa.Release ;
            NOTABSTOP


        @185,150 IMAGE Image_1 ;
            PICTURE '' ;
            WIDTH  400 HEIGHT 75
*             STRETCH     //AJUSTAR

        IF FILE(RUTALOGOPROG)=.T.
            W_Datos_Programa.Image_1.Picture:=RUTALOGOPROG
        ENDIF

    END WINDOW
    VentanaCentrar("W_Datos_Programa","Ventana1")
    CENTER WINDOW W_Datos_Programa
    ACTIVATE WINDOW W_Datos_Programa

RETURN Nil


STATIC FUNCTION VerRutasProg()
    cVerRutasProg:="GetCurrentFolder: "+GetCurrentFolder()+HB_OsNewLine()+ ;
        "Curdir: "+Curdir()+HB_OsNewLine()+ ;
        "SET(_SET_DEFAULT): "+SET(_SET_DEFAULT)+HB_OsNewLine()+ ;
        "gettempdir: "+gettempdir()+HB_OsNewLine()+HB_OsNewLine()+ ;
        "Ruta del programa: "+RUTAPROGRAMA+HB_OsNewLine()+ ;
        "Ruta de la empresa: "+IF(RUTAEMPRESA=NIL," ",RUTAEMPRESA)
    CopyToClipboard(cVerRutasProg)
    MsgInfo(cVerRutasProg ,"Rutas del programa")
Return Nil

STATIC FUNCTION CopiarDatosProg()
    TEXTO2:=""
    FOR N1=1 TO LEN(aDatosProg)
        FOR N2=1 TO LEN(aDatosProg[N1])
            TEXTO2:=TEXTO2+aDatosProg[N1,N2]
            IF LEN(aDatosProg[N1,N2])>0
                TEXTO2:=TEXTO2+HB_OsNewLine()
            ENDIF
        NEXT
    NEXT
    CopyToClipboard(TEXTO2)
    MSGBOX(TEXTO2,"Datos del portapapeles")
Return Nil

FUNCTION PROGRAMA(FPRO2)
    LOCAL FNUM2:=0,FNOM2:=""
    FPRO2:=IF(FPRO2=NIL," ",FPRO2)

    DO WHILE !(PROCNAME(FNUM2)=="")
        FNOM2:=PROCNAME(FNUM2)
        IF FNOM2==(FPRO2) .AND. LEN(FPRO2)>1
            RETURN .T.
        ENDIF
        FNUM2++
        IF UPPER(PROCNAME(FNUM2))=="MAIN" .AND. FPRO2==" "
            EXIT
        ENDIF
    ENDDO
    IF LEN(FPRO2)>1
        RETURN .F.
    ENDIF
Return(FNOM2)
*
FUNCTION Fichero_Programa(nomprog)
    nomprog:=UPPER( IF(nomprog=NIL,programa(),nomprog) )
    DO CASE
    CASE AT("ESCOLA",nomprog)<>0
        FNOM2:="Escola.exe"
    CASE AT("AGENDA",nomprog)<>0
        FNOM2:="Agenda.exe"
    CASE AT("COPIAPRG",nomprog)<>0
        FNOM2:="CopiaPRG.exe"
    CASE AT("COPIAZIP",nomprog)<>0
        FNOM2:="CopiaZIP.exe"
    CASE AT("BUSCAR",nomprog)<>0
        FNOM2:="Buscar.exe"
    CASE AT("FOTOS",nomprog)<>0
        FNOM2:="Fotos.exe"
    CASE AT("SUIALMA",nomprog)<>0
        FNOM2:="SuiAlma.exe"
    CASE AT("SUICONTA",nomprog)<>0
        FNOM2:="SuiConta.exe"
    CASE AT("CAJA",nomprog)<>0
        FNOM2:="Caja.exe"
    CASE AT("SUIWIN2",nomprog)<>0
        FNOM2:="SuiWin2.exe"
    CASE AT("VARIOS",nomprog)<>0
        FNOM2:="Varios.exe"
    CASE AT("RUTAFICHEROEXE",nomprog)<>0
        FNOM2:=GetExeFilename()
    OTHERWISE                                       //"FICHEROEXE"
        FNOM2:=GetExeFilename()
        IF RAT("\",FNOM2)<>0
            FNOM2:=SUBSTR(FNOM2,RAT("\",FNOM2)+1,LEN(FNOM2))
        ENDIF
    ENDCASE
RETURN(FNOM2)
*HB_ArgV(0) = GetExeFilename()
*

FUNCTION Ruta_Programa()
    FNOM2:=GetExeFilename()
    IF RAT("\",FNOM2)<>0
        FNOM2:=LEFT(FNOM2,RAT("\",FNOM2)-1)
    ENDIF
RETURN(FNOM2)


FUNCTION Nombre_Programa(nomprog)
    nomprog:=UPPER( IF(nomprog=NIL,programa(),nomprog) )
    DO CASE
    CASE AT("ESCOLA",nomprog)<>0
        FNOM2:="SUIZO Escola"
    CASE AT("AGENDA",nomprog)<>0
        FNOM2:="SUIZO Agenda"
    CASE AT("COPIAPRG",nomprog)<>0
        FNOM2:="SUIZO CopiaPRG"
    CASE AT("COPIAZIP",nomprog)<>0
        FNOM2:="SUIZO CopiaZIP"
    CASE AT("BUSCAR",nomprog)<>0
        FNOM2:="SUIZO Buscar"
    CASE AT("FOTOS",nomprog)<>0
        FNOM2:="SUIZO Fotos"
    CASE AT("SUIALMA",nomprog)<>0
        FNOM2:="SUIZO Almacen"
    CASE AT("SUICONTA",nomprog)<>0
        FNOM2:="SUIZO Contabilidad"
    CASE AT("CAJA",nomprog)<>0
        FNOM2:="SUIZO Caja"
    CASE AT("SUIWIN2",nomprog)<>0
        FNOM2:="SUIZO win-dos"
    CASE AT("VARIOS",nomprog)<>0
        FNOM2:="SUIZO varios"
    OTHERWISE
        FNOM2:=nomprog
    ENDCASE
RETURN(FNOM2)
*
FUNCTION Version_Programa(nomprog)
    nomprog:=UPPER( IF(nomprog=NIL,programa(),nomprog) )
    FNOM2:=Nombre_Programa(nomprog)
    ANOM1:=CTOD(DIACAMBIARFTO(__DATE__))
    ANOM2:=0
    DO CASE
    CASE AT("ESCOLA",UPPER(nomprog))<>0
        ANOM2:=2004                                 //cumples
        //02-08-04 Diario
        //07-01-05 indiceCDX
        //13-01-05 Cumple->Recibo
        //14-01-05 Facturas
    CASE AT("AGENDA",UPPER(nomprog))<>0
        ANOM2:=2004                                 //hyperlink
        //05-01-05 indiceCDX
    CASE AT("COPIAPRG",UPPER(nomprog))<>0
        ANOM2:=2004                                 //Red
    CASE AT("COPIAZIP",UPPER(nomprog))<>0
        ANOM2:=2005                                 //20-09-05 Empiezo
    CASE AT("BUSCAR",UPPER(nomprog))<>0
        ANOM2:=2005                                 //Red
    CASE AT("FOTOS",UPPER(nomprog))<>0
        ANOM2:=2004                                 //2impresoras
    CASE AT("SUIALMA",UPPER(nomprog))<>0
        ANOM2:=2008                                 //28-06-08 Empiezo
    CASE AT("SUICONTA",UPPER(nomprog))<>0
        ANOM2:=2007                                 //12-03-07 Empiezo
    CASE AT("CAJA",UPPER(nomprog))<>0
        ANOM2:=2005                                 //07-09-05 Empiezo
        //10-01-06 Listado de mayor
        //08-02-06 Listado de consumo de clientes
    CASE AT("SUIWIN2",UPPER(nomprog))<>0
        ANOM2:=2006                                 //17-02-06 Empiezo
    CASE AT("VARIOS",UPPER(nomprog))<>0
        ANOM2:=2006                                 //21-01-06 Empiezo
    OTHERWISE
        ANOM2:=YEAR(ANOM1)
    ENDCASE
    IF ANOM2>0
        FNOM2:=FNOM2+" v"+LTRIM(STR(YEAR(ANOM1)-ANOM2))+"."+LTRIM(STR(MONTH(ANOM1)))
    ENDIF
RETURN(FNOM2)
