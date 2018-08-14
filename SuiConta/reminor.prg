#include "minigui.ch"

procedure RemInor()
    TituloImp:="Cuaderno normas"

    DO CASE
    CASE UPPER(PROGRAMA())="SUICONTA"
        RUTAREMESA:=RUTAEMPRESA
        RUTAREMEMP:=RUTAPROGRAMA
    CASE FILE(RUTACONTABLE+"\REMESA.DBF")=.T.
        RUTAREMESA:=RUTACONTABLE
        RUTAREMEMP:=LEFT(RUTACONTABLE,RAT("\",RUTACONTABLE)-1)
    CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)+"\SUICONTA.EXE") = .T.
        RUTAREMEMP:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)
        AbrirDBF("empresa",,,,RUTAREMEMP,1)
        GO BOTT
        RUTAREMESA:=RUTAREMEMP+"\"+RTRIM(RUTA)
    CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)+"\REMESA.EXE") = .T.
        RUTAREMESA:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)
        RUTAREMEMP:=RUTAREMESA
    OTHERWISE
        RUTAREMESA:=RUTAEMPRESA
        RUTAREMEMP:=RUTAPROGRAMA
    ENDCASE

    aREMESAS:=Remesa_alis(RUTAREMESA)

    DEFINE WINDOW W_Imp1 ;
        AT 10,10 ;
        WIDTH 900 HEIGHT 510 ;
        TITLE 'Imprimir: '+TituloImp ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        @10,10 GRID GR_Remesas ;
            HEIGHT 250 ;
            WIDTH 460 ;
            HEADERS {'Serie','Numero','Remesa','Fecha','Banco','Importe','Documentos','Asiento'} ;
            WIDTHS {0,0,75,80,75,100,50,50 } ;
            ITEMS aREMESAS ;
            VALUE 1 ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','9,999,999','E'},{'TEXTBOX','NUMERIC'}} ;
            ON HEADCLICK {{|| Remesa_Grid_Ord(1)},{|| Remesa_Grid_Ord(2)},{|| Remesa_Grid_Ord(3)}, ;
            {|| Remesa_Grid_Ord(4)},{|| Remesa_Grid_Ord(5)},{|| Remesa_Grid_Ord(6)}, ;
            {|| Remesa_Grid_Ord(7)},{|| Remesa_Grid_Ord(8)}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
            BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
            ON CHANGE ( RemesaActRec() , RemInorAct1() )

        @10,480 GRID GR_Recibos ;
            HEIGHT 250 ;
            WIDTH 410 ;
            HEADERS {'Factura','Cliente','Importe','Vencimiento'} ;
            WIDTHS {50,150,100,80 } ;
            ITEMS {} ;
            COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','DATE'}} ;
            ON HEADCLICK {{|| Remesa_Grid_RecOrd(1)},{|| Remesa_Grid_RecOrd(2)},{|| Remesa_Grid_RecOrd(3)},{|| Remesa_Grid_RecOrd(4)}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER}

        @275,010 LABEL L_Orden VALUE 'Ordenar por' AUTOSIZE TRANSPARENT
        @270,110 RADIOGROUP R_Orden ;
            OPTIONS {'Codigo','Cliente','Factura','Vencimiento' } ;
            VALUE 1 WIDTH 85 ;
            HORIZONTAL

        @305,010 LABEL L_Norma VALUE 'Formato Fichero' AUTOSIZE TRANSPARENT
        @300,110 RADIOGROUP R_Norma ;
            OPTIONS {"NORMA-19 Recibos domiciliados a la vista" , ;
            "NORMA-32 Recibos domiciliados al vencimiento" , ;
            "NORMA-58 Descuento de recibos domiciliados" , ;
            "NORMA-34-1 Cheques y transferencias"} ;
            VALUE 1 WIDTH 300  ;
            ON CHANGE RemInorAct2()

        @410,10 BUTTONEX Bt_Fichero ;
            CAPTION 'Ruta' ;
            ICON icobus('buscar') ;
            WIDTH 90 HEIGHT 25 ;
            ACTION W_imp1.T_Fichero.value:=Getfile( {{'Texto','*.txt'},{'Todos','*.*'}} ,'Buscar fichero normas',W_Imp1.T_Fichero.value, .f. , .t. ) ;
            NOTABSTOP
        @410,110 TEXTBOX T_Fichero WIDTH 470 VALUE RUTAREMEMP+"\Rem_n19.txt" NOTABSTOP READONLY
/*            @410,590 BUTTONEX Bt_FicheroPP1 CAPTION 'portapapeles' ICON icobus('portapapeles') ;
            ACTION ( CopyToClipboard(W_imp1.T_Fichero.value) , MsgInfo(W_imp1.T_Fichero.value,"Portapapeles") ) ;
            WIDTH 100 HEIGHT 25 NOTABSTOP
*/
        @410,590 BUTTONEX Bt_FicheroPP1 CAPTION 'portapapeles' ICON icobus('portapapeles') ;
            ACTION ( SetClipboardText(W_imp1.T_Fichero.value) , MsgInfo(W_imp1.T_Fichero.value,"Portapapeles") ) ;
            WIDTH 100 HEIGHT 25 NOTABSTOP

        @410,700 BUTTONEX Bt_FicheroPP2 CAPTION 'Titulo portapapeles' ICON icobus('portapapeles') ;
            ACTION ( SetClipboardText(RemInor_Titulo()) , AutoMsgInfo(RemInor_Titulo(),"Portapapeles") ) ;
            WIDTH 150 HEIGHT 25 NOTABSTOP

        @445, 10 BUTTONEX B_Grabar CAPTION 'Grabar' ICON icobus('guardar') WIDTH 90 HEIGHT 25 ;
            ACTION RemInorImp()

        @445,110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
            ACTION W_Imp1.release



        RemInorAct1()
        RemesaActRec()

    END WINDOW
    VentanaCentrar("W_Imp1","Ventana1")
    CENTER WINDOW W_Imp1
    ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION RemInor_Titulo()
    NomRem2:="Remesa "+LTRIM(Hb_CSTR(W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)))+"-"+REM_NOM1(W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1))
RETURN(NomRem2)


STATIC FUNCTION RemInorAct1()
    //*** NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)

    NSER2:= GetColValue("GR_Remesas","W_Imp1",1)

    DO CASE
    CASE NSER2<=2
        W_Imp1.R_Norma.Value:=1
    CASE NSER2>=3 .AND. NSER2<=4
        W_Imp1.R_Norma.Value:=2
    CASE NSER2>=5
        W_Imp1.R_Norma.Value:=4
    OTHERWISE
        W_Imp1.R_Norma.Value:=1
    ENDCASE
    RemInorAct2()
Return Nil

STATIC FUNCTION RemInorAct2()
    DO CASE
    CASE W_Imp1.R_Norma.Value=1
        W_imp1.T_Fichero.value:=LEFT(W_imp1.T_Fichero.value,RAT("\",W_imp1.T_Fichero.value)-1)+"\Rem_n19.txt"
    CASE W_Imp1.R_Norma.Value=2
        W_imp1.T_Fichero.value:=LEFT(W_imp1.T_Fichero.value,RAT("\",W_imp1.T_Fichero.value)-1)+"\Rem_n32.txt"
    CASE W_Imp1.R_Norma.Value=3
        W_imp1.T_Fichero.value:=LEFT(W_imp1.T_Fichero.value,RAT("\",W_imp1.T_Fichero.value)-1)+"\Rem_n58.txt"
    CASE W_Imp1.R_Norma.Value=4
        W_imp1.T_Fichero.value:=LEFT(W_imp1.T_Fichero.value,RAT("\",W_imp1.T_Fichero.value)-1)+"\Rem_n34.txt"
    ENDCASE
Return Nil

FUNCTION RemesaActRec(RUTAREMESA)
    //*** NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
    //*** NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
    NSER2:= GetColValue("GR_Remesas","W_Imp1",1)
    NREM2:= GetColValue("GR_Remesas","W_Imp1",2)

    W_Imp1.GR_Recibos.DeleteAllItems
    AbrirDBF("REMESA",,,,RUTAREMESA)
    SEEK (NSER2*100000)+NREM2
    DO WHILE SERIE=NSER2 .AND. NREM=NREM2 .AND. .NOT. EOF()
        DO EVENTS
        W_Imp1.GR_Recibos.AddItem({RTRIM(NFRA),RTRIM(NOMCTA),IMPORTE,FVTO})
        SKIP
    ENDDO
Return Nil


*------------------------------------------------------------*
Procedure VerItem()
*------------------------------------------------------------*
    MsgInfo( 'Col 1: ' + GetColValue( "Grid_1", "Win_1", 1 )+'  ';
        + 'Col 2: ' + GetColValue( "Grid_1", "Win_1", 2 ) )
Return
*------------------------------------------------------------*
Function GetColValue( xObj, xForm, nCol )
*------------------------------------------------------------*
    Local nPos:= GetProperty(xForm, xObj, 'Value')
    Local aRet:= GetProperty(xForm, xObj, 'Item', nPos)
Return aRet[nCol]
*------------------------------------------------------------*
Function SetColValue( xObj, xForm, nCol, xValue )
*------------------------------------------------------------*
    Local nPos:= GetProperty(xForm, xObj, 'Value')
    Local aRet:= GetProperty(xForm, xObj, 'Item', nPos)
    aRet[nCol] := xValue
    SetProperty(xForm, xObj, 'Item', nPos, aRet)
Return NIL


STATIC FUNCTION RemInorImp()
    NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
    NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)

    IF FILE("FIN.DBF")
        AbrirDBF("Fin","SIN_INDICE")
        FIN->( DBCLOSEAREA() )
        ERASE FIN.DBF
        ERASE FIN.CDX
    ENDIF

    AbrirDBF("REMESA",,,,RUTAREMESA)
    COPY TO FIN FOR SERIE=NSER2 .AND. NREM=NREM2
    AbrirDBF("FIN","SIN_INDICE")

    DO CASE
    CASE W_Imp1.R_Orden.Value=1
        INDEX ON STR(CODCTA,8)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=2
        INDEX ON UPPER(NOMCTA)+NFRA TO FIN
    CASE W_Imp1.R_Orden.Value=3
        INDEX ON NFRA+STR(CODCTA,8) TO FIN
    CASE W_Imp1.R_Orden.Value=4
        INDEX ON DTOS(FVTO)+STR(CODCTA,8)+NFRA TO FIN
    ENDCASE

    IF FILE(W_imp1.T_Fichero.value)=.T.
        IF MSGYESNO("¿Desea sobreescribir el fichero?"+HB_OsNewLine()+W_imp1.T_Fichero.value)=.F.
            RETURN
        ENDIF
    ENDIF

    PonerEspera("Grabando fichero de normas"+HB_OsNewLine()+W_imp1.T_Fichero.value)
    FICIMP:=W_imp1.T_Fichero.value
    SET PRINTER TO &FICIMP
    DO CASE
    CASE W_Imp1.R_Norma.Value=1
        REMIN19()
    CASE W_Imp1.R_Norma.Value=2
        REMIN32()
    CASE W_Imp1.R_Norma.Value=3
        REMIN58()
    CASE W_Imp1.R_Norma.Value=4
        REMIN34()
    ENDCASE
    SET PRINTER TO
    QuitarEspera()
    MsgInfo("Fichero grabado correctamente"+HB_OsNewLine()+W_imp1.T_Fichero.value)

RETURN Nil



FUNCTION REM_NOM1(RNUM2)
    DO CASE
    CASE RNUM2=1
        RNOM2:="REC"
    CASE RNUM2=2
        RNOM2:="LET"
    CASE RNUM2=3
        RNOM2:="TAL"
    CASE RNUM2=4
        RNOM2:="PAG"
    CASE RNUM2=5
        RNOM2:="CHE"
    CASE RNUM2=6
        RNOM2:="TRF"
    OTHERWISE
        RNOM2:="***"
    ENDCASE
RETURN(RNOM2)
*
FUNCTION REM_NOM2(RNUM2)
    DO CASE
    CASE RNUM2=1
        RNOM2:="RECIBOS"
    CASE RNUM2=2
        RNOM2:="LETRAS "
    CASE RNUM2=3
        RNOM2:="TALONES"
    CASE RNUM2=4
        RNOM2:="PAGARES"
    CASE RNUM2=5
        RNOM2:="CHEQUES"
    CASE RNUM2=6
        RNOM2:="TRANSFERENCIAS"
    OTHERWISE
        RNOM2:="*******"
    ENDCASE
RETURN(RNOM2)
*
FUNCTION REM_NOM3()
    RNOM2:={"Recibos","Letras","Talones","Pagares","Cheques","Transferencias"}
RETURN(RNOM2)


FUNCTION Remesa_alis(RUTAREMESA)
    PonerEspera("Creando resumen de remesas")
    aREMESAS:={}
    AbrirDBF("REMESA",,,,RUTAREMESA)
    GO TOP
    NSER2:=SERIE
    NREM2:=NREM
    FREM2:=FREM
    CODBAN2:=CODBAN
    NASI2:=NASI
    TOTIMP:=0
    TOTDOC:=0
    DO WHILE .NOT. EOF()
        TOTIMP:=TOTIMP+IMPORTE
        TOTDOC++
        SKIP
        IF SERIE<>NSER2 .OR. NREM<>NREM2 .OR. EOF()
            AADD(aREMESAS,{NSER2,NREM2, LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2) ,FREM2,CODBAN2,TOTIMP,TOTDOC,NASI2})
            NSER2:=SERIE
            NREM2:=NREM
            FREM2:=FREM
            CODBAN2:=CODBAN
            NASI2:=NASI
            TOTIMP:=0
            TOTDOC:=0
        ENDIF
    ENDDO
    ASORT(aREMESAS,,, { |x, y| DTOS(x[4]) > DTOS(y[4]) })
    QuitarEspera()
RETURN(aREMESAS)


FUNCTION Remesa_Grid_Ord(LLAMADA)
    IF W_Imp1.GR_Remesas.ItemCount=0
        MSGSTOP("No hay registros para ordenar")
        RETURN
    ENDIF
    IF VALTYPE(W_Imp1.GR_Remesas.Cell(1,LLAMADA))<>"C" .AND. ;
        VALTYPE(W_Imp1.GR_Remesas.Cell(1,LLAMADA))<>"N" .AND. ;
        VALTYPE(W_Imp1.GR_Remesas.Cell(1,LLAMADA))<>"D"
        MSGSTOP("Esta columna no se puede ordenar")
        RETURN
    ENDIF
    PonerEspera("Ordenando la tabla por "+W_Imp1.GR_Remesas.Header(LLAMADA) )
    ESTOY:=IF(W_Imp1.GR_Remesas.Value=0,1,W_Imp1.GR_Remesas.Value)
    aOrdenar2:={}
    FOR N=1 TO W_Imp1.GR_Remesas.ItemCount
        AADD(aOrdenar2,W_Imp1.GR_Remesas.Item(N))
    NEXT
    DO CASE
    CASE LLAMADA<=3
        ASORT(aOrdenar2,,, { |x, y| (x[1]*100000)+x[2] < (y[1]*100000)+y[2] })
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="C"
        ASORT(aOrdenar2,,, { |x, y| UPPER(x[LLAMADA]) < UPPER(y[LLAMADA]) })
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="N"
        ASORT(aOrdenar2,,, { |x, y| x[LLAMADA] < y[LLAMADA] })
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="D"
        ASORT(aOrdenar2,,, { |x, y| DTOS(x[LLAMADA]) < DTOS(y[LLAMADA]) })
    ENDCASE
    W_Imp1.GR_Remesas.DeleteAllItems
    FOR N=1 TO LEN(aOrdenar2)
        W_Imp1.GR_Remesas.AddItem(aOrdenar2[N])
    NEXT
    W_Imp1.GR_Remesas.Value:=ESTOY
    QuitarEspera()

Return Nil

FUNCTION Remesa_Grid_RecOrd(LLAMADA)
    IF W_Imp1.GR_Recibos.ItemCount=0
        MSGSTOP("No hay registros para ordenar")
        RETURN
    ENDIF
    IF VALTYPE(W_Imp1.GR_Recibos.Cell(1,LLAMADA))<>"C" .AND. ;
        VALTYPE(W_Imp1.GR_Recibos.Cell(1,LLAMADA))<>"N" .AND. ;
        VALTYPE(W_Imp1.GR_Recibos.Cell(1,LLAMADA))<>"D"
        MSGSTOP("Esta columna no se puede ordenar")
        RETURN
    ENDIF
    PonerEspera("Ordenando la tabla por "+W_Imp1.GR_Recibos.Header(LLAMADA) )
    ESTOY:=IF(W_Imp1.GR_Recibos.Value=0,1,W_Imp1.GR_Recibos.Value)
    aOrdenar2:={}
    FOR N=1 TO W_Imp1.GR_Recibos.ItemCount
        AADD(aOrdenar2,W_Imp1.GR_Recibos.Item(N))
    NEXT
    DO CASE
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="C"
        ASORT(aOrdenar2,,, { |x, y| UPPER(x[LLAMADA]) < UPPER(y[LLAMADA]) })
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="N"
        ASORT(aOrdenar2,,, { |x, y| x[LLAMADA] < y[LLAMADA] })
    CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="D"
        ASORT(aOrdenar2,,, { |x, y| DTOS(x[LLAMADA]) < DTOS(y[LLAMADA]) })
    ENDCASE
    W_Imp1.GR_Recibos.DeleteAllItems
    FOR N=1 TO LEN(aOrdenar2)
        W_Imp1.GR_Recibos.AddItem(aOrdenar2[N])
    NEXT
    W_Imp1.GR_Recibos.Value:=ESTOY
    QuitarEspera()

    Return Nil