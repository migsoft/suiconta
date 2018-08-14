#include "minigui.ch"
#include "hbinkey.ch"
#include "dbinfo.ch"

Function L_Browse()
    NomAlias2:=""
    aDupReg:={}
    aRutas2:={}
    IF AT("ET",UPPER(RUTAPROGRAMA))<>0
        AADD(aRutas2,{ LEFT(RUTAPROGRAMA,AT("\SUIZO",UPPER(RUTAPROGRAMA))-1)+"\Suizo.tey\conta018" })
    ENDIF
    CLOSE DATABASES

    DEFINE WINDOW MiBrowse1 ;
        AT 0,0     ;
        WIDTH 800  ;
        HEIGHT 320 ;
        TITLE "Propiedades de ficheros " ;
        MODAL      ;
        NOSIZE ;
        ON RELEASE CloseTables()
**      NOSYSMENU

        @010,010 BUTTONEX Bt_Buscardir1 CAPTION 'Direcctorio' ICON icobus('buscar') WIDTH 90 HEIGHT 25 ;
            ACTION ( MiBrowse1.T_Ruta.Value:= GetFolder("Carpeta de ficheros",MiBrowse1.T_Ruta.Value) , L_Browse_ActRutas("AÑADIR") ) ;
            NOTABSTOP

        @010,100 BUTTONEX Bt_Buscardir2 CAPTION 'Red' ICON icobus('buscar') WIDTH 70 HEIGHT 25 ;
            ACTION ( MiBrowse1.T_Ruta.Value:= BrowseForFolder(0) , L_Browse_ActRutas("AÑADIR") ) ;
            NOTABSTOP

        @010,180 TEXTBOX T_Ruta ;
            WIDTH 610 ;
            TOOLTIP 'Ruta a buscar' ;
            NOTABSTOP ;
            ON CHANGE VerFicheros()

        @ 40,010 GRID GR_Rutas ;
            HEIGHT 200 ;
            WIDTH 280 ;
            HEADERS {'Rutas'} ;
            WIDTHS {270} ;
            ITEMS aRutas2 ;
            JUSTIFY {BROWSE_JTFY_LEFT} ;
            ON DBLCLICK L_Browse_ActRutas("GRID")

        Menu_Grid("MiBrowse1","GR_Rutas","MENU")


        aFic2:={}
        @ 40,295 GRID GR_Fichero ;
            HEIGHT 200 ;
            WIDTH 480 ;
            HEADERS {'Fichero','Tamaño','Fecha','Hora','Atributos'} ;
            WIDTHS {170,80,80,75,50} ;
            ITEMS aFic2 ;
            JUSTIFY {BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT} ;
            ON DBLCLICK L_BrowseF("BROWSE") ;
            VALUE 1

        Menu_Grid("MiBrowse1","GR_Fichero","MENU",,{"COPREGPP","COPTABPP"})

        @250, 10 CHECKBOX SiIndice CAPTION 'Abrir el primer indice' WIDTH 130 VALUE .T.

        @250,300 BUTTON Bt_Abrir ;
            CAPTION 'Abrir fichero' ;
            WIDTH 100 HEIGHT 25 ;
            ACTION L_BrowseF("BROWSE") ;
            NOTABSTOP

        @250,410 BUTTON Bt_Edit CAPTION 'Editar fichero' WIDTH 100 HEIGHT 25 ;
            ACTION L_BrowseF("EDIT") ;
            NOTABSTOP

        @250,700 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
            ACTION MiBrowse1.Release ;
            WIDTH 90 HEIGHT 25 ;
            TOOLTIP 'Salir' ;
            NOTABSTOP


        MiBrowse1.T_Ruta.Value:=GetCurrentFolder()
        L_Browse_ActRutas("AÑADIR")

        VerFicheros()

    END WINDOW
    VentanaCentrar("MiBrowse1","Ventana1","Alinear")
    CENTER WINDOW MiBrowse1
    ACTIVATE WINDOW MiBrowse1

Return Nil

STATIC FUNCTION L_Browse_ActRutas(LLAMADA)
    DO CASE
    CASE LLAMADA="AÑADIR"
        RUTA2:=MiBrowse1.T_Ruta.Value
        FOR N=1 TO MiBrowse1.GR_Rutas.ItemCount
            IF RUTA2=MiBrowse1.GR_Rutas.Cell(N,1)
                RUTA2:=""
                EXIT
            ENDIF
        NEXT
        IF LEN(RUTA2)>0
            MiBrowse1.GR_Rutas.AddItem({MiBrowse1.T_Ruta.Value})
            MiBrowse1.GR_Rutas.Value:=MiBrowse1.GR_Rutas.ItemCount
        ENDIF
    CASE LLAMADA="GRID"
        MiBrowse1.T_Ruta.Value:=MiBrowse1.GR_Rutas.Cell(MiBrowse1.GR_Rutas.Value,1)
        VerFicheros()
    ENDCASE
Return Nil

STATIC FUNCTION L_BrowseF(LLAMADA)
    LOCAL Color_Brow1 := { || IF( DELETE()=.T. , RGB(255,200,200), RGB(255,255,255)) }

    Fichero1:=MiBrowse1.GR_Fichero.Item(MiBrowse1.GR_Fichero.Value)
    Fichero2:=MiBrowse1.T_Ruta.Value+"\"+Fichero1[1]
    Fichero:=LEFT(Fichero1[1], AT(".",Fichero1[1])-1 )


    IF LLAMADA="EDIT"
        IF FILE(MiBrowse1.T_Ruta.Value+"\"+Fichero+".CDX") .AND. MiBrowse1.SiIndice.Value=.T.
            AbrirDBF(Fichero,,,,MiBrowse1.T_Ruta.Value)
        ELSE
            AbrirDBF(Fichero,"SIN_INDICE",,,MiBrowse1.T_Ruta.Value)
        ENDIF
        EDIT EXTENDED WORKAREA &Fichero TITLE Fichero2
        RETURN
    ENDIF

    DELETED2:=SET(_SET_DELETED)
    SET(_SET_DELETED,.F.)

    aINDEX2:={}
    nINDEX2:=1
    AADD(aINDEX2,"0-Sin Indice")
    IF FILE(MiBrowse1.T_Ruta.Value+"\"+Fichero+".CDX") .AND. MiBrowse1.SiIndice.Value=.T.
        AbrirDBF(Fichero,,,"EXCLUSIVE",MiBrowse1.T_Ruta.Value)
        NUM2:=1
        DO WHILE .T.
            NOM2:=INDEXKEY(NUM2++)
            IF LEN(LTRIM(NOM2))=0
                EXIT
            ELSE
                AADD(aINDEX2,LTRIM(STR(NUM2-1))+"-"+NOM2)
            ENDIF
        ENDDO
        IF LEN(aINDEX2)>1
            nINDEX2:=2
            DBSETORDER(1)
        ELSE
            nINDEX2:=1
            DBSETORDER(0)
        ENDIF
    ELSE
        AbrirDBF(Fichero,"SIN_INDICE",,"EXCLUSIVE",MiBrowse1.T_Ruta.Value)
    ENDIF

    StrucFic2:=DBSTRUCT()
    aCAB:={}
    aTAM:={}
    aCAM:={}
    aCAM2:={}
    aJUS:={}
    aCOLOR:={}
    FOR N=1 TO LEN(StrucFic2)
        AADD(aCAB, RTRIM(StrucFic2[N,1]) )
        AADD(aTAM, IF(StrucFic2[N,3]<=3,40,StrucFic2[N,3]*10 ))  //MINIMO 40
        AADD(aCAM, RTRIM(StrucFic2[N,1]) )
        AADD(aCAM2, RTRIM(StrucFic2[N,1])+" "+StrucFic2[N,2]+"-"+ ;
            LTRIM(STR(StrucFic2[N,3]))+","+LTRIM(STR(StrucFic2[N,4])) )
        DO CASE
        CASE StrucFic2[N,2]="C"
            AADD(aJUS, BROWSE_JTFY_LEFT )
        CASE StrucFic2[N,2]="N"
            AADD(aJUS, BROWSE_JTFY_RIGHT )
        CASE StrucFic2[N,2]="D"
            AADD(aJUS, BROWSE_JTFY_CENTER )
        OTHERWISE
            AADD(aJUS, BROWSE_JTFY_LEFT )
        ENDCASE
        AADD(aCOLOR, Color_Brow1 )
    NEXT

    DEFINE WINDOW MiBrowse2 ;
        AT 0,0     ;
        WIDTH 800  ;
        HEIGHT 630 ;
        TITLE "Propiedades de "+Fichero1[1] ;
        MODAL      ;
        NOSIZE     ;
        ON RELEASE CloseTables()

        ON KEY CONTROL+DELETE          ACTION Borrar_Reg1()
        ON KEY CONTROL+F4              ACTION Borrar_Reg2()
        ON KEY CONTROL+F9              ACTION Duplicar_Reg1()
        ON KEY F5                      ACTION PortaPapelesTitulo()
        ON KEY F6                      ACTION PortaPapelesMayMin()

        @010,10 TEXTBOX T_Ruta ;
            WIDTH 750 ;
            VALUE Fichero2 ;
            TOOLTIP 'Ruta del fichero' ;
            NOTABSTOP READONLY

        @ 40,10 BUTTONEX Bt_Indices CAPTION 'Indices' ICON icobus('portapapeles') ;
            ACTION MiBrowse2_Indices() WIDTH 80 HEIGHT 25 NOTABSTOP
        @ 40,90 COMBOBOX C_Index ;
            WIDTH 370 ITEMS aINDEX2 VALUE nINDEX2 ;
            ON CHANGE (DBSETORDER(MiBrowse2.C_Index.Value-1) , MiBrowse2.BR_Fichero.Refresh )

        @ 40,480 BUTTONEX Bt_Campos CAPTION 'Campos' ICON icobus('portapapeles') ;
            ACTION MiBrowse2_Campos() WIDTH 80 HEIGHT 25 NOTABSTOP
        @ 40,560 COMBOBOX C_Campos ;
            WIDTH 200 ITEMS aCAM2 VALUE 1

        @ 70,10 BROWSE BR_Fichero ;
            HEIGHT 450 ;
            WIDTH 750 ;
            HEADERS aCAB ;
            WIDTHS  aTAM ;
            WORKAREA &Fichero ;
            FIELDS  aCAM ;
            JUSTIFY aJUS ;
            VALUE 1 ;
            DYNAMICBACKCOLOR aCOLOR ;
            ON CHANGE MiBrowse2.Reg3.Value:=LTRIM(MIL(MiBrowse2.BR_Fichero.Value,12,0)) ;
            EDIT INPLACE APPEND

        @530,010 BUTTONEX Bt_Localizar CAPTION 'Localizar' ICON icobus('buscar') ;
            ACTION MiBrowse2_Localizar() WIDTH 90 HEIGHT 25 NOTABSTOP
        @530,100 TEXTBOX T_Buscar WIDTH 250 VALUE '' ON CHANGE { || Buscar() }
        @530,360 BUTTON Bt_Mayus1 CAPTION 'May' WIDTH 35 HEIGHT 25 ACTION MiBrowse2.T_Buscar.Value:=UPPER(MiBrowse2.T_Buscar.Value)
        @530,395 BUTTON Bt_Mayus2 CAPTION 'Min' WIDTH 35 HEIGHT 25 ACTION MiBrowse2.T_Buscar.Value:=LOWER(MiBrowse2.T_Buscar.Value)
        @530,430 BUTTON Bt_Mayus3 CAPTION 'Tit' WIDTH 35 HEIGHT 25 ACTION MiBrowse2.T_Buscar.Value:=NomPropio(MiBrowse2.T_Buscar.Value)

        @530,500 LABEL Reg1 VALUE "Registros:" AUTOSIZE TRANSPARENT
        @530,575 LABEL Reg2 VALUE LTRIM(MIL(LASTREC(),12,0)) AUTOSIZE FONTCOLOR MICOLOR("ROJO") TRANSPARENT
        @530,675 LABEL Reg3 VALUE "" AUTOSIZE FONTCOLOR MICOLOR("AZUL") TRANSPARENT
        @560,010 LABEL Mas1 VALUE "F5 -> Titulo" AUTOSIZE TRANSPARENT
        @580,010 LABEL Mas2 VALUE "F6 -> May/Min" AUTOSIZE TRANSPARENT
        @560,210 LABEL Mas3 VALUE "Alt+A   -> añadir registro" AUTOSIZE TRANSPARENT
        @580,210 LABEL Mas4 VALUE "Ctrl+F9 -> duplicar registro" AUTOSIZE TRANSPARENT
        @560,510 LABEL Mas5 VALUE "Ctrl+Sup -> marcar para borrar" AUTOSIZE TRANSPARENT
        @580,510 LABEL Mas6 VALUE "Ctrl+F4  -> borrar registros marcados" AUTOSIZE TRANSPARENT

        @560,610 LABEL L_Accion VALUE 'Ejecutar accion' AUTOSIZE TRANSPARENT INVISIBLE
        @560,660 TEXTBOX T_Accion WIDTH 120 VALUE '' INVISIBLE
        @580,610 BUTTON Bt_Accion CAPTION 'Ejecutar accion' WIDTH 100 HEIGHT 25 ;
            ACTION Accion_Reg1() ;
            NOTABSTOP INVISIBLE


        DEFINE CONTEXT MENU CONTROL BR_Fichero OF MiBrowse2
**         ITEM "Añadir registro"           ACTION (KEYBOARD HB_K_ALT_A)
        ITEM "Duplicar registro"       ACTION Duplicar_Reg1()
        ITEM "Marcar registro para borrar" ACTION Borrar_Reg1()
        ITEM "Desmarcar registro para borrar" ACTION Borrar_Reg1()
        ITEM "Borrar registros marcados" ACTION Borrar_Reg2()
        SEPARATOR
        ITEM "Marcar TODOS los registros para borrar" ACTION Borrar_Reg9()
        SEPARATOR
        ITEM "Reemplazar datos de los campos" ACTION MiBrowse3_Reemplazar()
        SEPARATOR
        ITEM "Copiar al portapapeles"  ACTION Copiar_RegPP()
        SEPARATOR
        ITEM "Vaciar datos para duplicar" ACTION Duplicar_RegistroPP("VACIAR")
        ITEM "Copiar para duplicar"    ACTION Duplicar_RegistroPP("COPIAR")
        ITEM "Pegar para duplicar"     ACTION Duplicar_RegistroPP("PEGAR")
    END MENU

    MiBrowse2.BR_Fichero.SetFocus

    END WINDOW
    VentanaCentrar("MiBrowse2","Ventana1","Alinear")
    CENTER WINDOW MiBrowse2
    ACTIVATE WINDOW MiBrowse2

    SET(_SET_DELETED,DELETED2)

Return Nil

STATIC FUNCTION MiBrowse2_Localizar()
    NOMBUSCAR:=RTRIM(MiBrowse2.T_Buscar.Value)
    IF LEN(NOMBUSCAR)=0
        RETURN
    ENDIF
    SELEC &Fichero
    IF VALTYPE(FIELDGET(MiBrowse2.C_Campos.Value))<>"C" .AND. ;
        VALTYPE(FIELDGET(MiBrowse2.C_Campos.Value))<>"N"
        MsgStop("Formato incorrecto del campo de busqueda"+HB_OsNewLine()+MiBrowse2.C_Campos.DisplayValue)
        RETURN
    ENDIF
    GO MiBrowse2.BR_Fichero.Value
    DO WHILE .T.
        DO EVENTS
        SKIP
        IF EOF()
            GO TOP
        ENDIF
        IF VALTYPE(FIELDGET(MiBrowse2.C_Campos.Value))="C"
            NOMBUSCAR2:=RTRIM(FIELDGET(MiBrowse2.C_Campos.Value))
        ELSE
            NOMBUSCAR2:=RTRIM(STR(FIELDGET(MiBrowse2.C_Campos.Value)))
        ENDIF
        IF AT(UPPER(NOMBUSCAR),UPPER(NOMBUSCAR2))<>0
            MiBrowse2.BR_Fichero.Value:=recno()
            EXIT
        ENDIF
        IF RECNO()=MiBrowse2.BR_Fichero.Value
            MSGSTOP("No se ha localizado ningun registro")
            EXIT
        ENDIF
    ENDDO

Return Nil


STATIC FUNCTION MiBrowse2_Indices()
    Texto1:=""
    FOR N=2 TO MiBrowse2.C_Index.ItemCount
        Texto1:=Texto1+MiBrowse2.C_Index.Item(N)+HB_OsNewLine()
    NEXT
    CopyToClipboard(Texto1)
    MsgInfo(Texto1,"Portapapeles")
Return Nil

STATIC FUNCTION MiBrowse2_Campos()
    Texto1:=""
    FOR N=1 TO MiBrowse2.C_Campos.ItemCount
        Texto1:=Texto1+MiBrowse2.C_Campos.Item(N)+HB_OsNewLine()
    NEXT
    CopyToClipboard(Texto1)
    MsgInfo(Texto1,"Portapapeles")
Return Nil


STATIC FUNCTION PortaPapelesTitulo()
    Texto1:=RetrieveTextFromClipboard()
    IF VALTYPE(Texto1)="C"
        Texto2:=NomPropio(Texto1)
        CopyToClipboard(Texto2)
        MsgBox(Texto1+HB_OsNewLine()+"---------"+HB_OsNewLine()+Texto2)
    ENDIF
RETURN Nil


STATIC FUNCTION PortaPapelesMayMin()
    Texto1:=RetrieveTextFromClipboard()
    IF VALTYPE(Texto1)="C"
        IF ISUPPER(Texto1)
            Texto2:=LOWER(Texto1)
        ELSE
            Texto2:=UPPER(Texto1)
        ENDIF
        CopyToClipboard(Texto2)
        MsgBox(Texto1+HB_OsNewLine()+"---------"+HB_OsNewLine()+Texto2)
    ENDIF
RETURN Nil


STATIC FUNCTION Borrar_Reg1()
    GO MiBrowse2.BR_Fichero.Value
    IF RLOCK()
        IF DELETE()=.F.
            DELETE
        ELSE
            RECALL
        ENDIF
        DBCOMMIT()
        DBUNLOCK()
    ENDIF
    MiBrowse2.BR_Fichero.Refresh
Return Nil

STATIC FUNCTION Borrar_Reg9()
    IF MsgYesNo("¿Desea marcar todos los registros para borrar?")=.F.
        RETURN
    ENDIF
    PonerEspera("Marcando registros para borrar...")
    IF FLOCK()
        GO TOP
        DO WHILE .NOT. EOF()
            IF DELETE()=.F.
                DELETE
            ENDIF
            SKIP
        ENDDO
        DBCOMMIT()
        DBUNLOCK()
    ENDIF
    MiBrowse2.BR_Fichero.Refresh
    QuitarEspera()
Return Nil


STATIC FUNCTION Borrar_Reg2()
    IF FLOCK()
        PACK
        DBCOMMIT()
        DBUNLOCK()
    ENDIF
    MiBrowse2.BR_Fichero.Refresh
    MiBrowse2.Reg2.Value:=LTRIM(MIL(LASTREC(),12,0))
Return Nil


STATIC FUNCTION VerFicheros()
    aFic1:=DIRECTORY(MiBrowse1.T_Ruta.Value+"\*.DBF")
    aFic2:={}
    FOR N=1 TO LEN(aFic1)
        AADD(aFic2,{aFic1[N,1],LTRIM(MIL(aFic1[N,2],12,0)),DIA(aFic1[N,3],10),aFic1[N,4],aFic1[N,5]})
    NEXT
    ASORT(aFic2,,, { |x, y| UPPER(x[1]) < UPPER(y[1]) })

    MiBrowse1.GR_Fichero.DeleteAllItems
    FOR N=1 TO LEN(aFic2)
        MiBrowse1.GR_Fichero.AddItem(aFic2[N])
    NEXT
Return Nil


STATIC FUNCTION Duplicar_Reg1()
    GO MiBrowse2.BR_Fichero.Value
    Duplicar_RegistroDBF()
    MiBrowse2.BR_Fichero.Value:=RECNO()
    MiBrowse2.BR_Fichero.Refresh
Return Nil


FUNCTION Duplicar_RegistroDBF(nDUPBASE1,nDUPBASE2)
    nDUPBASE1:=IF(nDUPBASE1=NIL,SELECT(),nDUPBASE1)
    nDUPBASE2:=IF(nDUPBASE2=NIL,SELECT(),nDUPBASE2)
    nDUPBASE1:=IF(VALTYPE(nDUPBASE1)="C",SELECT("nDUPBASE1"),nDUPBASE1)
    nDUPBASE2:=IF(VALTYPE(nDUPBASE2)="C",SELECT("nDUPBASE2"),nDUPBASE2)
    SELECT(nDUPBASE1)
    aField:={}
    FOR nField := 1 TO FCOUNT()
        AADD(aField,FIELDGET(nField))
    NEXT
    SELECT(nDUPBASE2)
    APPEND BLANK
    IF RLOCK()
        FOR nField := 1 TO FCOUNT()
            FIELDPUT(nField,aField[nField])
        NEXT
        DBCOMMIT()
        DBUNLOCK()
    ENDIF
Return Nil


STATIC FUNCTION Copiar_RegPP()
    GO MiBrowse2.BR_Fichero.Value
    aField:={}
    FOR nField := 1 TO FCOUNT()
        AADD(aField,FIELDGET(nField))
    NEXT
    TEXTO2:=""
    FOR nField := 1 TO LEN(aField)
        DO CASE
        CASE VALTYPE(aField[nField])="C"
            TEXTO2:=TEXTO2+aField[nField]
        CASE VALTYPE(aField[nField])="N"
            TEXTO2:=TEXTO2+STR(aField[nField])
        CASE VALTYPE(aField[nField])="D"
            TEXTO2:=TEXTO2+DTOC(aField[nField])
        CASE VALTYPE(aField[nField])="L"
            TEXTO2:=TEXTO2+IF(aField[nField]=.T.,"Si","No")
        ENDCASE
        IF nField<LEN(aField)
            TEXTO2:=TEXTO2+HB_K_TAB
        ENDIF
    NEXT
    CopyToClipboard(TEXTO2)
RETURN Nil




STATIC FUNCTION Duplicar_RegistroPP(LLAMADA)
    DO CASE
    CASE LLAMADA="VACIAR"
        aDupReg:={}
        MSGINFO("Se han vaciado los registros para duplicar")
    CASE LLAMADA="COPIAR"
        NomAlias2:=ALIAS()
        GO MiBrowse2.BR_Fichero.Value
        aDupReg2:={}
        FOR nField := 1 TO FCOUNT()
            AADD(aDupReg2,{FIELDNAME(nField),FIELDGET(nField)})
        NEXT
        AADD(aDupReg,aDupReg2)
        MSGINFO("Tiene "+LTRIM(STR(LEN(aDupReg)))+" registros para duplicar")
    CASE LLAMADA="PEGAR"
        IF NomAlias2<>ALIAS()
            MSGSTOP("El nombre de las bases de datos son distintos"+HB_OsNewLine()+ ;
                "Fichero origen: "+NomAlias2+HB_OsNewLine()+ ;
                "Fichero destino: "+ALIAS()+HB_OsNewLine() )
        ELSE
            FOR N=1 TO LEN(aDupReg)
                aDupReg2:=aDupReg[N]
                APPEND BLANK
                IF RLOCK()
                    FOR nField:=1 TO LEN(aDupReg2)
                        nField2:=FIELDNUM(aDupReg2[nField,1])
                        IF nField>0
                            FIELDPUT(nField2,aDupReg2[nField,2])
                        ENDIF
                    NEXT
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
            NEXT
            aDupReg:={}
            MiBrowse2.BR_Fichero.Value:=RECNO()
            MiBrowse2.BR_Fichero.Refresh
        ENDIF
    ENDCASE
Return Nil


STATIC FUNCTION Accion_Reg1()
    IF MSGYESNO(MiBrowse2.T_Accion.value+HB_OsNewLine()+"¿Desea ejecutar la accion?")=.F.
        RETURN
    ENDIF
Return Nil
*REPLACE DOCUMENTO WITH LTRIM(STR(FACTURA)) ALL

*cBloque := "{ ||" + MiBrowse2.T_Accion.value + "}"
*bBloque := &(cBloque)
*Eval( @cBloque)
*cBloque := MiBrowse2.T_Accion.value
*&cBloque


STATIC FUNCTION Buscar()
    IF MiBrowse2.C_Index.Value<=1
        MSGSTOP("No hay indice activado")
    ELSE
        SET SOFTSEEK ON
        DO CASE
        CASE DBORDERINFO(DBOI_KEYTYPE)="C"
            SEEK MiBrowse2.T_Buscar.Value
        CASE DBORDERINFO(DBOI_KEYTYPE)="N"
            SEEK VAL(MiBrowse2.T_Buscar.Value)
        OTHERWISE
            MSGSTOP("tipo de indice no valido")
        ENDCASE
        SET SOFTSEEK OFF
        MiBrowse2.BR_Fichero.Value:=RECNO()
   **MiBrowse2.BR_Fichero.Refresh
    ENDIF

Return Nil


STATIC FUNCTION MiBrowse3_Reemplazar()

    DEFINE WINDOW MiBrowse3 ;
        AT 0,0     ;
        WIDTH 600  ;
        HEIGHT 400 ;
        TITLE "Reemplazar campos en: "+Fichero1[1] ;
        MODAL      ;
        NOSIZE

        @ 15,010 LABEL L_Campos VALUE 'Campo' AUTOSIZE TRANSPARENT
        @ 10,100 COMBOBOX C_Campos WIDTH 200 ITEMS aCAB VALUE 1

        @ 45,010 LABEL L_Texto VALUE 'Texto' AUTOSIZE TRANSPARENT
        @ 40,100 TEXTBOX T_Texto WIDTH 450

        @ 70,010 CHECKBOX SiTodos CAPTION 'Todos los registros' WIDTH 130 VALUE .T.

        @120,010 LABEL L_Progress_1 VALUE 'Reemplazando' AUTOSIZE TRANSPARENT
        @120,100 PROGRESSBAR Progress_1 WIDTH 450 HEIGHT 20 SMOOTH


        @340,410 BUTTONEX B_Reemplaza CAPTION 'Reemplazar' ICON icobus('accion') WIDTH 90 HEIGHT 25 ;
            ACTION MiBrowse3_Reemplazar2()

        @340,510 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('Salir') WIDTH 80 HEIGHT 25 ;
            ACTION MiBrowse3.release


    END WINDOW
    VentanaCentrar("MiBrowse3","Ventana1","Alinear")
    CENTER WINDOW MiBrowse3
    ACTIVATE WINDOW MiBrowse3

Return Nil

STATIC FUNCTION MiBrowse3_Reemplazar2()
    IF USED()=.F.
        MsgStop("La base de datos esta cerrada","error")
        RETURN
    ENDIF

    IF MsgYesNo("Fichero: "+ALIAS()+HB_OsNewLine()+HB_OsNewLine()+ ;
        "Desea reemplazar el campo"+HB_OsNewLine()+ ;
        MiBrowse3.C_Campos.Item(MiBrowse3.C_Campos.Value)+HB_OsNewLine()+HB_OsNewLine()+ ;
        "con el texto: "+HB_OsNewLine()+ ;
        MiBrowse3.T_Texto.Value)=.F.
        RETURN
    ENDIF

    DO CASE
    CASE TYPE(FIELDNAME(MiBrowse3.C_Campos.Value))="C"
        TEXTO2:=MiBrowse3.T_Texto.Value
    CASE TYPE(FIELDNAME(MiBrowse3.C_Campos.Value))="N"
        TEXTO2:=VAL(MiBrowse3.T_Texto.Value)
    CASE TYPE(FIELDNAME(MiBrowse3.C_Campos.Value))="D"
        TEXTO2:=CTOD(MiBrowse3.T_Texto.Value)
    OTHERWISE
        MsgStop("El formato del campo no permite el reemplazo"+HB_OsNewLine()+ ;
            "Formato: "+TYPE(FIELDNAME(MiBrowse3.C_Campos.Value)),"error")
        RETURN
    ENDCASE

    MiBrowse3.Progress_1.RangeMin:=0
    MiBrowse3.Progress_1.RangeMax:=LASTREC()
    TANTO:=0
    GO TOP
    DO WHILE .NOT. EOF()
        MiBrowse3.Progress_1.Value:=++TANTO
        FIELDPUT(MiBrowse3.C_Campos.Value,TEXTO2)
        SKIP
    ENDDO

    MiBrowse2.BR_Fichero.Refresh

    MsgBox("Reemplazo realizado")

Return Nil
