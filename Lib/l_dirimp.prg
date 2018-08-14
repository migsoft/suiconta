#include "minigui.ch"
#include "winprint.ch"

FUNCTION DireccionImp(TIPO,aDireccionImp,aDireccionRte)
PUBLIC aFontD:={},aFontR:={}
   aDireccionRte:=IF(aDireccionRte=NIL,{"","","","","",""},aDireccionRte)


   DEFINE WINDOW W_DirImp1 ;
      AT 10,10 ;
      WIDTH 800 HEIGHT 560 ;
      TITLE 'Listado de direcciones' ;
      MODAL      ;
      NOSIZE

      @10,10 RADIOGROUP R_Tipo ;
      OPTIONS { 'Sobre' , 'Carta' , 'Etiqueta' } ;
      VALUE TIPO WIDTH 75 ON CHANGE Actualiz_SobreImp() HORIZONTAL

      @ 15,590 LABEL L_TotLis VALUE 'Registros: '+LTRIM(MIL(LEN(aDireccionImp),15,0)) AUTOSIZE TRANSPARENT

#ifdef __OOHG__
      @ 40,010 GRID GR_Dir1 ;
      HEIGHT 150 ;
      WIDTH 780 ;
      HEADERS {'Codigo','Nombre','Direccion','Cod.Postal','Poblacion','Provincia','Pais'} ;
      WIDTHS { 75,200,200,75,200,200,100 } ;
      ITEMS aDireccionImp ;
      VALUE 1 ;
      EDIT INPLACE ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
                    {'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}} ;
      ON HEADCLICK {{|| Grid1_Ord(1)},{|| Grid1_Ord(2)},{|| Grid1_Ord(3)},{|| Grid1_Ord(4)},{|| Grid1_Ord(5)},{|| Grid1_Ord(6)},{|| Grid1_Ord(7)}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT}
#else
      @ 40,010 GRID GR_Dir1 ;
      HEIGHT 150 ;
      WIDTH 780 ;
      HEADERS {'Codigo','Nombre','Direccion','Cod.Postal','Poblacion','Provincia','Pais'} ;
      WIDTHS { 75,200,200,75,200,200,100 } ;
      ITEMS aDireccionImp ;
      VALUE 1 ;
      EDIT INPLACE {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
                    {'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}} ;
      ON HEADCLICK {{|| Grid1_Ord(1)},{|| Grid1_Ord(2)},{|| Grid1_Ord(3)},{|| Grid1_Ord(4)},{|| Grid1_Ord(5)},{|| Grid1_Ord(6)},{|| Grid1_Ord(7)}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT}
#endif

      Menu_Grid("W_DirImp1","GR_Dir1","MENU")

      @190,10 LABEL L_ImpCol VALUE 'Imprimir: ' AUTOSIZE TRANSPARENT
      @190,100 CHECKBOX C_ImpColCod CAPTION 'Codigo'     HEIGHT 20 WIDTH 90 VALUE .F. TRANSPARENT NOTABSTOP
      @190,190 CHECKBOX C_ImpColNom CAPTION 'Nombre'     HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP
      @190,280 CHECKBOX C_ImpColDir CAPTION 'Direccion'  HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP
      @190,370 CHECKBOX C_ImpColCPo CAPTION 'Cod.Postal' HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP
      @190,460 CHECKBOX C_ImpColPob CAPTION 'Poblacion'  HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP
      @190,550 CHECKBOX C_ImpColPrv CAPTION 'Provincia'  HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP
      @190,640 CHECKBOX C_ImpColPai CAPTION 'Pais'       HEIGHT 20 WIDTH 90 VALUE .T. TRANSPARENT NOTABSTOP


      @210,10 RADIOGROUP R_Horizontal ;
      OPTIONS { 'Impresion vertial' , 'Impresion horizotal' } ;
      VALUE 1 WIDTH 150 HORIZONTAL

      aFontD:={'Arial',14,.F.,.F.,MICOLOR("NEGRO"),.F.,.F.,0}
      @240,010 BUTTON L_FuenteDir1 CAPTION 'Fuente direccion' WIDTH 120 HEIGHT 20 ;
      ACTION ( aFontD:=GetFont(aFontD[1],aFontD[2],aFontD[3],aFontD[4],aFontD[5],aFontD[6],aFontD[7],aFontD[8]) , ;
      aFontD:=IF(aFontD[1]=='' , {'Arial',14,.F.,.F.,MICOLOR("NEGRO"),.F.,.F.,0} , aFontD ) , ;
      W_DirImp1.L_FuenteDirF.Value:=aFontD[1] , ;
      W_DirImp1.L_FuenteDirN.Value:=LTRIM(STR(aFontD[2])) , ;
      W_DirImp1.L_FuenteDirF.FontColor:=aFontD[5] )
      @242,140 LABEL L_FuenteDirN VALUE LTRIM(STR(aFontD[2])) AUTOSIZE TRANSPARENT
      @242,160 LABEL L_FuenteDirF VALUE aFontD[1] AUTOSIZE TRANSPARENT

      aFontR:={'Arial',10,.F.,.F.,MICOLOR("NEGRO"),.F.,.F.,0}
      @270,010 BUTTON L_FuenteRte1 CAPTION 'Fuente remite' WIDTH 120 HEIGHT 20 ;
      ACTION ( aFontR:=GetFont(aFontR[1],aFontR[2],aFontR[3],aFontR[4],aFontR[5],aFontR[6],aFontR[7],aFontR[8]) , ;
      aFontR:=IF(aFontR[1]=='' , {'Arial',10,.F.,.F.,MICOLOR("NEGRO"),.F.,.F.,0} , aFontR ) , ;
      W_DirImp1.L_FuenteRteF.Value:=aFontR[1] , ;
      W_DirImp1.L_FuenteRteN.Value:=LTRIM(STR(aFontR[2])) , ;
      W_DirImp1.L_FuenteRteF.FontColor:=aFontR[5] )
      @272,140 LABEL L_FuenteRteN VALUE LTRIM(STR(aFontR[2])) AUTOSIZE TRANSPARENT
      @272,160 LABEL L_FuenteRteF VALUE aFontR[1] AUTOSIZE TRANSPARENT

      @300,010 LABEL L_Copias VALUE 'Numero de copias' AUTOSIZE TRANSPARENT
      @300,125 TEXTBOX T_Copias HEIGHT 20 WIDTH 50 VALUE 1 NUMERIC RIGHTALIGN

      @330,010 LABEL DirMarIzq VALUE 'Margen direccion' AUTOSIZE TRANSPARENT
      @330,140 LABEL L_DirMarIzq VALUE 'izquierdo' AUTOSIZE TRANSPARENT
      @330,200 TEXTBOX T_DirMarIzq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN
      @330,260 LABEL L_DirMarSup VALUE 'superior' AUTOSIZE TRANSPARENT
      @330,320 TEXTBOX T_DirMarSup HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN

      @355,010 LABEL RteMarIzq VALUE 'Margen remite' AUTOSIZE TRANSPARENT
      @355,140 LABEL L_RteMarIzq VALUE 'izquierdo' AUTOSIZE TRANSPARENT
      @355,200 TEXTBOX T_RteMarIzq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN
      @355,260 LABEL L_RteMarSup VALUE 'superior' AUTOSIZE TRANSPARENT
      @355,320 TEXTBOX T_RteMarSup HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN

      @385,010 CHECKBOX C_Remite CAPTION 'Imprimir direccion de remite' HEIGHT 20 WIDTH 200 VALUE .F. TRANSPARENT NOTABSTOP
      @410,010 TEXTBOX T_RteNom1 HEIGHT 20 WIDTH 300 VALUE aDireccionRte[1] ON LOSTFOCUS aDireccionRte[1]:=W_DirImp1.T_RteNom1.Value
      @430,010 TEXTBOX T_RteNom2 HEIGHT 20 WIDTH 300 VALUE aDireccionRte[2] ON LOSTFOCUS aDireccionRte[2]:=W_DirImp1.T_RteNom2.Value
      @450,010 TEXTBOX T_RteNom3 HEIGHT 20 WIDTH 050 VALUE aDireccionRte[3] ON LOSTFOCUS aDireccionRte[3]:=W_DirImp1.T_RteNom3.Value
      @450,070 TEXTBOX T_RteNom4 HEIGHT 20 WIDTH 240 VALUE aDireccionRte[4] ON LOSTFOCUS aDireccionRte[4]:=W_DirImp1.T_RteNom4.Value
      @470,010 TEXTBOX T_RteNom5 HEIGHT 20 WIDTH 300 VALUE aDireccionRte[5] ON LOSTFOCUS aDireccionRte[5]:=W_DirImp1.T_RteNom5.Value
      @490,010 TEXTBOX T_RteNom6 HEIGHT 20 WIDTH 300 VALUE aDireccionRte[6] ON LOSTFOCUS aDireccionRte[6]:=W_DirImp1.T_RteNom6.Value



      @215,410 LABEL L_Etiqueta VALUE 'Tipo de etiqueta' AUTOSIZE TRANSPARENT
      @210,520 COMBOBOX C_Etiqueta WIDTH 200 ITEMS Etiqueta_Tipos("DESCRIPCION") VALUE 6 ON CHANGE Actualiz_EtiqImp()

      @240,410 LABEL L_MargEtiq VALUE 'Margenes hoja' AUTOSIZE TRANSPARENT
      @240,560 LABEL L_IzqEtiq VALUE 'Izquierdo' AUTOSIZE TRANSPARENT
      @240,620 TEXTBOX T_IzqEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' ON LOSTFOCUS Actualiz_NumEtiq() RIGHTALIGN
      @240,680 LABEL L_SupEtiq VALUE 'Superior' AUTOSIZE TRANSPARENT
      @240,740 TEXTBOX T_SupEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' ON LOSTFOCUS Actualiz_NumEtiq() RIGHTALIGN

      @265,410 LABEL L_DimEtiq VALUE 'Dimensiones etiqueta' AUTOSIZE TRANSPARENT
      @265,560 LABEL L_AnchoEtiq VALUE 'Ancho' AUTOSIZE TRANSPARENT
      @265,620 TEXTBOX T_AnchoEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN
      @265,680 LABEL L_AltoEtiq VALUE 'Alto' AUTOSIZE TRANSPARENT
      @265,740 TEXTBOX T_AltoEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.9' FORMAT 'E' RIGHTALIGN

      @290,380 LABEL NumEtiq VALUE '' WIDTH 20 HEIGHT 20 TRANSPARENT RIGHTALIGN
      @290,410 LABEL L_NumEtiq VALUE 'Numero de etiquetas' AUTOSIZE TRANSPARENT
      @290,560 LABEL L_horizEtiq VALUE 'horizontal' AUTOSIZE TRANSPARENT
      @290,620 TEXTBOX T_horizEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999' FORMAT 'E' ON LOSTFOCUS Actualiz_NumEtiq() RIGHTALIGN
      @290,680 LABEL L_VertEtiq VALUE 'Vertical' AUTOSIZE TRANSPARENT
      @290,740 TEXTBOX T_VertEtiq HEIGHT 20 WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999' FORMAT 'E' ON LOSTFOCUS Actualiz_NumEtiq() RIGHTALIGN

      @315,410 LABEL L_Empezar VALUE 'Empezar por la etiqueta' AUTOSIZE TRANSPARENT
      @315,620 TEXTBOX T_Empezar HEIGHT 20 WIDTH 50 VALUE 1 NUMERIC RIGHTALIGN

      @340,410 CHECKBOX C_SiLin CAPTION 'Imprimir lineas de division' HEIGHT 20 WIDTH 160 VALUE .f.



      @380,410 PROGRESSBAR Progress_1 WIDTH 380 HEIGHT 10 SMOOTH

LIN2:=400
COL2:=400
draw rectangle in window W_DirImp1 at LIN2,COL2+10 to LIN2+2,COL2+390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @LIN2+15,COL2+10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @LIN2+10,COL2+100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @LIN2+45,COL2+220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @LIN2+40,COL2+280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @LIN2+40,COL2+10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_DirImp1.C_Impresora.Enabled:=IF(W_DirImp1.nImp.Value=.T.,.F.,.T.)

      @LIN2+70,COL2+10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @LIN2+100,COL2+10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION SobreImpi()

      @LIN2+100,COL2+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_DirImp1.release


      Actualiz_SobreImp()
      Actualiz_EtiqImp()

      END WINDOW
      VentanaCentrar("W_DirImp1","Ventana1")
      CENTER WINDOW W_DirImp1
      ACTIVATE WINDOW W_DirImp1

Return Nil


STATIC FUNCTION Menu_SobreImp(LLAMADA)
   IF W_DirImp1.GR_Dir1.Value<=0
      MSGSTOP("No hay ningun registro seleccionado")
      RETURN
   ENDIF
   DO CASE
   CASE LLAMADA="Nuevo"
      IF MSGYESNO("Desea añadir un registro en blanco")=.T.
         W_DirImp1.GR_Dir1.AddItem({"Nuevo","","","","","",""})
         W_DirImp1.GR_Dir1.Value:=W_DirImp1.GR_Dir1.ItemCount
         W_DirImp1.GR_Dir1.Refresh
      ENDIF
   CASE LLAMADA="Duplicar"
      IF MSGYESNO("Desea duplicar el registro activo")=.T.
         W_DirImp1.GR_Dir1.AddItem(W_DirImp1.GR_Dir1.Item(W_DirImp1.GR_Dir1.Value))
         W_DirImp1.GR_Dir1.Value:=W_DirImp1.GR_Dir1.ItemCount
         W_DirImp1.GR_Dir1.Refresh
      ENDIF
   CASE LLAMADA="Eliminar"
      IF MSGYESNO("Desea eliminar el registro activo")=.T.
         W_DirImp1.GR_Dir1.DeleteItem(W_DirImp1.GR_Dir1.Value)
         W_DirImp1.GR_Dir1.Refresh
      ENDIF
   CASE LLAMADA="CopiarRegistro"
      IF MSGYESNO("Desea copiar el registro activo al portapapeles")=.T.
         TEXTO2:=""
         aCopiar:=W_DirImp1.GR_Dir1.Item(W_DirImp1.GR_Dir1.Value)
         FOR N1=1 TO LEN(aCopiar)
            DO CASE
            CASE VALTYPE(aCopiar[N1])="C"
               TEXTO2:=TEXTO2+aCopiar[N1]
            CASE VALTYPE(aCopiar[N1])="N"
               TEXTO2:=TEXTO2+STR(aCopiar[N1])
            CASE VALTYPE(aCopiar[N1])="D"
               TEXTO2:=TEXTO2+DTOC(aCopiar[N1])
            CASE VALTYPE(aCopiar[N1])="L"
               TEXTO2:=TEXTO2+IF(aCopiar[N1]=.T.,"Si","No")
            ENDCASE
            IF N1<LEN(aCopiar)
               TEXTO2:=TEXTO2+CHR(9)
            ENDIF
         NEXT
         CopyToClipboard(TEXTO2)
      ENDIF
   CASE LLAMADA="CopiarTabla"
      IF MSGYESNO("Desea copiar la tabla al portapapeles")=.T.
         TEXTO2:=""
         FOR N2=1 TO W_DirImp1.GR_Dir1.ItemCount
         aCopiar:=W_DirImp1.GR_Dir1.Item(N2)
            FOR N1=1 TO LEN(aCopiar)
               DO CASE
               CASE VALTYPE(aCopiar[N1])="C"
                  TEXTO2:=TEXTO2+aCopiar[N1]
               CASE VALTYPE(aCopiar[N1])="N"
                  TEXTO2:=TEXTO2+STR(aCopiar[N1])
               CASE VALTYPE(aCopiar[N1])="D"
                  TEXTO2:=TEXTO2+DTOC(aCopiar[N1])
               CASE VALTYPE(aCopiar[N1])="L"
                  TEXTO2:=TEXTO2+IF(aCopiar[N1]=.T.,"Si","No")
               ENDCASE
               IF N1<LEN(aCopiar)
                  TEXTO2:=TEXTO2+CHR(9)
               ENDIF
            NEXT
            TEXTO2:=TEXTO2+HB_OsNewLine()
         NEXT
         CopyToClipboard(TEXTO2)
      ENDIF
   ENDCASE
Return Nil

STATIC FUNCTION Grid1_Ord(LLAMADA)
IF W_DirImp1.GR_Dir1.ItemCount=0
   MSGSTOP("No hay registros para ordenar")
   RETURN
ENDIF
IF VALTYPE(W_DirImp1.GR_Dir1.Cell(1,LLAMADA))<>"C" .AND. ;
   VALTYPE(W_DirImp1.GR_Dir1.Cell(1,LLAMADA))<>"N" .AND. ;
   VALTYPE(W_DirImp1.GR_Dir1.Cell(1,LLAMADA))<>"D"
   MSGSTOP("Esta columna no se puede ordenar")
   RETURN
ENDIF
PonerEspera("Ordenando la tabla por "+W_DirImp1.GR_Dir1.Header(LLAMADA) )
   ESTOY:=IF(W_DirImp1.GR_Dir1.Value=0,1,W_DirImp1.GR_Dir1.Value)
   aOrdenar2:={}
   FOR N=1 TO W_DirImp1.GR_Dir1.ItemCount
      AADD(aOrdenar2,W_DirImp1.GR_Dir1.Item(N))
   NEXT
   DO CASE
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="C"
      ASORT(aOrdenar2,,, { |x, y| UPPER(x[LLAMADA]) < UPPER(y[LLAMADA]) })
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="N"
      ASORT(aOrdenar2,,, { |x, y| x[LLAMADA] < y[LLAMADA] })
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="D"
      ASORT(aOrdenar2,,, { |x, y| DTOS(x[LLAMADA]) < DTOS(y[LLAMADA]) })
   ENDCASE
   W_DirImp1.GR_Dir1.DeleteAllItems
   FOR N=1 TO LEN(aOrdenar2)
      W_DirImp1.GR_Dir1.AddItem(aOrdenar2[N])
   NEXT
   W_DirImp1.GR_Dir1.Value:=ESTOY
QuitarEspera()
Return Nil

STATIC FUNCTION Actualiz_NumEtiq()
   W_DirImp1.NumEtiq.Value:=LTRIM(STR(W_DirImp1.T_horizEtiq.Value*W_DirImp1.T_VertEtiq.Value))
   W_DirImp1.T_AnchoEtiq.Value:=ROUND((210-(W_DirImp1.T_IzqEtiq.Value*2))/W_DirImp1.T_horizEtiq.Value,1)
   W_DirImp1.T_AltoEtiq.Value:=ROUND((297-(W_DirImp1.T_SupEtiq.Value*2))/W_DirImp1.T_VertEtiq.Value,1)
Return Nil

STATIC FUNCTION Actualiz_SobreImp()
DO CASE
CASE W_DirImp1.R_Tipo.value=1 // Sobre
   W_DirImp1.T_DirMarIzq.Value:=080
   W_DirImp1.T_DirMarSup.Value:=100
   W_DirImp1.T_RteMarIzq.Value:=015
   W_DirImp1.T_RteMarSup.Value:=120
   aFontD[2]:=14
   W_DirImp1.R_Horizontal.value:=2
   SiVer1=.F.
CASE W_DirImp1.R_Tipo.value=2 // folio
   W_DirImp1.T_DirMarIzq.Value:=145
   W_DirImp1.T_DirMarSup.Value:=40
   aFontD[2]:=14
   W_DirImp1.R_Horizontal.value:=1
   SiVer1=.F.
CASE W_DirImp1.R_Tipo.value=3 // etiqueta
   W_DirImp1.T_DirMarIzq.Value:=5
   W_DirImp1.T_DirMarSup.Value:=5
   aFontD[2]:=10
   W_DirImp1.R_Horizontal.value:=1
   SiVer1=.T.
ENDCASE
   W_DirImp1.L_FuenteDirN.Value:=LTRIM(STR(aFontD[2]))

   W_DirImp1.C_Etiqueta.Enabled:=SiVer1
   W_DirImp1.T_IzqEtiq.Enabled:=SiVer1
   W_DirImp1.T_SupEtiq.Enabled:=SiVer1
   W_DirImp1.T_AnchoEtiq.Enabled:=.F.
   W_DirImp1.T_AltoEtiq.Enabled:=.F.
   W_DirImp1.T_horizEtiq.Enabled:=SiVer1
   W_DirImp1.T_VertEtiq.Enabled:=SiVer1
   W_DirImp1.T_Empezar.Enabled:=SiVer1
   W_DirImp1.C_SiLin.Enabled:=SiVer1

   W_DirImp1.T_RteMarIzq.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteMarSup.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.C_Remite.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom1.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom2.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom3.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom4.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom5.Enabled:=IF(SiVer1=.T.,.F.,.T.)
   W_DirImp1.T_RteNom6.Enabled:=IF(SiVer1=.T.,.F.,.T.)
Return Nil

STATIC FUNCTION Actualiz_EtiqImp()
   aEti2:=Etiqueta_Tipos("TODO")

   W_DirImp1.T_SupEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,8]
   W_DirImp1.T_IzqEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,9]

   W_DirImp1.T_horizEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,6]
   W_DirImp1.T_VertEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,7]
   W_DirImp1.NumEtiq.Value:=LTRIM(STR(aEti2[W_DirImp1.C_Etiqueta.Value,5]))

   W_DirImp1.T_AnchoEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,3]
   W_DirImp1.T_AltoEtiq.Value:=aEti2[W_DirImp1.C_Etiqueta.Value,4]

   W_DirImp1.T_DirMarIzq.Value:=IF(5-W_DirImp1.T_IzqEtiq.Value<2,2,5-W_DirImp1.T_IzqEtiq.Value)
   W_DirImp1.T_DirMarSup.Value:=IF(5-W_DirImp1.T_SupEtiq.Value<2,2,5-W_DirImp1.T_SupEtiq.Value)
Return Nil


STATIC FUNCTION SobreImpi()
   local oprint

   IF W_DirImp1.GR_Dir1.ItemCount=0
      MsgExclamation("No hay direcciones seleccionadas","Informacion")
      RETURN
   ENDIF

   SiHorizontal:=IF(W_DirImp1.R_Horizontal.value=1,.F.,.T.)
   IF W_DirImp1.R_Tipo.value=3 // 16-etiqueta direccion 105x37
      SiHorizontal:=.F.
   ENDIF

   oprint:=tprint(UPPER(W_DirImp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_DirImp1.nImp.value , W_DirImp1.nVer.value , SiHorizontal , 9 , W_DirImp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(W_DirImp1.Title)
   oprint:setpreviewsize(2)  // tamaño del preview
   oprint:beginpage()


NCOPIAS1:=IF(W_DirImp1.T_Copias.value>0,W_DirImp1.T_Copias.value,1)
aDir1:={}
FOR Ndir1=1 TO W_DirImp1.GR_Dir1.ItemCount
   FOR NCOPIAS2=1 TO NCOPIAS1
      AADD(aDir1,W_DirImp1.GR_Dir1.Item(Ndir1))
   NEXT
NEXT
aDir1:=ASORT(aDir1,,, {|x,y| UPPER(x[3])<y[3]} )
W_DirImp1.Progress_1.RangeMin:=0
W_DirImp1.Progress_1.RangeMax:=LEN(aDir1)

DO CASE
CASE W_DirImp1.R_Tipo.value<=2 // Sobre y folio

   PAG:=1
   FOR Ndir1=1 TO LEN(aDir1)
      W_DirImp1.Progress_1.Value:=Ndir1

      DO EVENTS

      LIN2:=W_DirImp1.T_DirMarSup.Value
      COL2:=W_DirImp1.T_DirMarIzq.Value
      SEP2:=ROUND(aFontD[2]/3,0)

      SEPMAS:=0
      IF W_DirImp1.C_ImpColCod.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,1],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColNom.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,2],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColDir.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,3],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColCPo.Value=.T. .OR. W_DirImp1.C_ImpColPob.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,4]+" "+aDir1[Ndir1,5],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColPrv.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,6],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColPai.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,7],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF

   IF W_DirImp1.C_Remite.Value=.T.
      LIN2:=W_DirImp1.T_RteMarSup.Value
      COL2:=W_DirImp1.T_RteMarIzq.Value
      SEP2:=ROUND(aFontR[2]/3,0)
      oprint:printdata(LIN2+(SEP2*0),COL2,"Remite",aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
      oprint:printdata(LIN2+(SEP2*1),COL2,W_DirImp1.T_RteNom1.Value,aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
      oprint:printdata(LIN2+(SEP2*2),COL2,W_DirImp1.T_RteNom2.Value,aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
      oprint:printdata(LIN2+(SEP2*3),COL2,W_DirImp1.T_RteNom3.Value+" "+W_DirImp1.T_RteNom4.Value,aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
      oprint:printdata(LIN2+(SEP2*4),COL2,W_DirImp1.T_RteNom5.Value,aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
      oprint:printdata(LIN2+(SEP2*5),COL2,W_DirImp1.T_RteNom6.Value,aFontR[1],aFontR[2],aFontR[3],aFontR[5],"L",)
   ENDIF

      IF Ndir1<LEN(aDir1)
         oprint:endpage()
         oprint:beginpage()
      ENDIF

   NEXT

CASE W_DirImp1.R_Tipo.value=3 // etiqueta

   IF W_DirImp1.T_Empezar.value<0 .OR. W_DirImp1.T_Empezar.value>VAL(W_DirImp1.NumEtiq.Value)
      W_DirImp1.T_Empezar.value:=1
   ENDIF
   IF W_DirImp1.T_Empezar.value>1
      nEmpezar:=W_DirImp1.T_Empezar.value-1
      LIN1:=INT(nEmpezar/W_DirImp1.T_horizEtiq.Value)+1
      COL1:=(nEmpezar%W_DirImp1.T_horizEtiq.Value)+1
   ELSE
      LIN1:=1
      COL1:=1
   ENDIF
   PAG:=1
   FOR Ndir1=1 TO LEN(aDir1)
      W_DirImp1.Progress_1.Value:=Ndir1

      DO EVENTS

      C1:=((COL1-1)*W_DirImp1.T_AnchoEtiq.Value)+W_DirImp1.T_IzqEtiq.Value
      COL2:=W_DirImp1.T_DirMarIzq.Value+C1
      L1:=((LIN1-1)*W_DirImp1.T_AltoEtiq.Value)+W_DirImp1.T_SupEtiq.Value
      LIN2:=W_DirImp1.T_DirMarSup.Value+L1
      SEP2:=ROUND(aFontD[2]/3,0)
      IF W_DirImp1.C_SiLin.Value=.T.
         oprint:printrectangle(L1,C1,L1+W_DirImp1.T_AltoEtiq.Value,C1+W_DirImp1.T_AnchoEtiq.Value, ,0.5 )
      ENDIF

      SEPMAS:=0
      IF W_DirImp1.C_ImpColCod.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,1],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColNom.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,2],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColDir.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,3],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColCPo.Value=.T. .OR. W_DirImp1.C_ImpColPob.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,4]+" "+aDir1[Ndir1,5],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColPrv.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,6],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF
      IF W_DirImp1.C_ImpColPai.Value=.T.
         oprint:printdata(LIN2+(SEP2*SEPMAS++),COL2,aDir1[Ndir1,7],aFontD[1],aFontD[2],aFontD[3],aFontD[5],"L",)
      ENDIF

      COL1++
      IF COL1=W_DirImp1.T_horizEtiq.Value+1
         COL1:=1
         LIN1++
         IF LIN1=W_DirImp1.T_vertEtiq.Value+1 .AND. Ndir1<LEN(aDir1)
            LIN1:=1
            COL1:=1
            IF Ndir1<LEN(aDir1)
               oprint:endpage()
               oprint:beginpage()
            ENDIF
         ENDIF
      ENDIF
   NEXT



ENDCASE


   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

*   W_DirImp1.release

Return Nil



FUNCTION Etiqueta_Tipos(LLAMADA)
/*
1-Referencia
2-descripcion
3-Ancho etiqueta
4-Alto etiqueta
5-Etiq. por hoja
6-Núm. etiqs. horiz.
7-Núm. etiqs. vert.
8-Margen superior
9-Margen izquierdo
10-Dif. entre etiqs. vert.
11-Dif. entre etiqs. horiz.
*/

aETIQ1:={}
AADD(aETIQ1,{'01209','65 Etiquetas 38x21,2'  , 38  , 21.2,65,5,13,10.4,10,0,0})
AADD(aETIQ1,{'10825','44 Etiquetas 48,5x25,4', 48.5, 25.4,44,4,11, 8.5, 8,0,0})
AADD(aETIQ1,{'10559','33 Etiquetas 70x25,4'  , 70  , 25.4,33,3,11, 8.5, 0,0,0})
AADD(aETIQ1,{'10560','27 Etiquetas 70x30'    , 70  , 30  ,27,3, 9,13.2, 0,0,0})
AADD(aETIQ1,{'01212','24 Etiquetas 70x37'    , 70  , 37  ,24,3, 8, 0  , 0,0,0})
AADD(aETIQ1,{'01299','20 Etiquetas 105x29'   ,105  , 29  ,20,2,10, 3.2, 0,0,0})
AADD(aETIQ1,{'01214','16 Etiquetas 105x37'   ,105  , 37  ,16,2, 8, 0  , 0,0,0})
AADD(aETIQ1,{'01289','12 Etiquetas 105x48'   ,105  , 48  ,12,2, 6, 4.2, 0,0,0})
AADD(aETIQ1,{'01213','12 Etiquetas 97x42,4'  , 97  , 42.4,12,2, 6,21  , 8,0,0})
AADD(aETIQ1,{'10826', '8 Etiquetas 97x67,7'  , 97  , 67.7, 8,2, 4,12.8, 8,0,0})
AADD(aETIQ1,{'10827', '4 Etiquetas 105x148'  ,105  ,148  , 4,2, 2, 0  , 0,0,0})
AADD(aETIQ1,{'10919', '2 Etiquetas 210x148'  ,210  ,148  , 2,1, 2, 0  , 0,0,0})
AADD(aETIQ1,{'01215', '1 Etiquetas 297x210'  ,297  ,210  , 1,1, 1, 0  , 0,0,0})
AADD(aETIQ1,{'00000', 'Personalizado'        ,297  ,210  , 1,1, 1, 0  , 0,0,0})

aETIQ2:={}
FOR N=1 TO LEN(aETIQ1)
   AADD(aETIQ2,aETIQ1[N,2])
NEXT

IF LLAMADA="TODO"
   RETURN(aETIQ1)
ELSE
   RETURN(aETIQ2)
ENDIF
