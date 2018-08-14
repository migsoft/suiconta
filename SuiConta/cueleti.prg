#include "minigui.ch"
#include "winprint.ch"

procedure Cueleti()
   TituloImp:="Listado de etiquetas"

   AbrirDBF("Cuentas")
   DBSETORDER(2)

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 800 HEIGHT 420 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @015,10 LABEL L_CodCta1 VALUE 'Desde el codigo' AUTOSIZE TRANSPARENT
      @010,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @010,225 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION ( br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") , ;
                              W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value) ) ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @045,10 LABEL L_CodCta2 VALUE 'Hasta el codigo' AUTOSIZE TRANSPARENT
      @040,110 TEXTBOX T_CodCta2 WIDTH 100 VALUE "99999999" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta2.Value)
      @040,225 BUTTONEX Bt_BuscarCue2 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta2.Value),"W_Imp1","T_CodCta2") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP


      @ 10,590 BUTTONEX B_Actualizar CAPTION 'Actializar lista' WIDTH 90 HEIGHT 25 ;
               ACTION CueletiSel("TODO")

      @ 10,690 BUTTONEX B_Borrar CAPTION 'Borrar lista' WIDTH 90 HEIGHT 25 ;
               ACTION CueletiSel("BORRARTODO")

      @ 45,590 LABEL L_TotLis VALUE 'Registros:' AUTOSIZE TRANSPARENT


      @70,10 BROWSE BR_Cue1 ;
      HEIGHT 270 ;
      WIDTH 380 ;
      HEADERS {'Codigo','Nombre','Pob.'} ;
      WIDTHS { 75 , 210 , 65 } ;
      WORKAREA Cuentas ;
      FIELDS { 'Cuentas->CODCTA' , 'Cuentas->NOMCTA' , "Cuentas->POBCTA"} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT} ;
      VALUE 1 ;
      ON DBLCLICK CueletiSel("AÑADIR") ;
      ON HEADCLICK { {|| AbrirDBF("Cuentas"),DBSETORDER(1), W_Imp1.BR_Cue1.Refresh} , ;
                     {|| AbrirDBF("Cuentas"),DBSETORDER(2), W_Imp1.BR_Cue1.Refresh} }

      @100,390 BUTTON B_Mas1 CAPTION '->' WIDTH 20 HEIGHT 25 ACTION CueletiSel("AÑADIR")
      @130,390 BUTTON B_Mas2 CAPTION '<-' WIDTH 20 HEIGHT 25 ACTION CueletiSel("BORRAR")

      @ 70,410 GRID GR_Cue1 ;
      HEIGHT 270 ;
      WIDTH 380 ;
      HEADERS {'Codigo','Nombre','Direccion','Cod.Postal','Poblacion','Provincia','Pais'} ;
      WIDTHS { 75 , 210 ,0,0, 65,0,0 } ;
      ITEMS {} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,,,,} ;
      ON DBLCLICK CueletiSel("BORRAR") ;
      VALUE 1

      @340,410 PROGRESSBAR Prog_FacL WIDTH 380 HEIGHT 10 SMOOTH

      @355,10 LABEL L_BuscarNom VALUE 'Buscar' AUTOSIZE TRANSPARENT
      @350,70 TEXTBOX T_BuscarNom WIDTH 250 VALUE '' ;
              MAXLENGTH 30 ;
              ON CHANGE Buscar_Alu()

      @360,600 BUTTONEX B_Sel CAPTION 'Seleccionar' ICON icobus('consultar') WIDTH 90 HEIGHT 25 ;
               ACTION Cueletii()

      @360,700 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION Cueletii()
   IF W_Imp1.GR_Cue1.ItemCount=0
      MsgExclamation("No hay datos seleccionados","Informacion")
      RETURN
   ENDIF

   aDireccionImp:={}
   W_Imp1.Prog_FacL.RangeMin:=0
   W_Imp1.Prog_FacL.RangeMax:=W_Imp1.GR_Cue1.ItemCount
   FOR Nf=1 TO W_Imp1.GR_Cue1.ItemCount
      W_Imp1.Prog_FacL.Value:=Nf
      AADD(aDireccionImp,W_Imp1.GR_Cue1.Item(Nf))
   NEXT

AbrirDBF("empresa",,,,RUTAPROGRAMA)
GO RECNOEMPRESA
aDireccionRte:={ ;
    RTRIM(NOMEMP), ;
    RTRIM(Direccion), ;
    RTRIM(CodPos), ;
    RTRIM(Poblacion), ;
    RTRIM(Provincia), ;
    '' }

DireccionImp(1,aDireccionImp,aDireccionRte)

Return Nil



STATIC FUNCTION CueletiSel(LLAMADA)
   DO CASE
   CASE LLAMADA="TODO"
*      W_Imp1.GR_Cue1.DeleteAllItems //PODER IMPRIMIR VARIAS CUENTAS
      AbrirDBF("Cuentas")
      W_Imp1.Prog_FacL.RangeMin:=0
      W_Imp1.Prog_FacL.RangeMax:=LASTREC()
      W_Imp1.Prog_FacL.Value:=0
      NUM2:=1
      GO TOP
      DO WHILE .NOT. EOF()
         DO EVENTS
         W_Imp1.Prog_FacL.Value:=NUM2++
         IF CODCTA>=VAL(W_Imp1.T_CodCta1.Value) .AND. CODCTA<=VAL(W_Imp1.T_CodCta2.Value)
            W_Imp1.GR_Cue1.AddItem({ LTRIM(STR(CODCTA)), ;
               RTRIM(NOMCTA),RTRIM(DIRCTA),RTRIM(CODPOS),RTRIM(POBCTA),RTRIM(PROCTA),RTRIM(PAIS)})
         ENDIF
         SKIP
      ENDDO
      W_Imp1.GR_Cue1.Refresh
   CASE LLAMADA="AÑADIR"
      AbrirDBF("Cuentas")
      GO W_Imp1.BR_Cue1.Value
      W_Imp1.GR_Cue1.AddItem({ LTRIM(STR(CODCTA)), ;
         RTRIM(NOMCTA),RTRIM(DIRCTA),RTRIM(CODPOS),RTRIM(POBCTA),RTRIM(PROCTA),RTRIM(PAIS)})
      W_Imp1.GR_Cue1.Value:=W_Imp1.GR_Cue1.ItemCount
      W_Imp1.GR_Cue1.Refresh
   CASE LLAMADA=="BORRAR"
      W_Imp1.GR_Cue1.DeleteItem(W_Imp1.GR_Cue1.Value)
      W_Imp1.GR_Cue1.Refresh
   CASE LLAMADA=="BORRARTODO"
      W_Imp1.Prog_FacL.Value:=0
      W_Imp1.GR_Cue1.DeleteAllItems
      W_Imp1.GR_Cue1.Refresh
   ENDCASE
   W_Imp1.L_TotLis.Value:="Registros: "+LTRIM(STR(W_Imp1.GR_Cue1.ItemCount))

STATIC FUNCTION Buscar_Alu()
   AbrirDBF("Cuentas")
   IF PCODCTA1(W_Imp1.T_BuscarNom.Value)<>0
      DBSETORDER(1)
      W_Imp1.BR_Cue1.Refresh
      Cuentas->(DBSeek( PCODCTA1(W_Imp1.T_BuscarNom.Value) ))
   ELSE
      DBSETORDER(2)
      W_Imp1.BR_Cue1.Refresh
      SET SOFTSEEK ON
      Cuentas->(DBSeek(UPPER(W_Imp1.T_BuscarNom.Value)))
      SET SOFTSEEK OFF
   ENDIF
   W_Imp1.BR_Cue1.Value := RECNO()
   W_Imp1.BR_Cue1.Refresh



