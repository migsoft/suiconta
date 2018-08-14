#include "minigui.ch"

Function br_rec1(CODIGO1,VENTANA1,CONTROL1)
   LOCAL Color_rec1 := { || IIF( PEND<>0 , RGB(255,200,200), RGB(255,255,255)) }
   CODRECBR:=1
   AbrirDBF("FACREB",,,,,1)
   SEEK CODIGO1
   IF .NOT. EOF()
      CODRECBR:=RECNO()
   ELSE
      GO BOTT
      CODRECBR:=RECNO()
   ENDIF

   DEFINE WINDOW WinBRrec1 ;
      AT 0,0     ;
      WIDTH 700 HEIGHT 470 ;
      TITLE 'Facturas recibidas' ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("VERDECLARO")

      @ 10,10 BROWSE br_rec1 ;
      WIDTH 676 HEIGHT 380 ;
      HEADERS {'Registro','Fecha','Nombre','Factura','Importe','Pendiente','Asiento'} ;
      WIDTHS { 55,80,175,70,100,100,75 } ;
      WORKAREA FACREB ;
      FIELDS { 'FACREB->NREG','FACREB->FREG','FACREB->NOMCTA','FACREB->REF','LTRIM(MIL(FACREB->TFAC,15,2))','LTRIM(MIL(FACREB->PEND,15,2))','ASIENTO' } ;
      DYNAMICBACKCOLOR { Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1 } ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE CODRECBR ;
      ON DBLCLICK br_rec1Fin(VENTANA1,CONTROL1) ;
      ON HEADCLICK { {|| AbrirDBF("FACREB",,,,,1), WinBRrec1.br_rec1.Refresh} , , ;
                     {|| AbrirDBF("FACREB",,,,,2), WinBRrec1.br_rec1.Refresh} }

      DEFINE CONTEXT MENU CONTROL br_rec1 OF WinBRrec1
         ITEM "Copiar registro"            ACTION br_rec1_CopReg("REG1")
         ITEM "Copiar registros proveedor" ACTION br_rec1_CopReg("REGPROV")
      END MENU

   @400,110 TEXTBOX T_BuscarNom WIDTH 250 VALUE '' MAXLENGTH 30 ;
           ON CHANGE { || br_rec1Bus() }

   @400,370 RADIOGROUP R_TipLoc OPTIONS { 'Nombre','Factura' } ;
           VALUE 1 WIDTH 80 HORIZONTAL

   @400,600 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
      ACTION WinBRrec1.Release ;
      WIDTH 90 HEIGHT 25 NOTABSTOP

   @400,010 BUTTONEX Bt_Localizar CAPTION 'Localizar' ICON icobus('buscar') ;
         ACTION br_rec1Loc() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

   WinBRrec1.T_BuscarNom.SetFocus

   END WINDOW
   VentanaCentrar("WinBRrec1","Ventana1")
   CENTER WINDOW WinBRrec1
   ACTIVATE WINDOW WinBRrec1

Return Nil

STATIC FUNCTION br_rec1Bus()
   AbrirDBF("FACREB")
   SET SOFTSEEK ON
   IF VAL(WinBRrec1.T_BuscarNom.Value)<>0
      DBSETORDER(1)
      SEEK VAL(WinBRrec1.T_BuscarNom.Value)
   ELSE
      DBSETORDER(2)
      SEEK UPPER(WinBRrec1.T_BuscarNom.Value)
   ENDIF
   SET SOFTSEEK OFF
   WinBRrec1.br_rec1.Value := Recno()
   WinBRrec1.br_rec1.Refresh
Return Nil

STATIC FUNCTION br_rec1Loc()

   NOMBUSCAR:=RTRIM(WinBRrec1.T_BuscarNom.Value)
   IF LEN(NOMBUSCAR)=0
      RETURN
   ENDIF
   AbrirDBF("FACREB")
   IF WinBRrec1.br_rec1.Value=0
      GO TOP
      WinBRrec1.br_rec1.Value:=RECNO()
   ENDIF
   GO WinBRrec1.br_rec1.Value
   DO WHILE .T.
      DO EVENTS
      SKIP
      IF EOF()
         GO TOP
      ENDIF
      DO CASE
      CASE WinBRrec1.R_TipLoc.Value=1
         IF AT(UPPER(NOMBUSCAR),UPPER(NOMCTA))<>0
            WinBRrec1.br_rec1.Value:=recno()
            WinBRrec1.br_rec1.Refresh
            EXIT
         ENDIF
      CASE WinBRrec1.R_TipLoc.Value=2
         IF AT(UPPER(NOMBUSCAR),UPPER(REF))<>0
            WinBRrec1.br_rec1.Value:=recno()
            WinBRrec1.br_rec1.Refresh
            EXIT
         ENDIF
      ENDCASE
      IF RECNO()=WinBRrec1.br_rec1.Value
         MSGSTOP("No se ha localizado ningun registro")
         EXIT
      ENDIF
   ENDDO
Return Nil

STATIC FUNCTION br_rec1Fin(VENTANA1,CONTROL1)
   AbrirDBF("FACREB")
   GO WinBRrec1.br_rec1.Value
   SetProperty(VENTANA1,CONTROL1,"Value",NREG)
   WinBRrec1.Release
Return Nil

STATIC FUNCTION br_rec1_CopReg(LLAMADA)
   TEXTO2:=""
   AbrirDBF("FACREB")
   GO WinBRrec1.br_rec1.Value
   DO CASE
   CASE LLAMADA="REG1"
      TEXTO2:=LTRIM(STR(NREG))+CHR(9)+LTRIM(STR(CODIGO))+CHR(9)+RTRIM(NOMCTA)+CHR(9)+ ;
              DIA(FREG)+CHR(9)+RTRIM(REF)+CHR(9)+STRTRAN(LTRIM(STR(TFAC)),".",",")+CHR(9)+STRTRAN(LTRIM(STR(PEND)),".",",")
   CASE LLAMADA="REGPROV"
      PonerEspera("Copiando datos al portapapeles")
      CODPROV2:=CODIGO
      DO WHILE CODIGO=CODPROV2 .AND. .NOT. BOF()
         SKIP -1
      ENDDO
      SKIP
      DO WHILE CODIGO=CODPROV2 .AND. .NOT. EOF()
         TEXTO2:=TEXTO2+LTRIM(STR(NREG))+CHR(9)+LTRIM(STR(CODIGO))+CHR(9)+RTRIM(NOMCTA)+CHR(9)+ ;
                 DIA(FREG)+CHR(9)+RTRIM(REF)+CHR(9)+STRTRAN(LTRIM(STR(TFAC)),".",",")+CHR(9)+STRTRAN(LTRIM(STR(PEND)),".",",")+HB_OsNewLine()
         SKIP
      ENDDO
      QuitarEspera()
   ENDCASE
//*** CopyToClipboard(TEXTO2)
SetClipboardText(TEXTO2)
MsgInfo("El texto se ha copiado al portapapeles")
Return Nil
