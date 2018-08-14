#include "minigui.ch"

Function br_cue1(CODIGO1,VENTANA1,CONTROL1,CONTROL2)
	IF VALTYPE(CODIGO1)="C"
		CODIGO2:=VAL(CODIGO1)
	ELSE
		CODIGO2:=CODIGO1
	ENDIF
   CODALUBR:=1
   AbrirDBF("CUENTAS",,,,,1)
   SEEK CODIGO2
   IF .NOT. EOF()
      CODALUBR:=RECNO()
   ENDIF

   DBSETORDER(2)

   DEFINE WINDOW WinBRcue1 ;
      AT 0,0     ;
      WIDTH 500 HEIGHT 460 ;
      TITLE 'Subcuentas' ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("AZULCLARO")

      @ 10,10 BROWSE BR_Cue1 ;
      HEIGHT 380 ;
      WIDTH 475 ;
      HEADERS {'Codigo','Nombre','Saldo'} ;
      WIDTHS { 100,250,100 } ;
      WORKAREA CUENTAS ;
      FIELDS { 'CUENTAS->CODCTA','CUENTAS->NOMCTA','LTRIM(MIL(CUENTAS->SALDO,15,2))' } ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT} ;
      VALUE CODALUBR ;
      ON DBLCLICK Terminar_CodSui1(VENTANA1,CONTROL1,CONTROL2) ;
      ON HEADCLICK { {|| AbrirDBF("CUENTAS"),DBSETORDER(1), WinBRcue1.BR_Cue1.Refresh} , ;
                     {|| AbrirDBF("CUENTAS"),DBSETORDER(2), WinBRcue1.BR_Cue1.Refresh} }

   @400,010 BUTTONEX Bt_Localizar ;
         CAPTION 'Localizar' ;
         ICON icobus('buscar') ;
         ACTION Localizar_Cta() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

   @400,110 TEXTBOX T_BuscarNom ;
           WIDTH 250 ;
           VALUE '' ;
           TOOLTIP 'Buscar subcuenta' ;
           MAXLENGTH 30 ;
           ON CHANGE { || Bus_SuiCodCta() }

   WinBRcue1.T_BuscarNom.SetFocus

   END WINDOW
   VentanaCentrar("WinBRcue1","Ventana1")
   ACTIVATE WINDOW WinBRcue1

Return

STATIC FUNCTION Bus_SuiCodCta()
   AbrirDBF("CUENTAS")
   SET SOFTSEEK ON
   IF VAL(PCODCTA3(WinBRcue1.T_BuscarNom.Value))<>0
      DBSETORDER(1)
      CODCTA2:=VAL(PCODCTA3(WinBRcue1.T_BuscarNom.Value))
      SEEK CODCTA2
   ELSE
      DBSETORDER(2)
      SEEK UPPER(WinBRcue1.T_BuscarNom.Value)
   ENDIF
   SET SOFTSEEK OFF
   WinBRcue1.BR_Cue1.Value := Recno()
   WinBRcue1.BR_Cue1.Refresh


STATIC FUNCTION Localizar_Cta()
   NOMBUSCAR:=RTRIM(WinBRcue1.T_BuscarNom.Value)
   IF LEN(NOMBUSCAR)=0
      RETURN
   ENDIF
   AbrirDBF("CUENTAS")
   GO WinBRcue1.BR_Cue1.Value
   DO WHILE .T.
      DO EVENTS
      SKIP
      IF EOF()
         GO TOP
      ENDIF
      IF AT(UPPER(NOMBUSCAR),UPPER(NOMCTA))<>0
         WinBRcue1.BR_Cue1.Value:=recno()
         WinBRcue1.BR_Cue1.Refresh
         EXIT
      ENDIF
      IF RECNO()=WinBRcue1.BR_Cue1.Value
         MSGSTOP("No se ha localizado ningun registro")
         EXIT
      ENDIF
   ENDDO


STATIC FUNCTION Terminar_CodSui1(VENTANA1,CONTROL1,CONTROL2)
   CONTROL2:=IF(CONTROL2=NIL,"",CONTROL2)
   AbrirDBF("CUENTAS")
   GO WinBRcue1.BR_Cue1.Value
   SetProperty(VENTANA1,CONTROL1,"Value",LTRIM(STR(CODCTA)))
   IF IsControlDefined(&CONTROL2,&VENTANA1)=.T.
      SetProperty(VENTANA1,CONTROL2,"Value",RTRIM(NOMCTA))
   ENDIF
   WinBRcue1.Release



FUNCTION Codigos_NomCta(Cod_Codigo)
   NUMSEL:=SELECT()
   AbrirDBF("CUENTAS")
   ESTOYCLI2:=RECNO()
   SEEK VAL(Cod_Codigo)
   DO CASE
   CASE VAL(Cod_Codigo)=0
      Nom_Codigo:=""
   CASE EOF()
      Nom_Codigo:="Codigo no dado de alta"
   OTHERWISE
      Nom_Codigo:=NOMCTA
   ENDCASE
   GO ESTOYCLI2
   SELEC(NUMSEL)

RETURN(Nom_Codigo)

