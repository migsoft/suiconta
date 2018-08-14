#include "minigui.ch"

Function br_suizocue(RUTA1,CODIGO1,VENTANA1,CONTROL1,CONTROL2)
   VENTANA1:=IF(VENTANA1=NIL,"",VENTANA1)
   CONTROL1:=IF(CONTROL1=NIL,"",CONTROL1)
   CONTROL2:=IF(CONTROL2=NIL,"",CONTROL2)
   CODALUBR:=1
   CODIGOFIN:=""
   DO CASE
   CASE AT("SUIZO",UPPER(RUTA1))<>0
      RUTA2:=RTRIM(RUTA1)
      IF FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
         FILE(RUTA2+"\CUENTAS.CDX")
         AbrirDBF("CUENTAS",,,,RUTA2)
         SEEK CODIGO1
         IF .NOT. EOF()
            CODALUBR:=RECNO()
         ENDIF
      ELSE
         RETURN(CODIGOFIN)
      ENDIF
   CASE AT("GRUPOSP",UPPER(RUTA1))<>0
      RUTA2:=RTRIM(RUTA1)
      IF FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
         FILE(RUTA2+"\SUBCTA.CDX")
         AbrirDBF("SUBCTA",,,,RUTA2)
         DBSETORDER(1)
         SEEK PADR(LTRIM(STR(CODIGO1)),LEN(COD)," ")
         IF .NOT. EOF()
            CODALUBR:=RECNO()
         ENDIF
      ELSE
         RETURN(CODIGOFIN)
      ENDIF
   OTHERWISE
      RETURN(CODIGOFIN)
   ENDCASE

   DBSETORDER(2)


   DEFINE WINDOW WinBRcue1 ;
      AT 0,0     ;
      WIDTH 500  ;
      HEIGHT 500 ;
      TITLE "Consulta de cuentas" ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("AZULCLARO") ;
      ON RELEASE Terminar_CodSui2()

   @ 10,10 LABEL L_Ruta1 VALUE RUTA1 AUTOSIZE TRANSPARENT

   IF SELEC("CUENTAS")<>0
      @ 40,10 BROWSE BR_Cue1 ;
      HEIGHT 350 ;
      WIDTH 475 ;
      HEADERS {'Codigo','Nombre','Saldo'} ;
      WIDTHS { 100,250,100 } ;
      WORKAREA CUENTAS ;
      FIELDS { 'CUENTAS->CODCTA','CUENTAS->NOMCTA','LTRIM(MIL(CUENTAS->SALDO,15,2))' } ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT} ;
      VALUE CODALUBR ;
      ON DBLCLICK Terminar_CodSui1(VENTANA1,CONTROL1,CONTROL2) ;
      ON HEADCLICK { {|| AbrirDBF("CUENTAS",,,,RUTA2),DBSETORDER(1), WinBRcue1.BR_Cue1.Refresh} , ;
                     {|| AbrirDBF("CUENTAS",,,,RUTA2),DBSETORDER(2), WinBRcue1.BR_Cue1.Refresh} }
   ELSE
      @ 40,10 BROWSE BR_Cue1 ;
      HEIGHT 350 ;
      WIDTH 475 ;
      HEADERS {'Codigo','Nombre','Saldo'} ;
      WIDTHS { 100,250,100 } ;
      WORKAREA SUBCTA ;
      FIELDS { 'SUBCTA->COD','SUBCTA->TITULO','LTRIM(MIL(SUBCTA->SUMADBEU-SUBCTA->SUMAHBEU,15,2))' } ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT} ;
      VALUE CODALUBR ;
      ON DBLCLICK Terminar_CodSui1(VENTANA1,CONTROL1,CONTROL2) ;
      ON HEADCLICK { {|| AbrirDBF("SUBCTA",,,,RUTA2),DBSETORDER(1), WinBRcue1.BR_Cue1.Refresh} , ;
                     {|| AbrirDBF("SUBCTA",,,,RUTA2),DBSETORDER(2), WinBRcue1.BR_Cue1.Refresh} }
   ENDIF

   @400,010 BUTTONEX Bt_Localizar CAPTION 'Localizar' ICON icobus('buscar') ;
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

RETURN(CODIGOFIN)

PROCEDURE Bus_SuiCodCta()
   IF SELEC("CUENTAS")<>0
      SELEC CUENTAS
   ELSE
      SELEC SUBCTA
   ENDIF
   SET SOFTSEEK ON
   IF SELEC("CUENTAS")<>0
      IF VAL(PCODCTA3(WinBRcue1.T_BuscarNom.Value))<>0
         DBSETORDER(1)
         CODCTA2:=VAL(PCODCTA3(WinBRcue1.T_BuscarNom.Value))
         SEEK CODCTA2
      ELSE
         DBSETORDER(2)
         SEEK UPPER(WinBRcue1.T_BuscarNom.Value)
      ENDIF
   ELSE
      IF VAL(PCODCTA3(WinBRcue1.T_BuscarNom.Value))<>0
         DBSETORDER(1)
         CODCTA2:=PCODCTA3(WinBRcue1.T_BuscarNom.Value)
         SEEK CODCTA2
      ELSE
         DBSETORDER(2)
         SEEK UPPER(WinBRcue1.T_BuscarNom.Value)
      ENDIF
   ENDIF
   SET SOFTSEEK OFF
   WinBRcue1.BR_Cue1.Value := Recno()
   WinBRcue1.BR_Cue1.Refresh

STATIC FUNCTION Terminar_CodSui1(VENTANA1,CONTROL1,CONTROL2)
   GO WinBRcue1.BR_Cue1.Value
   DO CASE
   CASE SELEC("CUENTAS")<>0
      IF IsControlDefined(&CONTROL1,&VENTANA1)=.T.
         SetProperty(VENTANA1,CONTROL1,"Value",LTRIM(STR(CODCTA)))
      ENDIF
      IF IsControlDefined(&CONTROL2,&VENTANA1)=.T.
         SetProperty(VENTANA1,CONTROL2,"Value",NOMCTA)
      ENDIF
      CODIGOFIN:=LTRIM(STR(CODCTA))
   CASE SELEC("SUBCTA")<>0
      IF IsControlDefined(&CONTROL1,&VENTANA1)=.T.
         SetProperty(VENTANA1,CONTROL1,"Value",LTRIM(COD))
      ENDIF
      IF IsControlDefined(&CONTROL2,&VENTANA1)=.T.
         SetProperty(VENTANA1,CONTROL2,"Value",TITULO)
      ENDIF
      CODIGOFIN:=LTRIM(COD)
   ENDCASE
   Terminar_CodSui2()
   WinBRcue1.Release

STATIC FUNCTION Terminar_CodSui2()
   IF SELEC("CUENTAS")<>0
      CUENTAS->( DBCLOSEAREA() )
   ENDIF
   IF SELEC("SUBCTA")<>0
      SUBCTA->( DBCLOSEAREA() )
   ENDIF
RETURN


FUNCTION BusCtaSuizo(RutaSuizo2,CtaSuizo2,Ventana2,Control2)
   IF VAL(CtaSuizo2)=0
      SetProperty(Ventana2,Control2,"Value","")
   ENDIF
   IF AT("SUIZO",UPPER(RutaSuizo2))<>0
      RUTA2:=RTRIM(RutaSuizo2)
      IF FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
         FILE(RUTA2+"\CUENTAS.CDX")
         AbrirDBF("CUENTAS",,,,RUTA2,1)
         SEEK VAL(CtaSuizo2)
         IF .NOT. EOF()
            SetProperty(Ventana2,Control2,"Value",CUENTAS->NOMCTA)
         ENDIF
         IF SELEC("CUENTAS")<>0
            CUENTAS->( DBCLOSEAREA() )
         ENDIF
      ENDIF
   ENDIF
   IF AT("GRUPOSP",UPPER(RutaSuizo2))<>0
      RUTA2:=RTRIM(RutaSuizo2)
      IF FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
         FILE(RUTA2+"\SUBCTA.CDX")
         AbrirDBF("SUBCTA",,,,RUTA2,1)
         SEEK PADR(CtaSuizo2,LEN(COD)," ")
         IF .NOT. EOF()
            SetProperty(Ventana2,Control2,"Value",SUBCTA->TITULO)
         ENDIF
         IF SELEC("SUBCTA")<>0
            SUBCTA->( DBCLOSEAREA() )
         ENDIF
      ENDIF
   ENDIF

STATIC FUNCTION Localizar_Cta()
   NOMBUSCAR:=RTRIM(WinBRcue1.T_BuscarNom.Value)
   IF LEN(NOMBUSCAR)=0
      RETURN
   ENDIF
   IF SELEC("CUENTAS")<>0
      AbrirDBF("CUENTAS",,,,RUTA2)
      GO WinBRcue1.BR_Cue1.Value
      DO WHILE .T.
         DO EVENTS
         SKIP
         IF EOF()
            GO TOP
         ENDIF
         IF AT(UPPER(NOMBUSCAR),UPPER(CUENTAS->NOMCTA))<>0
            WinBRcue1.BR_Cue1.Value:=recno()
            WinBRcue1.BR_Cue1.Refresh
            EXIT
         ENDIF
         IF RECNO()=WinBRcue1.BR_Cue1.Value
            MSGSTOP("No se ha localizado ningun registro")
            EXIT
         ENDIF
      ENDDO
   ELSE
      AbrirDBF("SUBCTA",,,,RUTA2)
      GO WinBRcue1.BR_Cue1.Value
      DO WHILE .T.
         DO EVENTS
         SKIP
         IF EOF()
            GO TOP
         ENDIF
         IF AT(UPPER(NOMBUSCAR),UPPER(SUBCTA->TITULO))<>0
            WinBRcue1.BR_Cue1.Value:=recno()
            WinBRcue1.BR_Cue1.Refresh
            EXIT
         ENDIF
         IF RECNO()=WinBRcue1.BR_Cue1.Value
            MSGSTOP("No se ha localizado ningun registro")
            EXIT
         ENDIF
      ENDDO
   ENDIF
