#include "minigui.ch"

PROCEDURE CueCamb()

IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
   RETURN
ENDIF

   DEFINE WINDOW W_CueCamb ;
      AT 10,10 ;
      WIDTH 600 HEIGHT 400 ;
      TITLE 'Cambiar codigo cuenta' ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @010,10 BUTTONEX Bt_BuscarCue1 CAPTION 'Codigo Actual' ICON icobus('buscar') ;
         ACTION ( br_cue1(VAL(W_CueCamb.T_CodAct.Value),"W_CueCamb","T_CodAct","T_NomAct") , CueCamb_Bus() ) ;
         WIDTH 110 HEIGHT 25 NOTABSTOP
      @010,130 TEXTBOX T_CodAct WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS ( W_CueCamb.T_CodAct.Value:=PCODCTA3(W_CueCamb.T_CodAct.Value) , CueCamb_Bus() )
      @010,245 TEXTBOX T_NomAct WIDTH 250 VALUE "" READONLY NOTABSTOP

      @040,10 BUTTONEX Bt_BuscarCue2 CAPTION 'Codigo Futuro' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_CueCamb.T_CodFut.Value),"W_CueCamb","T_CodFut","T_NomFut") ;
         WIDTH 110 HEIGHT 25 NOTABSTOP
      @040,130 TEXTBOX T_CodFut WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_CueCamb.T_CodFut.Value:=PCODCTA3(W_CueCamb.T_CodFut.Value)
      @040,245 TEXTBOX T_NomFut WIDTH 250 VALUE "" READONLY NOTABSTOP


      @105,10 LABEL L_Fic1 VALUE 'Subcuentas' AUTOSIZE TRANSPARENT
      @135,10 LABEL L_Fic2 VALUE 'Asientos' AUTOSIZE TRANSPARENT
      @165,10 LABEL L_Fic3 VALUE 'Cierre' AUTOSIZE TRANSPARENT
      @195,10 LABEL L_Fic4 VALUE 'Facturas emitidas' AUTOSIZE TRANSPARENT
      @225,10 LABEL L_Fic5 VALUE 'Facturas recibidas' AUTOSIZE TRANSPARENT
      @255,10 LABEL L_Fic6 VALUE 'Vencimientos Fac.Rec.' AUTOSIZE TRANSPARENT
      @285,10 LABEL L_Fic7 VALUE 'Remesas' AUTOSIZE TRANSPARENT

      @075,150 LABEL L_Loc VALUE 'Localizados' AUTOSIZE TRANSPARENT
      @100,150 TEXTBOX T_LocFic1 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @130,150 TEXTBOX T_LocFic2 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @160,150 TEXTBOX T_LocFic3 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @190,150 TEXTBOX T_LocFic4 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @220,150 TEXTBOX T_LocFic5 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @250,150 TEXTBOX T_LocFic6 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @280,150 TEXTBOX T_LocFic7 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP

      @075,250 LABEL L_Mod VALUE 'Modificados' AUTOSIZE TRANSPARENT
      @100,250 TEXTBOX T_ModFic1 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @130,250 TEXTBOX T_ModFic2 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @160,250 TEXTBOX T_ModFic3 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @190,250 TEXTBOX T_ModFic4 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @220,250 TEXTBOX T_ModFic5 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @250,250 TEXTBOX T_ModFic6 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP
      @280,250 TEXTBOX T_ModFic7 WIDTH 50 VALUE 0 READONLY NUMERIC RIGHTALIGN NOTABSTOP

      @100,310 PROGRESSBAR Progres1 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @130,310 PROGRESSBAR Progres2 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @160,310 PROGRESSBAR Progres3 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @190,310 PROGRESSBAR Progres4 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @220,310 PROGRESSBAR Progres5 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @250,310 PROGRESSBAR Progres6 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH
      @280,310 PROGRESSBAR Progres7 RANGE 0 , 100 WIDTH 250 HEIGHT 25 SMOOTH

      @310,010 CHECKBOX SiFicCta CAPTION 'Modificar el codigo en la ficha de subcuentas'  WIDTH 270 VALUE .T.
      @340,010 CHECKBOX SiEjercicios CAPTION 'Cambiar en ejercicios anteriores' WIDTH 200 VALUE .T.

      @340,410 BUTTONEX B_Cambiar CAPTION 'Cambiar' WIDTH 90 HEIGHT 25 ACTION CueCamb_Cam()

      @340,510 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 80 HEIGHT 25 ;
               ACTION W_CueCamb.release

      END WINDOW
      VentanaCentrar("W_CueCamb","Ventana1")
      ACTIVATE WINDOW W_CueCamb

Return Nil

STATIC FUNCTION CueCamb_Bus()
FOR N=1 TO 7
   DO CASE
   CASE N=1
      AbrirDBF("CUENTAS")
      CAMPO:="CODCTA"
   CASE N=2
      AbrirDBF("APUNTES")
      CAMPO:="CODCTA"
   CASE N=3
      AbrirDBF("CIERRE")
      CAMPO:="CODCTA"
   CASE N=4
      AbrirDBF("FAC92")
      CAMPO:="CODCTA"
   CASE N=5
      AbrirDBF("FACREB")
      CAMPO:="CODIGO"
   CASE N=6
      AbrirDBF("FACVTO")
      CAMPO:="CODIGO"
   CASE N=7
      AbrirDBF("REMESA")
      CAMPO:="CODCTA"
   ENDCASE

   IF LEN(RTRIM(W_CueCamb.T_CodAct.Value))<>0
	   COUNT TO CODHAY FOR &CAMPO=VAL(W_CueCamb.T_CodAct.Value)
   ELSE
      CODHAY:=0
   ENDIF
   SetProperty("W_CueCamb","T_LocFic"+LTRIM(STR(N)),"Value",CODHAY)
   SetProperty("W_CueCamb","T_ModFic"+LTRIM(STR(N)),"Value", 0 )

NEXT

Return Nil







STATIC FUNCTION CueCamb_Cam()

   IF VAL(W_CueCamb.T_CodFut.Value)=0
      MsgStop("La cuenta futura debe de ser distinta de cero","error")
      RETURN
   ENDIF
   IF VAL(W_CueCamb.T_CodFut.Value)=VAL(W_CueCamb.T_CodAct.Value)
      MsgStop("La cuenta futura debe de ser distinta de la cuenta actual","error")
      RETURN
   ENDIF

   AbrirDBF("CUENTAS",,,,,1)

   SEEK VAL(W_CueCamb.T_CodFut.Value)
   IF .NOT. EOF()
      IF MSGYESNO("La cuenta futura ya esta creada"+HB_OsNewLine()+ ;
                  W_CueCamb.T_CodFut.Value+" "+W_CueCamb.T_NomFut.Value+HB_OsNewLine()+ ;
                  "¿Desea continuar?","Atencion")=.F.
         RETURN
      ENDIF
   ENDIF

   IF MSGYESNO("¿Desea cambiar la cuenta?"+HB_OsNewLine()+ ;
               "Actual "+W_CueCamb.T_CodAct.Value+" "+W_CueCamb.T_NomAct.Value+HB_OsNewLine()+ ;
               "Futura "+W_CueCamb.T_CodFut.Value+" "+W_CueCamb.T_NomFut.Value)=.F.
      RETURN
   ENDIF


ANY2:=EJERANY

DO WHILE .T.
   IF W_CueCamb.SiEjercicios.Value=.F. .AND. ANY2<EJERANY
      EXIT
   ENDIF
   RUTA2:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,ANY2,"SUICONTA")
   RUTA2:=RUTA2[1]
   IF LEN(RTRIM(RUTA2))=0
      EXIT
   ENDIF

   IF ANY2<EJERANY
      IF MSGYESNO("Desea cambiar la cuenta en el ejercicio "+LTRIM(STR(ANY2))+"?"+HB_OsNewLine()+ ;
                   RUTA2,"Atencion")=.F.
         EXIT
      ENDIF
   ENDIF

   ANY2--

	FOR N=1 TO 7
	   DO CASE
	   CASE N=1
			IF W_CueCamb.SiFicCta.Value=.F.
				LOOP
			ENDIF
	      AbrirDBF("CUENTAS",,,,RUTA2,2)
	      CAMPO:="CODCTA"
	   CASE N=2
	      AbrirDBF("APUNTES",,,,RUTA2)
	      CAMPO:="CODCTA"
	   CASE N=3
	      AbrirDBF("CIERRE",,,,RUTA2)
	      CAMPO:="CODCTA"
	   CASE N=4
	      AbrirDBF("FAC92",,,,RUTA2)
	      CAMPO:="CODCTA"
	   CASE N=5
	      AbrirDBF("FACREB",,,,RUTA2)
	      CAMPO:="CODIGO"
	   CASE N=6
	      AbrirDBF("FACVTO",,,,RUTA2)
	      CAMPO:="CODIGO"
	   CASE N=7
	      AbrirDBF("REMESA",,,,RUTA2)
	      CAMPO:="CODCTA"
	   ENDCASE
	
	   SetProperty("W_CueCamb","Progres"+LTRIM(STR(N)),"RangeMax",LASTREC())
	   SetProperty("W_CueCamb","Progres"+LTRIM(STR(N)),"Value",0)
	   SET DELETE OFF
	   GO TOP
	   DO WHILE .NOT. EOF()
	      SetProperty("W_CueCamb","Progres"+LTRIM(STR(N)),"Value", GetProperty("W_CueCamb","Progres"+LTRIM(STR(N)),"Value")+1 )
	      DO EVENTS
	      IF DELETED()=.T.
	         SKIP
	         LOOP
	      ENDIF
	      IF &CAMPO=VAL(W_CueCamb.T_CodAct.Value)
	         SetProperty("W_CueCamb","T_ModFic"+LTRIM(STR(N)),"Value", GetProperty("W_CueCamb","T_ModFic"+LTRIM(STR(N)),"Value")+1 )
	         IF RLOCK()
	            REPLACE &CAMPO WITH VAL(W_CueCamb.T_CodFut.Value)
	            DBCOMMIT()
	            DBUNLOCK()
	         ENDIF
	      ENDIF
	      SKIP
	   ENDDO
	   SET DELETE ON
	
	NEXT

ENDDO

MsgInfo("El cambio de codigo de la cuenta ha finalizado correctamente")



