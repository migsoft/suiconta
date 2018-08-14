#include "minigui.ch"
#include "winprint.ch"

procedure balmay2()
   TituloImp:="Libro mayor de subcuentas"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 420 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_CodCta1 VALUE 'Desde el codigo' AUTOSIZE TRANSPARENT
      @ 10,110 TEXTBOX T_CodCta1 WIDTH 100 VALUE "" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta1.Value:=PCODCTA3(W_Imp1.T_CodCta1.Value)
      @ 10,225 BUTTONEX Bt_BuscarCue1 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta1.Value),"W_Imp1","T_CodCta1") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @ 45,10 LABEL L_CodCta2 VALUE 'Hasta el codigo' AUTOSIZE TRANSPARENT
      @ 40,110 TEXTBOX T_CodCta2 WIDTH 100 VALUE "99999999" MAXLENGTH 8 ;
               ON LOSTFOCUS W_Imp1.T_CodCta2.Value:=PCODCTA3(W_Imp1.T_CodCta2.Value)
      @ 40,225 BUTTONEX Bt_BuscarCue2 ;
         CAPTION 'Buscar' ICON icobus('buscar') ;
         ACTION br_cue1(VAL(W_Imp1.T_CodCta2.Value),"W_Imp1","T_CodCta2") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @ 75,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 70,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
      @105,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @100,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

      @135,10 LABEL L_Cierre VALUE 'Incluir asientos de cierre' AUTOSIZE TRANSPARENT
      @130,150 COMBOBOX C_Cierre WIDTH 100 ITEMS {"Ninguno","Ajustes","Regulacion","Cierre","Todos"} VALUE 1

      @165,10 LABEL L_FecLis VALUE 'Fecha listado' AUTOSIZE TRANSPARENT
      @160,110 DATEPICKER D_FecLis WIDTH 100 VALUE DATE()

      @195,10 LABEL L_PagIni VALUE 'Numero pagina' AUTOSIZE TRANSPARENT
      @190,110 TEXTBOX T_PagIni WIDTH 100 HEIGHT 25 VALUE 1 ;
              NUMERIC INPUTMASK '99,999,999,999' FORMAT 'E' RIGHTALIGN

      @225,10 LABEL L_Progres VALUE 'Cuenta' AUTOSIZE TRANSPARENT
      @220,110 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 280 HEIGHT 20 SMOOTH

LINW:=260
COLW:=0
draw rectangle in window W_Imp1 at LINW,COLW+10 to LINW+2,COLW+390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @LINW+15,COLW+10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @LINW+10,COLW+100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @LINW+45,COLW+220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @LINW+40,COLW+280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @LINW+40,COLW+10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @LINW+70,COLW+10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @LINW+100,COLW+10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION ( balmay2i() , W_Imp1.release )

      @LINW+100,COLW+110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION balmay2i()
   local oprint

STORE 0 TO PAG,PAG1,LIN

AbrirDBF("CUENTAS",,,,,1)
SEEK VAL(W_Imp1.T_CodCta1.Value)
IF EOF()
   GO TOP
   DO WHILE CODCTA<VAL(W_Imp1.T_CodCta1.Value) .AND. .NOT. EOF()
      SKIP
   ENDDO
ENDIF

IF EOF()
	MsgStop("No hay datos para imprimir")
	RETURN
ENDIF

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloqImp(W_Imp1.Title))
   oprint:setpreviewsize(2)  // tamaño del preview
   oprint:beginpage()

   W_Imp1.P_Progres.RangeMax:=LASTREC()
   W_Imp1.P_Progres.Value:=0
CODCTA2:=-1
DO WHILE CODCTA<=VAL(W_Imp1.T_CodCta2.Value) .AND. .NOT. EOF()
   DO EVENTS

   IF LIN>=250 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(280,105,"Sigue en la hoja: "+LTRIM(STR(PAG2+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1
		PAG2:=PAG+(W_Imp1.T_PagIni.value-1)

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG2)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(W_Imp1.D_FecLis.Value,10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,105,TituloqImp(W_Imp1.Title),"times new roman",18,.F.,,"C",)

      LIN:=40
		IF PAG=1
	      oprint:printdata(LIN  ,20,"Desde fecha "+DIA(W_Imp1.D_Fec1.Value,10),"times new roman",12,.F.,,"L",)
	      oprint:printdata(LIN+5,20,"Hasta fecha "+DIA(W_Imp1.D_Fec2.Value,10),"times new roman",12,.F.,,"L",)
	      oprint:printdata(LIN  ,90,"Desde cuenta "+LTRIM(W_Imp1.T_CodCta1.Value),"times new roman",12,.F.,,"L",)
	      oprint:printdata(LIN+5,90,"Hasta cuenta "+LTRIM(W_Imp1.T_CodCta2.Value),"times new roman",12,.F.,,"L",)
			LIN:=LIN+5
		ENDIF
	ENDIF

	IF CODCTA<>CODCTA2
		AbrirDBF("CUENTAS",,,,,1)
		CODCTA2:=CODCTA
		NOMCTA2:=NOMCTA
		DEB2:=0
		HAB2:=0

		AbrirDBF("APUNTES",,,,,2)
		SEEK STR(CODCTA2,8)
		DO WHILE CODCTA=CODCTA2 .AND. FECHA<W_Imp1.D_Fec1.Value .AND. .NOT. EOF()
			DEB2:=DEB2+DEBE
			HAB2:=HAB2+HABER
			SKIP
		ENDDO

		IF DEB2<>0 .OR. HAB2<>0 .OR. CODCTA=CODCTA2
	      LIN:=LIN+10
	      oprint:printdata(LIN, 50,CODCTA2,"times new roman",12,.F.,,"R",)
	      oprint:printdata(LIN, 52,NOMCTA2,"times new roman",12,.F.,,"L",)
	      LIN:=LIN+5
	      oprint:printdata(LIN, 30,"Asiento","times new roman",10,.F.,,"R",)
	      oprint:printdata(LIN, 50,"Fecha","times new roman",10,.F.,,"R",)
	      oprint:printdata(LIN, 52,"Descripcion","times new roman",10,.F.,,"L",)
	      oprint:printdata(LIN,130,"Debe","times new roman",10,.F.,,"R",)
	      oprint:printdata(LIN,150,"Haber","times new roman",10,.F.,,"R",)
	      oprint:printdata(LIN,170,"Saldo","times new roman",10,.F.,,"R",)
	      oprint:printline(LIN+4,20,LIN+4,170,,0.5)
	      LIN:=LIN+5
			IF DEB2>HAB2
				DEB2:=DEB2-HAB2
				HAB2:=0
			ELSE
				DEB2:=0
				HAB2:=HAB2-DEB2
			ENDIF
			SALDO2:=DEB2-HAB2
			IF SALDO<>0
		      oprint:printdata(LIN, 52,"Saldo inicial","times new roman",10,.F.,,"L",)
		      oprint:printdata(LIN,130,MIL(DEB2,14,2),"times new roman",10,.F.,,"R",)
		      oprint:printdata(LIN,150,MIL(HAB2,14,2),"times new roman",10,.F.,,"R",)
		      oprint:printdata(LIN,170,MIL(SALDO2,14,2),"times new roman",10,.F.,,"R",)
		      LIN:=LIN+5
			ENDIF
		ENDIF
	ENDIF

	AbrirDBF("APUNTES",,,,,2)
	DO WHILE CODCTA=CODCTA2
	   IF LIN>=260
			EXIT
		ENDIF
		IF FECHA<=W_Imp1.D_Fec2.Value
	      oprint:printdata(LIN, 30,NASI,"times new roman",10,.F.,,"R",)
	      oprint:printdata(LIN, 50,DIA(FECHA,10),"times new roman",10,.F.,,"R",)
		   oprint:printdata(LIN, 52,NOMAPU,"times new roman",10,.F.,,"L",)
		   oprint:printdata(LIN,130,MIL(DEBE,14,2),"times new roman",10,.F.,,"R",)
		   oprint:printdata(LIN,150,MIL(HABER,14,2),"times new roman",10,.F.,,"R",)
			SALDO2:=SALDO2+DEBE-HABER
		   oprint:printdata(LIN,170,MIL(SALDO2,14,2),"times new roman",10,.F.,,"R",)
		   LIN:=LIN+5
		ENDIF
		SKIP
	ENDDO
	IF CODCTA=CODCTA2 .AND. LIN>=260
		AbrirDBF("CUENTAS",,,,,1)
		LOOP
	ENDIF

	IF W_Imp1.C_Cierre.Value>1
		AbrirDBF("CIERRE",,,,,2)
		SEEK STR(CODCTA2,8)
		DO WHILE CODCTA=CODCTA2
		   IF LIN>=260
				EXIT
			ENDIF
			IF W_Imp1.C_Cierre.Value=2 .AND. NASI<999997 .OR. ;
				W_Imp1.C_Cierre.Value=3 .AND. NASI=999997 .OR. ;
				W_Imp1.C_Cierre.Value=4 .AND. NASI=999998 .OR. ;
				W_Imp1.C_Cierre.Value=5

		      oprint:printdata(LIN, 30,NASI,"times new roman",10,.F.,,"R",)
		      oprint:printdata(LIN, 50,DIA(FECHA,10),"times new roman",10,.F.,,"R",)
			   oprint:printdata(LIN, 52,NOMAPU,"times new roman",10,.F.,,"L",)
			   oprint:printdata(LIN,130,MIL(DEBE,14,2),"times new roman",10,.F.,,"R",)
			   oprint:printdata(LIN,150,MIL(HABER,14,2),"times new roman",10,.F.,,"R",)
				SALDO2:=SALDO2+DEBE-HABER
			   oprint:printdata(LIN,170,MIL(SALDO2,14,2),"times new roman",10,.F.,,"R",)
			   LIN:=LIN+5
			ENDIF
			SKIP
		ENDDO
		IF CODCTA=CODCTA2 .AND. LIN>=260
			AbrirDBF("CUENTAS",,,,,1)
			LOOP
		ENDIF
	ENDIF

	AbrirDBF("CUENTAS",,,,,1)
	SKIP
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
	W_Imp1.L_Progres.value:=STR(CODCTA)

ENDDO

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

Return Nil
