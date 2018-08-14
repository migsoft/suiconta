#include "minigui.ch"
** RUTAEB:=RUTAPROGRAMA

Function br_recbus(RUTAEB,NEMP2,NFAC2,NOMPROV2,POSICION)
   LOCAL Color_rec1 := { || IIF( PEND<>0 , RGB(255,200,200), RGB(255,255,255)) }
	DEFAULT NOMPROV2 TO ""
	DEFAULT POSICION TO ""

   NEMP2:=IF(NEMP2=0,NUMEMP,NEMP2)

   aNFAC:={0,"",""}

   IREMPRESA:=1
   aRUTAEMP1:={}
   aRUTAEMP2:={}
   aRUTAEMP3:={}
   aRUTAEMP:=aRUTAEMP(RUTAEB,"SUICONTA")
   IF LEN(aRUTAEMP)>=1
      FOR N=1 TO LEN(aRUTAEMP)
         AADD(aRUTAEMP1,aRUTAEMP[N,1])
         AADD(aRUTAEMP2,aRUTAEMP[N,2])
         AADD(aRUTAEMP3,LTRIM(aRUTAEMP[N,4]))
         IF VAL(aRUTAEMP[N,4])=NEMP2
            IREMPRESA=N
         ENDIF
      NEXT
   ENDIF
   
   AutoMsgInfo(aRUTAEMP2[IREMPRESA])

   AbrirDBF("FACREB",,,,RUTAEB+"\"+aRUTAEMP2[IREMPRESA])
   IF NFAC2<>0
      DBSETORDER(1)
      SEEK NFAC2
      CODRECBR:=RECNO()
   ELSE
      IF LEN(RTRIM(NOMPROV2))<>0
         DBSETORDER(2)
         SET SOFTSEEK ON
         SEEK UPPER(NOMPROV2)
         SET SOFTSEEK OFF
			IF UPPER(POSICION)="ULTIMA" .AND. UPPER(FACREB->NOMCTA)=UPPER(NOMPROV2)
				DO WHILE UPPER(FACREB->NOMCTA)=UPPER(NOMPROV2) .AND. .NOT. EOF()
					SKIP
				ENDDO
				SKIP -1
			ENDIF
         CODRECBR:=RECNO()
      ELSE
         CODRECBR:=1
      ENDIF
   ENDIF
   DBSETORDER(2)

   DEFINE WINDOW Winfacbus1 ;
      AT 50,0     ;
      WIDTH 700 HEIGHT 470 ;
      TITLE "Facturas recibidas" ;
      MODAL ;
      NOSIZE BACKCOLOR MiColor("VERDECLARO")

      @0,0 TEXTBOX RutaEB WIDTH 250 VALUE RUTAEB INVISIBLE

      @ 10,10 BROWSE br_rec1 ;
      WIDTH 676 HEIGHT 350 ;
      HEADERS {'Codigo','Fecha','Nombre','Factura','Importe','Pendiente','Asiento'} ;
      WIDTHS { 55,80,175,70,100,100,75 } ;
      WORKAREA FACREB ;
      FIELDS { 'FACREB->NREG','FACREB->FREG','FACREB->NOMCTA','FACREB->REF','LTRIM(MIL(FACREB->TFAC,15,2))','LTRIM(MIL(FACREB->PEND,15,2))','ASIENTO' } ;
      DYNAMICBACKCOLOR { Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1,Color_rec1 } ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE CODRECBR ;
      ON DBLCLICK br_recbusFin() ;
      ON HEADCLICK { {|| AbrirDBF("FACREB",,,,RUTAEB+"\"+Winfacbus1.C_RutaEmp.DisplayValue),DBSETORDER(1), Winfacbus1.br_rec1.Refresh} , , ;
                     {|| AbrirDBF("FACREB",,,,RUTAEB+"\"+Winfacbus1.C_RutaEmp.DisplayValue),DBSETORDER(2), Winfacbus1.br_rec1.Refresh} }

      DEFINE CONTEXT MENU CONTROL br_rec1 OF Winfacbus1
         ITEM "Copiar registro"            ACTION Winfacbus1_CopReg("REG1")
         ITEM "Copiar registros proveedor" ACTION Winfacbus1_CopReg("REGPROV")
      END MENU

      @370,010 BUTTONEX Bt_Localizar CAPTION 'Localizar' ICON icobus('buscar') ;
               ACTION br_recbusLoc() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP
      @370,110 TEXTBOX T_BusNom WIDTH 250 VALUE '' MAXLENGTH 30 ;
               ON CHANGE br_recbusfac("BUSCAR")

      @370,370 RADIOGROUP R_TipLoc OPTIONS { 'Nombre','Factura' } ;
              VALUE 1 WIDTH 80 HORIZONTAL

      @405,010 LABEL L_Empresa VALUE 'Empresa:' AUTOSIZE TRANSPARENT
      @400,110 COMBOBOX C_Empresa WIDTH 300 ITEMS aRUTAEMP1 VALUE IREMPRESA NOTABSTOP ;
              ON CHANGE (Winfacbus1.C_RutaEmp.Value:=Winfacbus1.C_Empresa.Value , ;
                         Winfacbus1.C_CodEmp.Value :=Winfacbus1.C_Empresa.Value , br_recbusfac("RUTA") )
      @400,110 COMBOBOX C_RutaEmp WIDTH 270 ITEMS aRUTAEMP2 VALUE IREMPRESA INVISIBLE
      @400,110 COMBOBOX C_CodEmp  WIDTH  50 ITEMS aRUTAEMP3 VALUE IREMPRESA INVISIBLE

      @400,600 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
         ACTION Winfacbus1.Release ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

   Winfacbus1.T_BusNom.SetFocus

   END WINDOW
   VentanaCentrar("Winfacbus1","Ventana1")
   ACTIVATE WINDOW Winfacbus1

   RETURN(aNFAC)

STATIC FUNCTION br_recbusFin()
aNFAC:={Winfacbus1.br_rec1.Value ,;
        Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue ,;
        VAL(Winfacbus1.C_CodEmp.DisplayValue) }
   Winfacbus1.Release

STATIC FUNCTION br_recbusfac(LLAMADA)
   AbrirDBF("FACREB",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
   IF LLAMADA="RUTA"
      Winfacbus1.br_rec1.Refresh
      RETURN
   ENDIF

   SET SOFTSEEK ON
   IF VAL(Winfacbus1.T_BusNom.Value)<>0
      DBSETORDER(1)
      SEEK VAL(Winfacbus1.T_BusNom.Value)
   ELSE
      DBSETORDER(2)
      SEEK UPPER(Winfacbus1.T_BusNom.Value)
   ENDIF
   SET SOFTSEEK OFF

   Winfacbus1.br_rec1.Value:=RECNO()
   Winfacbus1.br_rec1.Refresh



STATIC FUNCTION br_recbusLoc()
   NOMBUSCAR:=RTRIM(Winfacbus1.T_BusNom.Value)
   IF LEN(NOMBUSCAR)=0
      RETURN
   ENDIF
   AbrirDBF("FACREB",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
   IF Winfacbus1.br_rec1.Value=0
      GO TOP
      Winfacbus1.br_rec1.Value:=RECNO()
   ENDIF
   GO Winfacbus1.br_rec1.Value
   DO WHILE .T.
      DO EVENTS
      SKIP
      IF EOF()
         GO TOP
      ENDIF
      DO CASE
      CASE Winfacbus1.R_TipLoc.Value=1
         IF AT(UPPER(NOMBUSCAR),UPPER(NOMCTA))<>0
            Winfacbus1.br_rec1.Value:=recno()
            Winfacbus1.br_rec1.Refresh
            EXIT
         ENDIF
      CASE Winfacbus1.R_TipLoc.Value=2
         IF AT(UPPER(NOMBUSCAR),UPPER(REF))<>0
            Winfacbus1.br_rec1.Value:=recno()
            Winfacbus1.br_rec1.Refresh
            EXIT
         ENDIF
      ENDCASE
      IF RECNO()=Winfacbus1.br_rec1.Value
         MSGSTOP("No se ha localizado ningun registro")
         EXIT
      ENDIF
   ENDDO


STATIC FUNCTION Winfacbus1_CopReg(LLAMADA)
   TEXTO2:=""
   AbrirDBF("FACREB",,,,Winfacbus1.RutaEB.Value+"\"+Winfacbus1.C_RutaEmp.DisplayValue)
   GO Winfacbus1.br_rec1.Value
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

