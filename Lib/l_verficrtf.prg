#include "minigui.ch"

procedure VerFicRtf(NomFicRtf,AnchoFicRtf,AltoFicRtf)

DEFAULT NomFicRtf   to "Abrir fichero"
DEFAULT AnchoFicRtf to 1024
DEFAULT AltoFicRtf  to 768-30

   DEFINE WINDOW W_VerFicRtf ;
      AT 10,10 ;
      WIDTH AnchoFicRtf ;
      HEIGHT AltoFicRtf ;
      TITLE NomFicRtf ;
      CHILD ;
      ON SIZE VerFicRtf_Tam("SIZE") ;
		ON MAXIMIZE VerFicRtf_Tam("MAXIMIZE") ;
		ON RESTORE VerFicRtf_Tam("RESTAURAR")

      DEFINE MAIN MENU

      POPUP 'Fichero'
         MENUITEM 'Abrir'          ACTION VerFicRtf_Fic()
         MENUITEM 'Volver a abrir' ACTION VerFicRtf_Leer()
         MENUITEM 'Editor externo' ACTION VerFicRtf_Editar()
         MENUITEM 'Guardar'        ACTION VerFicRtf_Guardar()
         MENUITEM 'Guardar como'   ACTION VerFicRtf_Guardar("GUARDARCOMO")
         SEPARATOR
         MENUITEM 'Salir'    ACTION W_VerFicRtf.release
      END POPUP
      POPUP 'Fuente'
         MENUITEM 'Seleccionar'    ACTION VerFicRtf_Fuente()
      END POPUP

      END MENU

      @10,10 BUTTONEX B_AbrirFic CAPTION '' ICON icobus('buscar') WIDTH 30 HEIGHT 25 TOOLTIP 'Abrir fichero' ;
               ACTION VerFicRtf_Fic()

      @10,50 BUTTONEX B_EditarFic CAPTION '' ICON icobus('modificar') WIDTH 30 HEIGHT 25 TOOLTIP 'Editor externo' ;
               ACTION VerFicRtf_Editar()

      @10,90 BUTTONEX B_GuardarFic CAPTION '' ICON icobus('guardar') WIDTH 30 HEIGHT 25 TOOLTIP 'Guardar' ;
               ACTION VerFicRtf_Guardar()

      @10,130 BUTTONEX B_CancelarFic CAPTION '' ICON icobus('cancelar') WIDTH 30 HEIGHT 25 TOOLTIP 'Cancelar' ;
               ACTION VerFicRtf_Leer()

      @10,170 BUTTONEX B_Fuente CAPTION '' ICON icobus('fuente') WIDTH 30 HEIGHT 25 TOOLTIP 'Fuente' ;
			      ACTION VerFicRtf_Fuente()

      @10,210 SPINNER FontSize RANGE 1,100 VALUE 10 WIDTH 90 HEIGHT 25 ;
               ON CHANGE W_VerFicRtf.FicRtf.FontSize:=W_VerFicRtf.FontSize.Value

      @10,310 BUTTONEX B_Salir CAPTION '' ICON icobus('salir') WIDTH 30 HEIGHT 25 TOOLTIP 'Salir' ;
               ACTION W_VerFicRtf.release


      @ 40,10 RICHEDITBOX FicRtf WIDTH W_VerFicRtf.Width-30 HEIGHT W_VerFicRtf.Height-110 VALUE memoRead(NomFicRtf) FONT NomFuente('Courier New') SIZE 10 BOLD

      DEFINE CONTEXT MENU CONTROL FicRtf OF W_VerFicRtf
         ITEM "Copiar seleccion"                     ACTION VerFicRtf_Menu("SELECCION")
         ITEM "Copiar toda la linea al portapapeles" ACTION VerFicRtf_Menu("LINEA")
      END MENU


      END WINDOW
      VentanaCentrar("W_VerFicRtf","Ventana1")
      ACTIVATE WINDOW W_VerFicRtf

Return Nil



STATIC FUNCTION VerFicRtf_Fuente()
	aFontD:=GetFont(W_VerFicRtf.FicRtf.FontName, ;
						 W_VerFicRtf.FicRtf.FontSize, ;
						 W_VerFicRtf.FicRtf.FontBold, ;
						 W_VerFicRtf.FicRtf.FontItalic, ;
						 W_VerFicRtf.FicRtf.FontColor, ;
						 W_VerFicRtf.FicRtf.FontUnderline, ;
						 W_VerFicRtf.FicRtf.FontStrikeout)
	IF LEN(aFontD[1])>0
		W_VerFicRtf.FicRtf.FontName:=aFontD[1]
		W_VerFicRtf.FicRtf.FontSize:=aFontD[2]
		W_VerFicRtf.FicRtf.FontBold:=aFontD[3]
		W_VerFicRtf.FicRtf.FontItalic:=aFontD[4]
		W_VerFicRtf.FicRtf.FontColor:=aFontD[5]
		W_VerFicRtf.FicRtf.FontUnderline:=aFontD[6]
		W_VerFicRtf.FicRtf.FontStrikeout:=aFontD[7]
		W_VerFicRtf.FontSize.Value:=W_VerFicRtf.FicRtf.FontSize
		W_VerFicRtf.FicRtf.SetFocus
	ENDIF




STATIC FUNCTION VerFicRtf_Tam(LLAMADA)
DO CASE
CASE LLAMADA="SIZE"
   IF (W_VerFicRtf.Width)%20<>0 .AND. (W_VerFicRtf.Height)%15<>0
      RETURN
   ENDIF
CASE LLAMADA="RESTAURAR"
/*
   W_VerFicRtf.Width :=1024
   W_VerFicRtf.Height:=768-30
   VentanaCentrar("W_VerFicRtf","Ventana1")
*/
CASE LLAMADA="MAXIMIZE"
/*
   W_VerFicRtf.Col   :=0
   W_VerFicRtf.Row   :=0
   W_VerFicRtf.Width :=GetDesktopWidth()
   W_VerFicRtf.Height:=GetDesktopHeight()
*/
ENDCASE

IF IsControlDefined(FicRtf,W_VerFicRtf)=.F.
	RETURN
ENDIF

W_VerFicRtf.FicRtf.Width :=(W_VerFicRtf.Width)-30
W_VerFicRtf.FicRtf.Height:=(W_VerFicRtf.Height)-110

DO EVENTS
DO CASE
CASE W_VerFicRtf.Width<=800
   W_VerFicRtf.FicRtf.FontSize:=8
CASE W_VerFicRtf.Width<=900
   W_VerFicRtf.FicRtf.FontSize:=9
CASE W_VerFicRtf.Width<=1000
   W_VerFicRtf.FicRtf.FontSize:=10
CASE W_VerFicRtf.Width<=1100
   W_VerFicRtf.FicRtf.FontSize:=11
CASE W_VerFicRtf.Width<=1200
   W_VerFicRtf.FicRtf.FontSize:=12
OTHERWISE
   W_VerFicRtf.FicRtf.FontSize:=14
ENDCASE
W_VerFicRtf.FontSize.Value:=W_VerFicRtf.FicRtf.FontSize
W_VerFicRtf.FicRtf.SetFocus



STATIC FUNCTION VerFicRtf_Fic()
   W_VerFicRtf.Title:=Getfile(  {{'Ficheros de texto (*.txt;*.rtf)','*.txt;*.rtf'}, ;
											{'Ficheros programacion (*.ini;*.bat;*.prg)','*.ini;*.bat;*.prg'}, ;
											{'Todos los ficheros','*.*'}} ,'Buscar fichero',W_VerFicRtf.Title, .f. , .t. )
	VerFicRtf_Leer()


STATIC FUNCTION VerFicRtf_Leer()
   W_VerFicRtf.FicRtf.Value:=memoRead(W_VerFicRtf.Title)
   W_VerFicRtf.FicRtf.SetFocus




STATIC FUNCTION VerFicRtf_Editar()
   NomFichero2:='"'+W_VerFicRtf.Title+'"'
   rutaArcProg2:=GetProgramFilesFolder()
IF RIGHT(UPPER(W_VerFicRtf.Title),4)=".RTF"
   IF ShellExecute(0, "open", "soffice.exe", NomFichero2 , , 1)<=32
      IF ShellExecute(0, "open", "WinWord.exe", NomFichero2 , , 1)<=32
         IF ShellExecute(0, "open", NomFichero2 , , , 1)<=32
            MSGINFO("No se ha localizado el programa asociado a la extension"+HB_OsNewLine()+HB_OsNewLine()+NomFichero2)
         ENDIF
      ENDIF
   ENDIF
ELSE
   IF ShellExecute(0, "open", rutaArcProg2+"\Crimson Editor\cedt.exe", NomFichero2 , , 1)<=32
      IF ShellExecute(0, "open", rutaArcProg2+"\Emerald Editor Community\Crimson Editor 3.72\cedt.exe", NomFichero2 , , 1)<=32
         IF ShellExecute(0, "open", "NOTEPAD.EXE" , NomFichero2 , , 1)<=32
            IF ShellExecute(0, "open", NomFichero2 , , , 1)<=32
               MSGINFO("No se ha localizado el programa asociado a la extension"+HB_OsNewLine()+HB_OsNewLine()+NomFichero2)
            ENDIF
         ENDIF
      ENDIF
   ENDIF
ENDIF



STATIC FUNCTION VerFicRtf_Guardar(LLAMADA)
	DEFAULT LLAMADA TO ""

	IF LLAMADA==UPPER("GUARDARCOMO")
	   NomFichero2:=Getfile( {{'Ficheros de texto','*.*'}} ,'Buscar fichero',W_VerFicRtf.Title, .f. , .t. )
		IF LEN(NomFichero2)=0
			MsgStop("Proceso cancelado")
			RETURN
		ENDIF
		IF FILE(NomFichero2)=.T.
			IF MsgYesNo(NomFichero2+HB_OsNewLine()+"¿Desea reemplazar el fichero?","Guardar")=.F.
				RETURN
			ENDIF
		ENDIF
		W_VerFicRtf.Title:=NomFichero2
	ELSE
		IF LEN(W_VerFicRtf.Title)=0
			MsgStop("No se ha abierto ningun fichero")
			RETURN
		ENDIF
		IF MsgYesNo(W_VerFicRtf.Title+HB_OsNewLine()+"¿Desea guardar el fichero?","Guardar")=.F.
			RETURN
		ENDIF
	ENDIF

#ifdef __OOHG__
	MsgStop("no se graba en oohg")
#else
	IF Upper(AllTrim(SubStr(W_VerFicRtf.Title, Rat('.',W_VerFicRtf.Title) + 1))) == 'RTF'
		MemoWrit( W_VerFicRtf.Title, W_VerFicRtf.FicRtf.RichValue )
	ELSE
		MemoWrit( W_VerFicRtf.Title, W_VerFicRtf.FicRtf.Value )
   ENDIF
	MsgInfo(W_VerFicRtf.Title+HB_OsNewLine()+"Fichero guardado")
#endif



STATIC FUNCTION VerFicRtf_Menu(LLAMADA)
LOCAL INILIN:=0,FINLIN:=0,TOTLIN:=0,Texto1:=""
   TDATOS := W_VerFicRtf.FicRtf.Value
   FOR N=1 TO LEN(TDATOS)
      IF SUBSTR(TDATOS,N,1)==CHR(13)
         IF N<W_VerFicRtf.FicRtf.CaretPos+TOTLIN+1
            INILIN:=N
            TOTLIN++
         ELSE
            FINLIN:=N
            EXIT
         ENDIF
      ENDIF
   NEXT

DO CASE
CASE LLAMADA="SELECCION"
   nCONTROL:=GetControlHandle("FicRtf","W_VerFicRtf") //GetFormHandle("form1")
   INISEL:=LoWord ( SendMessage( nCONTROL , 176 , 0 , 0 ) )
   FINSEL:=HiWord ( SendMessage( nCONTROL , 176 , 0 , 0 ) )
   Texto1:=SUBSTR(TDATOS,INISEL+TOTLIN+1,FINSEL-INISEL)
CASE LLAMADA="LINEA"
   Texto1:=SUBSTR(W_VerFicRtf.FicRtf.Value,INILIN+2,FINLIN-INILIN-2)
ENDCASE
CopyToClipboard(Texto1)
MsgInfo(Texto1,"Portapapeles")
