#include "minigui.ch"

Function CrearAccesoDir(nomprog)

   LOCAL RUTALOGOPROG:=IF(RUTALOGOPROG=NIL,'',RUTALOGOPROG)
   nomprog:=UPPER( IF(nomprog=NIL,programa(),nomprog) )

   FicProg:=Fichero_Programa(nomprog)
   FicProg:=LEFT(FicProg,AT(".EXE",UPPER(FicProg))-1)+"*.EXE"
   aFicProg1:=DIRECTORY(RUTAPROGRAMA+"\"+FicProg)
   ASORT( aFicProg1,,, {| x, y | x[1] < y[1] } )

   aFicProg2:={}
   FOR N=1 TO LEN(aFicProg1)
      AADD(aFicProg2,aFicProg1[N,1]+" "+DIA(aFicProg1[N,3],10))
   NEXT

   DEFINE WINDOW W_AccDir ;
      AT 10,10 ;
      WIDTH 600 HEIGHT 300 ;
      TITLE 'Crear acceso directo del programa' ;
      MODAL   ;
      NOSIZE BACKCOLOR {255,225,150} //Naranja claro

      LIN:=15
      @LIN,10 LABEL L_A1 ;
              VALUE 'Programa:' ;
              WIDTH 140 HEIGHT 25 TRANSPARENT
      @LIN,150 LABEL L_A2 ;
              VALUE Version_Programa(nomprog) ;
              WIDTH 450 HEIGHT 25 TRANSPARENT FONTCOLOR {255,0,0} //rojo
      LIN:=LIN+15
      @LIN,150 LABEL L_A3 ;
              VALUE Fichero_Programa("RUTAFICHEROEXE") ;
              WIDTH 450 HEIGHT 25 TRANSPARENT FONTCOLOR {153,0,153} //Violeta Oscuro

      LIN:=LIN+30
      @LIN,10 LABEL L_B1 ;
              VALUE "Ruta escritorio:" ;
              WIDTH 140 HEIGHT 25 TRANSPARENT
      @LIN,150 LABEL L_B2 ;
              VALUE GetDesktopFolder() ;
              WIDTH 435 HEIGHT 15 TRANSPARENT FONTCOLOR {0,153,51} //Verde Hierba

      LIN:=LIN+30
      @LIN+5,10 LABEL L_C1 ;
              VALUE "Versiones de programa:" ;
              WIDTH 140 HEIGHT 25 TRANSPARENT
      @LIN,150 COMBOBOX C_Fic WIDTH 200 ITEMS aFicProg2 VALUE 1


      @230, 10 BUTTONEX B_Imp CAPTION 'Crear' WIDTH 90 HEIGHT 25 ;
               ACTION CrearAccesoDir2()

      @230,110 BUTTON Bt_Salir CAPTION 'Salir' WIDTH 90 HEIGHT 25 ;
               ACTION W_AccDir.Release TOOLTIP 'Salir'


      END WINDOW
      VentanaCentrar("W_AccDir","Ventana1")
      CENTER WINDOW W_AccDir
      ACTIVATE WINDOW W_AccDir

RETURN Nil

*
STATIC Function CrearAccesoDir2(nomprog)
   LOCAL olnk

   FicProg:=W_AccDir.C_Fic.Item(W_AccDir.C_Fic.Value)
   FicProg:=LEFT(FicProg,AT(".EXE",UPPER(FicProg))+3)
   DATOSFIC1:=RUTAPROGRAMA+"\"+FicProg

   olnk := TOleAuto():New( "WScript.Shell" )
   strDesktop :=olnk:SpecialFolders():Item("desktop")

   strPathToAccesoDir := olnk:ExpandEnvironmentStrings(DATOSFIC1)

   oShellLink := olnk:CreateShortcut(strDesktop + "\"+NOMPROGRAMA+".lnk")
   oShellLink:TargetPath := strPathToAccesoDir

   *-------------Esto no es indispensable-----------
   oShellLink:WindowStyle := 1
*   oShellLink:Hotkey := "CTRL+SHIFT+F"
   oShellLink:IconLocation := DATOSFIC1+", 0"
   oShellLink:Description := Version_Programa(nomprog)
   oShellLink:WorkingDirectory := RUTAPROGRAMA
   *-----------------------------------------------------------
   oShellLink:Save()

   MSGINFO("Se ha creado el acceso directo")
   W_AccDir.Release

RETURN Nil
