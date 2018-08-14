#include "minigui.ch"

FUNCTION VentanaCentrar(NonVentanaCen,NonVentanaPrin,AliVentanaCen)
AliVentanaCen:=IF(AliVentanaCen=NIL,"CENTRAR",UPPER(AliVentanaCen))
IF IsWindowDefined(&NonVentanaCen)=.T.
   IF IsWindowDefined(&NonVentanaPrin)=.T.
      AnchoVentanaPrin:=GetProperty(NonVentanaPrin,"Width")
      AltoVentanaPrin :=GetProperty(NonVentanaPrin,"Height")
      ColVentanaPrin  :=GetProperty(NonVentanaPrin,"Col")
      LinVentanaPrin  :=GetProperty(NonVentanaPrin,"Row")
   ELSE
      AnchoVentanaPrin:=GetDesktopWidth()
      AltoVentanaPrin :=GetDesktopHeight()
      ColVentanaPrin  :=0
      LinVentanaPrin  :=0
   ENDIF

   DO CASE
   CASE AliVentanaCen="CENTRAR"
**      SET CENTERWINDOW RELATIVE PARENT
      SetProperty(NonVentanaCen,"Col",((AnchoVentanaPrin-GetProperty(NonVentanaCen,"Width"))/2)+ColVentanaPrin)
      SetProperty(NonVentanaCen,"Row",((AltoVentanaPrin-GetProperty(NonVentanaCen,"Height"))/2)+LinVentanaPrin)
   CASE AliVentanaCen="HORIZONTAL"
      IF AnchoVentanaPrin-GetProperty(NonVentanaCen,"Width")>0
         SetProperty(NonVentanaCen,"Col",((AnchoVentanaPrin-GetProperty(NonVentanaCen,"Width"))/2)+ColVentanaPrin)
      ELSE
         SetProperty(NonVentanaCen,"Col",0)
      ENDIF
   CASE AliVentanaCen="VERTICAL"
      IF AltoVentanaPrin-GetProperty(NonVentanaCen,"Height")>0
         SetProperty(NonVentanaCen,"Row",((AltoVentanaPrin-GetProperty(NonVentanaCen,"Height"))/2)+LinVentanaPrin)
      ELSE
         SetProperty(NonVentanaCen,"Row",0)
      ENDIF
   CASE AliVentanaCen="ALINEAR"
      SetProperty(NonVentanaCen,"Col",GetProperty(NonVentanaCen,"Col")+ColVentanaPrin)
      SetProperty(NonVentanaCen,"Row",GetProperty(NonVentanaCen,"Row")+LinVentanaPrin)
   ENDCASE
ENDIF




/*
**Mostrar ventanas del escritorio***
hWnd2 := GetWindow(hWnd1, 2) //2-GW_HWNDNEXT
msgbox(str(hWnd1)+" "+GetWindowText(hWnd1)+" "+str(hWnd2)+" "+GetWindowText(hWnd2))
SetWindowPos(hWnd2,0,0,0,0,0)
hWnd2 := GetWindow(hWnd1, 2) //2-GW_HWNDNEXT
msgbox(str(hWnd1)+" "+GetWindowText(hWnd1)+" "+str(hWnd2)+" "+GetWindowText(hWnd2))

ShowWindow(hWnd2, 1) //0-SW_HIDE (ocultar) 9-SW_RESTORE 3-SW_SHOWMAXIMIZED 1-SW_SHOWNORMAL
DO EVENTS
msgbox(str(hWnd1)+" "+GetWindowText(hWnd1)+" "+str(hWnd2)+" "+GetWindowText(hWnd2))

**Mostrar todas las ventanas
hWnd4:=hWnd1
hWnd5:=-1
Ventana1.win1.AddItem({str(hWnd1),GetWindowText(hWnd1)})
FOR N=1 TO 100
   DO EVENTS
   hWnd5 := GetWindow(hWnd4, 2)
   IF LEN(RTRIM(GetWindowText(hWnd5)))>0
      Ventana1.win1.AddItem({str(hWnd5),GetWindowText(hWnd5)})
   ENDIF
   hWnd4:=hWnd5
NEXT

STATIC FUNCTION INI_HANDS_MOST()
hWnd5:=VAL(Ventana1.win1.Cell(Ventana1.win1.Value,1))
SetWindowPos(hWnd5,0,0,0,0,0,1)
**Mostrar todas las ventanas
Ventana1.win2.DeleteAllItems
hWnd4:=hWnd1
hWnd5:=-1
Ventana1.win2.AddItem({str(hWnd1),GetWindowText(hWnd1)})
FOR N=1 TO 100
   DO EVENTS
   hWnd5 := GetWindow(hWnd4, 2)
   IF LEN(RTRIM(GetWindowText(hWnd5)))>0
      Ventana1.win2.AddItem({str(hWnd5),GetWindowText(hWnd5)})
   ENDIF
   hWnd4:=hWnd5
NEXT
*/

