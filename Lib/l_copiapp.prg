#include "minigui.ch"
Function CopiaPP(LLAMADA,Ventana1)

Texto1:=""

//***aCtrls := _GetArrayOfAllControlsForForm ( Ventana1 )
oVentana1 := GetExistingFormObject( Ventana1 )

//***for nI := 1 to len( aCtrls )
for nI := 1 to len( oVentana1:aControls )
   //***CtrlName := aCTrls[nI]
   CtrlName := oVentana1:aControls[nI]:Name
   xVal := GetProperty( Ventana1, CtrlName, "value" )
   IF LLAMADA="TODOS" .OR. UPPER(CtrlName)="T_"
      Texto1:=Texto1+CtrlName+" "+xChar( xVal )+HB_OsNewLine()
   ENDIF
next

//***CopyToClipboard(Texto1)
SetClipboardText(Texto1)
MsgInfo(Texto1,"Portapapeles")

Return Nil

/******************************************************************************/
Function _GetArrayOfAllControlsForForm ( Ventana1 )
/******************************************************************************/
Local nFormHandle , i , nControlCount , aRetVal := {} , x

oVentana1 := GetExistingFormObject( Ventana1 )
//***nFormHandle := GetFormHandle ( Ventana1 )
nFormHandle := oVentana1:hWnd
//***nControlCount := Len ( _HMG_aControlHandles )
nControlCount := Len ( oVentana1:aControls )
For i := 1 To nControlCount
   //***If _HMG_aControlParentHandles[i] == nFormHandle
    If oVentana1:aControls[i] == nFormHandle
      //***If ValType( _HMG_aControlHandles[i] ) == 'N'
       If ValType( oVentana1:aControls[i] ) == 'N'
         //***IF ! Empty( _HMG_aControlNames[i] )
         IF ! Empty( oVentana1:aControls[i]:names )
            //***If Ascan( aRetVal, _HMG_aControlNames[i] ) == 0
            If Ascan( aRetVal, oVentana1:aControls[i]:names ) == 0
               //***Aadd( aRetVal, _HMG_aControlNames[i] )
               Aadd( aRetVal, oVentana1:aControls[i]:names )
            EndIf
         ENDIF
      //***ElseIf ValType( _HMG_aControlHandles [i] ) == 'A'
      ElseIf ValType( oVentana1:aControls[i] ) == 'A'
         //***For x := 1 To Len ( _HMG_aControlHandles[i] )
         For x := 1 To Len ( oVentana1:aControls[i] )
            //***IF !Empty( _HMG_aControlNames[i] )
            IF !Empty( oVentana1:aControls[i]:names )
               //***If Ascan( aRetVal, _HMG_aControlNames[i] ) == 0
               If Ascan( aRetVal,oVentana1:aControls[i]:names ) == 0
                  //***Aadd( aRetVal, _HMG_aControlNames [i] )
                  Aadd( aRetVal, oVentana1:aControls[i]:names )
               EndIf
            ENDIF
         Next x
      EndIf
   EndIf
Next i
Return Asort( aRetVal )
