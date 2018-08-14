#include "minigui.ch"

FUNCTION MsgEspera(cMensaje,nTiempo)
   DEFINE WINDOW WinEspera1 AT 0,0 WIDTH 650 HEIGHT 100 TITLE 'Mensajes del Sistema...' MODAL NOSYSMENU
   @ 20,5 LABEL lblMensajes VALUE AllTrim(cMensaje) WIDTH 640 HEIGHT 40 CENTERALIGN BOLD
   END WINDOW
   VentanaCentrar("WinEspera1","Ventana1")
   ACTIVATE WINDOW WinEspera1 NOWAIT

   IF Empty(nTiempo)
      nTiempo:=3
   ENDIF
   DO WHILE nTiempo>=0
      DO EVENTS
      Inkey(.5)
      nTiempo:=nTiempo-0.5
   ENDDO

   RELEASE WINDOW WinEspera1
RETURN(NIL)




FUNCTION PonerEspera(cMensaje)
   IF IsWindowDefined(WinEspera2)=.T.
      DECLARE WINDOW WinEspera2
      WinEspera2.restore
      WinEspera2.Show
      WinEspera2.SetFocus
   ELSE
      DEFINE WINDOW WinEspera2 AT 0,0 WIDTH 650 HEIGHT 100 TITLE 'Mensajes del Sistema...' MODAL NOSYSMENU BACKCOLOR MiColor("AZULCLARO")
      END WINDOW
      VentanaCentrar("WinEspera2","Ventana1")
      ACTIVATE WINDOW WinEspera2 NOWAIT
   ENDIF
   IF IsControlDefined(lblMensajes,WinEspera2)=.T.
      WinEspera2.lblMensajes.Value:=AllTrim(cMensaje)
   ELSE
      @ 20,5 LABEL lblMensajes OF WinEspera2 VALUE AllTrim(cMensaje) WIDTH 640 HEIGHT 40 CENTERALIGN BOLD TRANSPARENT
   ENDIF
   CURSORWAIT()
   DO EVENTS
RETURN(NIL)


FUNCTION PonerEsperaTexto(cMensaje)
   IF IsWindowDefined(WinEspera2)=.T.
      IF IsControlDefined(lblMensajes,WinEspera2)=.T.
         WinEspera2.lblMensajes.Value:=AllTrim(cMensaje)
      ENDIF
   ENDIF
   DO EVENTS
RETURN(NIL)


FUNCTION QuitarEspera()
   CURSORARROW()
   IF IsWindowDefined(WinEspera2)=.T.
      WinEspera2.Release
   ENDIF
   DO EVENTS
RETURN(NIL)


/*
 * MsgOptions([cText], [cTitle], [cImage], aOptions, [nDefaultOption], [nSeconds])

DECLARE WINDOW _Options

Function MsgOptions(cText, cTitle, cImage, aOptions, nDefaultOption, nSeconds )
   Local nItem:=0, nBtnWidth:=0, aBtn:=Array(Len(aOptions)), aImgInfo
   Local nBtnPosX:=10, nBtnPosY:=85, cOption:=""

   Default cText           To "Seleccione una opción..."
   Default cTitle          To "Error !" //***UPPER(_HMG_BRWLangError [10])+"!"
   Default cImage          To ""
   Default nDefaultOption  To 1
   Default nSeconds        To 0

   DEFINE FONT _Font_Options FONTNAME "MS Sans Serif" SIZE 9

   //Calcular anchura máxima de un botón para igualarlos todos
   For nItem:=1 To Len(aOptions)
      aOptions[nItem]:=Alltrim(aOptions[nItem])
      nBtnWidth:=Max( GetTextWidth(, aOptions[nItem], GetFontHandle("_Font_Options")), nBtnWidth )
   Next
   nBtnWidth+=5

   DEFINE WINDOW _Options  ;
      AT 0,0               ;
      WIDTH (Len(aOptions)*(10+nBtnWidth))+15 ;
      HEIGHT 155           ;
      TITLE cTitle         ;
      ICON "demo.ico"      ;
      MODAL                ;
      NOSIZE               ;
      ON RELEASE IF( IsControlDefined( Timer_1, _Options ), _Options.Timer_1.Release, )

      ON KEY ESCAPE ACTION _Options.Release

      If !Empty(cImage)
         aImgInfo := BmpSize(cImage)
         If !Empty(aImgInfo [BM_WIDTH])
           @ 20, 10 IMAGE _Image PICTURE (cImage) WIDTH aImgInfo [BM_WIDTH] HEIGHT aImgInfo [BM_HEIGHT]
           @ 40, 55 LABEL _Label VALUE cText WIDTH (Len(aOptions)*(10+nBtnWidth))-50 HEIGHT 30 ;
             TRANSPARENT CENTERALIGN FONT "_Font_Options"
         Endif
      Else
           @ 40, 10 LABEL _Label VALUE cText WIDTH (Len(aOptions)*(10+nBtnWidth))-10 HEIGHT 30 ;
             TRANSPARENT CENTERALIGN FONT "_Font_Options"
      Endif

      For nItem:=1 To Len(aOptions)
         aBtn[nItem]:="_Btn_"+Ltrim(Str(nItem))
         cOption:=aBtn[nItem]
         @ nBtnPosY, nBtnPosX BUTTON &cOption CAPTION aOptions[nItem] WIDTH nBtnWidth HEIGHT 25 FONT "_Font_Options" ;
                   ACTION ( cOption:=GetProperty("_Options", This.Name, "Caption"), _Options.Release )
         nBtnPosX+=nBtnWidth+10
      Next

      DoMethod("_Options", aBtn[nDefaultOption], "SetFocus")

      If nSeconds>0
         DEFINE TIMER Timer_1 Interval nSeconds*1000  ;
                ACTION ( cOption:=aOptions[nDefaultOption], _Options.Release )
      Endif

   END WINDOW
   VentanaCentrar("_Options","Ventana1")
   ACTIVATE WINDOW _Options

   RELEASE FONT _Font_Options

Return Ascan(aOptions,Alltrim(cOption))

*/
