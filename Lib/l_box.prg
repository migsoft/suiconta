#include "minigui.ch"

FUNCTION Box3botones(titulo,atexto,boton1,boton2,boton3)
   Box3boton:=0
   COL1:=90+(LEN(atexto)*20)
   COL2:=20+(LEN(atexto)*20)

   DEFINE WINDOW W_Box3botones ;
      AT 0,0 ;
      WIDTH 450 HEIGHT COL1 ;
      TITLE titulo ;
      MODAL      ;
      NOSIZE

      @ 10,10 LABEL L_Box1 VALUE atexto[1] AUTOSIZE
   IF LEN(atexto)>=2
      @ 30,10 LABEL L_Box2 VALUE atexto[2] AUTOSIZE
   ENDIF
   IF LEN(atexto)>=3
      @ 50,10 LABEL L_Box3 VALUE atexto[3] AUTOSIZE
   ENDIF
   IF LEN(atexto)>=4
      @ 70,10 LABEL L_Box4 VALUE atexto[4] AUTOSIZE
   ENDIF
   IF LEN(atexto)>=5
      @ 90,10 LABEL L_Box5 VALUE atexto[5] AUTOSIZE
   ENDIF

      @COL2,010 BUTTON B_Boton1 CAPTION boton1 ACTION Box3botonesF(1) ;
               WIDTH 125 HEIGHT 25
      @COL2,160 BUTTON B_Boton2 CAPTION boton2 ACTION Box3botonesF(2) ;
               WIDTH 125 HEIGHT 25
      @COL2,310 BUTTON B_Boton3 CAPTION boton3 ACTION Box3botonesF(3) ;
               WIDTH 125 HEIGHT 25

   END WINDOW
   VentanaCentrar("W_Box3botones","Ventana1")
   ACTIVATE WINDOW W_Box3botones
   Return(Box3boton)

FUNCTION Box3botonesF(boton)
   Box3boton:=boton
   W_Box3botones.release
   Return(Box3boton)
