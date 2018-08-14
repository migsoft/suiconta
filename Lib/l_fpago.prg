#include "minigui.ch"

FUNCTION Fpago(CodPago,Cod_Ventana1,Cod_Control1,Cod_Control2)
   CODPAGO2:=PADR(LTRIM(STR(CODPAGO)),4," ")
   Y1:=VAL(SUBSTR(CODPAGO2,1,1))+1 //FORMA DE PAGO
   Y2:=VAL(SUBSTR(CODPAGO2,2,1))+1 //NUMERO DE GIROS
   Y3:=VAL(SUBSTR(CODPAGO2,3,1))+1 //DIAS 1º VTO.
   Y4:=VAL(SUBSTR(CODPAGO2,4,1))+1 //DIAS RESTO

   IF Y2<=0 .OR. Y2>9
      Y2:=1
   ENDIF
   IF Y2<=0 .OR. Y2>7
      Y2:=1
   ENDIF
   IF Y3<=0 .OR. Y3>10
      Y3:=1
   ENDIF
   IF Y4<=0 .OR. Y4>10
      Y4:=1
   ENDIF


   aFpago1:={"Contado","Reposicion","Transferencia","Giro","Talon","Pagare","Reembolso","Cont.Leasing","Cont.Metalico"}
   aFpago2:={"0 Vto.","1 Vto.","2 Vtos.","3 Vtos.","4 Vtos.","5 Vtos.","6 Vtos."}
   aFpago3:={"0 dias","15 dias","30 dias","45 dias","60 dias","75 dias","90 dias","105 dias","120 dias","135 dias"}
   aFpago4:={"0 dias","15 dias","30 dias","45 dias","60 dias","75 dias","90 dias","105 dias","120 dias","135 dias"}

   DEFINE WINDOW WinFpago ;
      AT 0,0     ;
      WIDTH 500 HEIGHT 350 ;
      TITLE 'Forma de pago' ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("VERDECLARO")

      @ 10,010 LABEL L_Fpago1 VALUE 'Forma de pago' AUTOSIZE TRANSPARENT
      @ 10,130 LABEL L_Fpago2 VALUE 'Vencimientos' AUTOSIZE TRANSPARENT
      @ 10,230 LABEL L_Fpago3 VALUE 'Dias 1º vto.' AUTOSIZE TRANSPARENT
      @ 10,330 LABEL L_Fpago4 VALUE 'Dias resto vtos.' AUTOSIZE TRANSPARENT

      @ 30,010 RADIOGROUP Fpago1 OPTIONS aFpago1 VALUE Y1 WIDTH 100 SPACING 20 ON CHANGE Fpago_Ver()
      @ 30,130 RADIOGROUP Fpago2 OPTIONS aFpago2 VALUE Y2 WIDTH 100 SPACING 20 ON CHANGE Fpago_Ver()
      @ 30,230 RADIOGROUP Fpago3 OPTIONS aFpago3 VALUE Y3 WIDTH 100 SPACING 20 ON CHANGE Fpago_Ver()
      @ 30,330 RADIOGROUP Fpago4 OPTIONS aFpago4 VALUE Y4 WIDTH 100 SPACING 20 ON CHANGE Fpago_Ver() 

      @220,10 TEXTBOX T_FPagoN WIDTH 50 VALUE CodPago NUMERIC READONLY //INVISIBLE
      @250,10 TEXTBOX T_FPagoC WIDTH 400 VALUE VENCI_NC(CodPago) READONLY

      @280,010 BUTTONEX Bt_Aceptar CAPTION 'Aceptar' ICON icobus('aceptar') WIDTH 90 HEIGHT 25 ;
               ACTION Fpago_Terminar(CodPago,Cod_Ventana1,Cod_Control1,Cod_Control2)

      @280,110 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION WinFpago.release

   END WINDOW
   VentanaCentrar("WinFpago","Ventana1")
   ACTIVATE WINDOW WinFpago

RETURN



STATIC FUNCTION Fpago_Terminar(Cod_Codigo,Cod_Ventana1,Cod_Control1,Cod_Control2)
   SetProperty(Cod_Ventana1,Cod_Control1,"Value",WinFpago.T_FPagoN.Value)
   IF IsControlDefined(&Cod_Control2,&Cod_Ventana1)=.T.
      SetProperty(Cod_Ventana1,Cod_Control2,"Value",VENCI_NC(WinFpago.T_FPagoN.Value))
   ENDIF
   WinFpago.Release
RETURN


STATIC FUNCTION Fpago_Ver()
IF WinFpago.Fpago1.Value=1 //Contado
   WinFpago.T_FPagoN.Value:=0
ELSE
   WinFpago.T_FPagoN.Value:=VAL( ;
   LTRIM(STR(WinFpago.Fpago1.Value-1))+ ;
   LTRIM(STR(WinFpago.Fpago2.Value-1))+ ;
   LTRIM(STR(WinFpago.Fpago3.Value-1))+ ;
   LTRIM(STR(WinFpago.Fpago4.Value-1)) )
ENDIF
WinFpago.T_FPagoC.Value:=VENCI_NC(WinFpago.T_FPagoN.Value)
RETURN

