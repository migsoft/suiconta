#include "minigui.ch"

procedure Asimpor()
   TituloImp:="Importar asientos"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 390 ;
      TITLE TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 10,10 BUTTONEX Bt_Buscardir CAPTION 'Ruta' ICON icobus('buscar') ;
              WIDTH 90 HEIGHT 25 ;
              ACTION W_Imp1.T_Ruta.Value:=GetFolder(,W_Imp1.T_Ruta.Value) ;
              NOTABSTOP
      @ 10,100 TEXTBOX T_Ruta WIDTH 290 VALUE ''


      @ 45,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 40,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1
      @ 75,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @ 70,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

      @105,10 LABEL L_Nasi1 VALUE 'Desde asiento' AUTOSIZE TRANSPARENT
      @100,110 TEXTBOX T_Nasi1 WIDTH 100 VALUE 0 NUMERIC RIGHTALIGN
      @135,10 LABEL L_Nasi2 VALUE 'Hasta asiento' AUTOSIZE TRANSPARENT
      @130,110 TEXTBOX T_Nasi2 WIDTH 100 VALUE 999999999 NUMERIC RIGHTALIGN


      @160,10 CHECKBOX Asientos  CAPTION 'Fichero de asientos' WIDTH 220 VALUE .T.
      @190,10 CHECKBOX emitidas  CAPTION 'Fichero de facturas emitidas' WIDTH 220 VALUE .T.
      @220,10 CHECKBOX cobros    CAPTION 'Fichero de cobros facturas emitidas' WIDTH 220 VALUE .T.
      @250,10 CHECKBOX recibidas CAPTION 'Fichero de facturas recibidas' WIDTH 220 VALUE .T.
      @280,10 CHECKBOX pagos     CAPTION 'Fichero de pagos facturas recibidas' WIDTH 220 VALUE .T.

      @330, 10 BUTTONEX B_Importar CAPTION 'Importar' ICON icobus('guardar') WIDTH 90 HEIGHT 25 ;
               ACTION ( Asimpori() , W_Imp1.release )

      @330,110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION Asimpori()
IF W_Imp1.T_Ruta.Value=RUTAEMPRESA
   MSGSTOP("No se pueden importar asientos de la misma empresa")
   RETURN
ENDIF

PonerEspera("Procesando los datos....")

RUTA2:=RTRIM(W_Imp1.T_Ruta.Value)
IF W_Imp1.Asientos.Value=.T.
   IF FILE(RUTA2+"\APUNTES.DBF")=.T.
      AbrirDBF("APUNTES")
      APPEND FROM &RUTA2\APUNTES FOR FECHA>=W_Imp1.D_Fec1.Value .AND. FECHA<=W_Imp1.D_Fec2.Value .AND. ;
               NASI>=W_Imp1.T_Nasi1.Value .AND. NASI<=W_Imp1.T_Nasi2.Value
   ENDIF
ENDIF

IF W_Imp1.emitidas.Value=.T.
   IF FILE(RUTA2+"\FAC92.DBF")=.T.
      AbrirDBF("FAC92")
      APPEND FROM &RUTA2\FAC92 FOR FFAC>=W_Imp1.D_Fec1.Value .AND. FFAC<=W_Imp1.D_Fec2.Value .AND. ;
               ASIENTO>=W_Imp1.T_Nasi1.Value .AND. ASIENTO<=W_Imp1.T_Nasi2.Value
   ENDIF
ENDIF

IF W_Imp1.cobros.Value=.T.
   IF FILE(RUTA2+"\COBROS.DBF")=.T.
      AbrirDBF("COBROS")
      APPEND FROM &RUTA2\COBROS FOR FCOB>=W_Imp1.D_Fec1.Value .AND. FCOB<=W_Imp1.D_Fec2.Value .AND. ;
               NASI>=W_Imp1.T_Nasi1.Value .AND. NASI<=W_Imp1.T_Nasi2.Value
   ENDIF
ENDIF

IF W_Imp1.recibidas.Value=.T.
   IF FILE(RUTA2+"\FACREB.DBF")=.T.
      AbrirDBF("FACREB")
      APPEND FROM &RUTA2\FACREB FOR FREG>=W_Imp1.D_Fec1.Value .AND. FREG<=W_Imp1.D_Fec2.Value .AND. ;
               ASIENTO>=W_Imp1.T_Nasi1.Value .AND. ASIENTO<=W_Imp1.T_Nasi2.Value
   ENDIF
ENDIF

IF W_Imp1.pagos.Value=.T.
   IF FILE(RUTA2+"\PAGOS.DBF")=.T.
      AbrirDBF("PAGOS")
      APPEND FROM &RUTA2\PAGOS FOR FPAG>=W_Imp1.D_Fec1.Value .AND. FPAG<=W_Imp1.D_Fec2.Value .AND. ;
               NASI>=W_Imp1.T_Nasi1.Value .AND. NASI<=W_Imp1.T_Nasi2.Value
   ENDIF
ENDIF

QuitarEspera()

MSGINFO("Datos importados")

RETURN



