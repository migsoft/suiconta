#include "minigui.ch"

procedure AsiCie()

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 650 HEIGHT 450 ;
      TITLE 'Asiento de cierre automatico' ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_Fec1 VALUE 'Fecha asiento' AUTOSIZE TRANSPARENT
      @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA2

      @ 15,300 LABEL L_Nasi1 VALUE 'Asiento regularizacion' AUTOSIZE TRANSPARENT
      @ 10,430 TEXTBOX T_Nasi1 WIDTH 100 VALUE 999997 NUMERIC RIGHTALIGN

      @ 45,300 LABEL L_Nasi2 VALUE 'Asiento de cierre' AUTOSIZE TRANSPARENT
      @ 40,430 TEXTBOX T_Nasi2 WIDTH 100 VALUE 999998 NUMERIC RIGHTALIGN

      @ 70,010 GRID GR_Cierre ;
      HEIGHT 250 ;
      WIDTH 600 ;
      HEADERS {'Cuenta','Descripcion','Asiento','Debe','Haber'} ;
      WIDTHS {80,170,120,100,100 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT, ;
               BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT}

      Menu_Grid("W_Imp1","GR_Cierre","MENU",,{"COPREGPP","COPTABPP"})

      @330,010 PROGRESSBAR P_Progres1 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH
      @330,310 PROGRESSBAR P_Progres2 RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH
      @360,010 PROGRESSBAR P_Progres3 RANGE 0 , 100 WIDTH 600 HEIGHT 20 SMOOTH

      @390, 10 BUTTON B_Calcular CAPTION 'Calcular' WIDTH 90 HEIGHT 25 ;
               ACTION AsiCie1()

      @390,110 BUTTON B_Conta CAPTION 'Contabilizar' WIDTH 90 HEIGHT 25 ;
               ACTION AsiCie2()

      @390,210 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION AsiCie1()
FIC_CIERRE(00000000,99999999,DMA,DMA2,1,1)
AbrirDBF1("FINSUIC",,,"Exclusive",gettempdir(),1)
SELEC 1
W_Imp1.GR_Cierre.DeleteAllItems
W_Imp1.P_Progres1.RangeMax:=LASTREC()
W_Imp1.P_Progres1.Value:=0
W_Imp1.P_Progres2.Value:=0
aSAL1:={}
aSAL2:={}
TOT1:=0
TOT2:=0
GO TOP
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres1.Value:=W_Imp1.P_Progres1.Value+1
   DO EVENTS

   TOTC:=SALDO+DEBE-HABER
   IF TOTC=0
      SKIP
      LOOP
   ENDIF

   IF TOTC<0
      DEB2:=TOTC*-1
      HAB2:=0
   ELSE
      DEB2:=0
      HAB2:=TOTC
   ENDIF

   IF CODCTA>=60000000
      AADD(aSAL1,{CODCTA,RTRIM(NOMCTA),"Asiento de Regularizacion",DEB2,HAB2})
      TOT1:=TOT1+TOTC
   ELSE
      AADD(aSAL2,{CODCTA,RTRIM(NOMCTA),"Asiento de Cierre",DEB2,HAB2})
      TOT2:=TOT2+TOTC
   ENDIF

   SKIP
ENDDO

ASORT(aSAL1,,, { |x, y| x[1] < y[1] })

IF TOT1>0
   DEB2:=TOT1
   HAB2:=0
ELSE
   DEB2:=0
   HAB2:=TOT1*-1
ENDIF
AADD(aSAL1,{12900000,"Perdidas y ganancias","Asiento de Regularizacion",DEB2,HAB2})

***INTRODUCIR SALDO PARA CIERRE***
CTA129:=ASCAN(aSAL2,{|AVAL| AVAL[1]=12900000})
IF CTA129>0
   SALDO129:=HAB2-DEB2+aSAL2[CTA129,4]-aSAL2[CTA129,5] //HAB2,DEB2 SALDO AL REVES
   IF SALDO129>0
      aSAL2[CTA129,4]:=SALDO129
      aSAL2[CTA129,5]:=0
   ELSE
      aSAL2[CTA129,4]:=0
      aSAL2[CTA129,5]:=SALDO129*-1
   ENDIF
ELSE
   AADD(aSAL2,{12900000,"Perdidas y ganancias","Asiento de Cierre",HAB2,DEB2}) //SALDO AL REVES
ENDIF
TOT2:=TOT2+DEB2-HAB2
***FIN INTRODUCIR SALDO PARA CIERRE***

ASORT(aSAL2,,, { |x, y| x[1] < y[1] })

IF ROUND(TOT2,4)<>0
   IF TOT2>0
      DEB2:=TOT2
      HAB2:=0
   ELSE
      DEB2:=0
      HAB2:=TOT2*-1
   ENDIF
   AADD(aSAL2,{55559999,"Asiento de Cierre DESCUADRE","Asiento de Cierre",HAB2,DEB2})
ENDIF


W_Imp1.P_Progres2.RangeMax:=LEN(aSAL1)+LEN(aSAL2)
W_Imp1.P_Progres2.Value:=0
FOR N=1 TO LEN(aSAL1)
   W_Imp1.P_Progres2.Value:=W_Imp1.P_Progres2.Value+1
   DO EVENTS
   W_Imp1.GR_Cierre.AddItem(aSAL1[N])
   W_Imp1.GR_Cierre.Value:=W_Imp1.GR_Cierre.ItemCount
   W_Imp1.GR_Cierre.Refresh
NEXT

FOR N=1 TO LEN(aSAL2)
   W_Imp1.P_Progres2.Value:=W_Imp1.P_Progres2.Value+1
   DO EVENTS
   W_Imp1.GR_Cierre.AddItem(aSAL2[N])
   W_Imp1.GR_Cierre.Value:=W_Imp1.GR_Cierre.ItemCount
   W_Imp1.GR_Cierre.Refresh
NEXT

Return Nil


STATIC FUNCTION AsiCie2()
IF MSGYESNO("Desea hacer el asiento de regulacion y cierre")=.F.
   RETURN
ENDIF

AbrirDBF("Cierre")
GO TOP
DO WHILE .NOT. EOF()
   DO EVENTS
   IF NASI=999997 .OR. NASI=999998
      IF RLOCK()
         DELETE
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
   ENDIF
   SKIP
ENDDO

W_Imp1.P_Progres3.RangeMax:=W_Imp1.GR_Cierre.ItemCount
W_Imp1.P_Progres3.Value:=0
NASI2:=1
NAPU2:=1
FOR N=1 TO W_Imp1.GR_Cierre.ItemCount
   W_Imp1.P_Progres3.Value:=W_Imp1.P_Progres3.Value+1
   DO EVENTS

   IF AT("REG",UPPER(W_Imp1.GR_Cierre.Cell(N,3)))<>0
      NASI2:=999997
   ELSE
      IF NASI2=999997
         NAPU2:=1
      ENDIF
      NASI2:=999998
   ENDIF

   APPEND BLANK
   IF RLOCK()
      REPLACE NASI WITH NASI2
      REPLACE APU WITH IF(NAPU2>=990,990,NAPU2++)
      REPLACE NEMP WITH NUMEMP
      REPLACE FECHA WITH W_Imp1.D_Fec1.Value
      REPLACE CODCTA WITH W_Imp1.GR_Cierre.Cell(N,1)
      REPLACE NOMAPU WITH W_Imp1.GR_Cierre.Cell(N,3)
      REPLACE DEBE   WITH W_Imp1.GR_Cierre.Cell(N,4)
      REPLACE HABER  WITH W_Imp1.GR_Cierre.Cell(N,5)
      DBCOMMIT()
      DBUNLOCK()
   ENDIF
NEXT

MSGBOX("Se ha realizado el asiento de regulacion y cierre con exito")

