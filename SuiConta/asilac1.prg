#include "minigui.ch"
#include "winprint.ch"

procedure AsilApeCie1()
   LOCAL Color_gr1 := { |x,nItem| IF( x[5]<>0 , RGB(255,200,200) , RGB(255,255,255)) }

   TituloImp:="Comprobacion asiento cierre/apertura"
   FEC1:=CTOD( "01-"+STR(MONTH(DATE()),2)+"-"+STR(YEAR(DATE()),4) )

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 800 HEIGHT 600 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @010,010 GRID GR_Fic1 ;
      HEIGHT 350 ;
      WIDTH 775 ;
      HEADERS {'Codigo','Descripcion','Cierre '+STR(EJERANY-1,4),'Apertura '+STR(EJERANY,4),'Diferencia'} ;
      WIDTHS { 100,300,100,100,100 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      DYNAMICBACKCOLOR { Color_gr1,Color_gr1,Color_gr1,Color_gr1,Color_gr1 }

      @366,010 PROGRESSBAR Progress3 WIDTH 775 HEIGHT 15 SMOOTH

      @375,150 LABEL L_Error VALUE 'ATENCION existen diferencias en el cierre y la apertura' ;
               WIDTH 650 HEIGHT 30 FONT "Arial" SIZE 18 BOLD FONTCOLOR MICOLOR("ROJO") TRANSPARENT CENTERALIGN INVISIBLE

      @380,010 CHECKBOX SiCero CAPTION 'Incluir saldo cero' width 140 value .F.

      @405,010 LABEL L_FicCie VALUE 'Empresa de asiento de cierre' AUTOSIZE TRANSPARENT
      @430,010 TEXTBOX T_FicCie WIDTH 300 MAXLENGTH 30
      @460,010 PROGRESSBAR Progress1 WIDTH 300 HEIGHT 15 SMOOTH

      @485,010 LABEL L_FicApe VALUE 'Empresa de asiento de apertura' AUTOSIZE TRANSPARENT
      @510,010 TEXTBOX T_FicApe WIDTH 300 MAXLENGTH 30
      @540,010 PROGRESSBAR Progress2 WIDTH 300 HEIGHT 15 SMOOTH


*draw rectangle in window W_Imp1 at 230,410 to 232,790 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @415,410 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @410,500 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @445,620 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @440,680 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @440,410 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @470,410 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @500,410 LABEL L_Progress_1 VALUE 'Imprimiendo' AUTOSIZE TRANSPARENT
      @500,500 PROGRESSBAR Progress_1 WIDTH 280 HEIGHT 15 SMOOTH

      @525,410 BUTTONEX B_Comprobar CAPTION 'Comprobar' WIDTH 90 HEIGHT 25 ;
               ACTION AsilApeCie1_INI()

      @525,510 BUTTONEX B_Imprimir CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION AsilApeCie1i("IMPRESORA")
W_Imp1.B_Imprimir.Enabled  := .F.

      @525,610 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil




procedure AsilApeCie1_INI()
W_Imp1.B_Comprobar.Enabled := .F.

AbrirDBF("EMPRESA",,,,RUTAPROGRAMA,1)
SEEK NUMEMP
NOMEMPAPE:=LTRIM(STR(NEMP))+"-"+RTRIM(EMP)+" ("+LTRIM(STR(EJERCICIO))+")"
W_Imp1.T_FicApe.Value:=NOMEMPAPE
NUMEMPCIE:=NEMPANT
SEEK NUMEMPCIE
IF NUMEMPCIE=0 .OR. EOF()
   MSGSTOP("No se ha encontrado la empresa anterior"+HB_OsNewLine()+"Empresa numero: "+LTRIM(STR(NUMEMPCIE)) )
   W_Imp1.release
   RETURN
ENDIF
NOMEMPCIE:=LTRIM(STR(NEMP))+"-"+RTRIM(EMP)+" ("+LTRIM(STR(EJERCICIO))+")"
W_Imp1.T_FicCie.Value:=NOMEMPCIE
RUTEMPCIE:=RUTAPROGRAMA+"\"+ALLTRIM(RUTA)


***ASIENTO DE CIERRE***
AbrirDBF("CIERRE",,,,RUTEMPCIE)
AbrirDBF("CUENTAS",,,,RUTEMPCIE)
AbrirDBF("APUNTES")

aFIN:={}

AbrirDBF("CIERRE",,,,RUTEMPCIE)
SET DELETED OFF
GO TOP
W_Imp1.Progress1.RangeMax:=LASTREC()
W_Imp1.Progress1.Value:=0
DO WHILE .NOT. EOF()
   W_Imp1.Progress1.Value:=W_Imp1.Progress1.Value+1
   DO EVENTS

   IF DELETED()=.T.
      SKIP
      LOOP
   ENDIF
   IF NASI<>999998 .OR. DEBE-HABER=0
      SKIP
      LOOP
   ENDIF
   CODCTA2:=CODCTA
   SALDO2:=DEBE-HABER

   CODCTAB:=ASCAN(aFIN,{ |AVAL| AVAL[1]==CODCTA} )
   IF CODCTAB=0 .OR. CODCTAB=NIL
      AbrirDBF("CUENTAS",,,,RUTEMPCIE)
      SEEK CODCTA2
      AADD(aFIN,{CODCTA2,NOMCTA,0,0,0})
      CODCTAB:=LEN(aFIN)
   ENDIF
   aFIN[CODCTAB,3]:=aFIN[CODCTAB,3]+(SALDO2*-1) //CAMBIAR SIGNO

   AbrirDBF("CIERRE",,,,RUTEMPCIE)
   SKIP
ENDDO
SET DELETED ON
***FIN ASIENTO DE CIERRE***


***ASIENTO DE APERTURA***
AbrirDBF("APUNTES")
SET DELETED OFF
GO TOP
W_Imp1.Progress2.RangeMax:=LASTREC()
W_Imp1.Progress2.Value:=0
DO WHILE .NOT. EOF()
   W_Imp1.Progress2.Value:=W_Imp1.Progress2.Value+1
   DO EVENTS

   IF DELETED()=.T.
      SKIP
      LOOP
   ENDIF
   IF NASI<>1 .OR. DEBE-HABER=0
      SKIP
      LOOP
   ENDIF
   CODCTA2:=CODCTA
   SALDO2:=DEBE-HABER

   CODCTAB:=ASCAN(aFIN,{ |AVAL| AVAL[1]==CODCTA} )
   IF CODCTAB=0 .OR. CODCTAB=NIL
      AbrirDBF("CUENTAS",,,,RUTEMPCIE)
      SEEK CODCTA2
      AADD(aFIN,{CODCTA2,NOMCTA,0,0,0})
      CODCTAB:=LEN(aFIN)
   ENDIF
   aFIN[CODCTAB,4]:=aFIN[CODCTAB,4]+SALDO2

   AbrirDBF("APUNTES")
   SKIP
ENDDO
SET DELETED ON
***FIN ASIENTO DE APERTURA***


MAL:=0
W_Imp1.GR_Fic1.DeleteAllItems
W_Imp1.Progress3.RangeMax:=LEN(aFIN)
W_Imp1.Progress3.Value:=0
IF LEN(aFIN)>0
   FOR N=1 TO LEN(aFIN)
   W_Imp1.Progress3.Value:=N
      DO EVENTS
      aFIN[N,5]:=aFIN[N,3]-aFIN[N,4]
      IF W_Imp1.SiCero.Value=.T. .OR. aFIN[N,5]<>0
         W_Imp1.GR_Fic1.AddItem(aFIN[N])
      ENDIF
      IF aFIN[N,5]<>0 .AND. MAL=0
         MAL:=1
      ENDIF
   NEXT
ENDIF

IF MAL=1
   W_Imp1.L_Error.Value:="ATENCION existen diferencias en el cierre y la apertura"
   W_Imp1.L_Error.FontColor:=MICOLOR("ROJO")
   W_Imp1.L_Error.Visible:=.T.
ELSE
   W_Imp1.L_Error.Value:="No existen diferencias en el cierre y la apertura"
   W_Imp1.L_Error.FontColor:=MICOLOR("VERDE")
   W_Imp1.L_Error.Visible:=.T.
ENDIF

W_Imp1.B_Imprimir.Enabled  := .T.



procedure AsilApeCie1i(LLAMADA)
   local oprint

   IF W_Imp1.GR_Fic1.ItemCount=0
      MsgExclamation("No hay datos en los parametros introducidos","Informacion")
      RETURN
   ENDIF

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .F. , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloImp)
   oprint:setpreviewsize(2)  // tamaño del preview
   oprint:beginpage()

PAG:=0
LIN:=0
aCOL:={}

AADD(aCOL, 10) //1-tamaño letra
AADD(aCOL, 20) //2-codigo cuenta
AADD(aCOL, 40) //3-nombre cuenta
AADD(aCOL,130) //4-saldo cierre
AADD(aCOL,160) //5-saldo apertura
AADD(aCOL,190) //6-diferencia

W_Imp1.Progress_1.RangeMax:=W_Imp1.GR_Fic1.ItemCount
W_Imp1.Progress_1.Value:=0
FOR N=1 TO W_Imp1.GR_Fic1.ItemCount
   W_Imp1.Progress_1.Value:=W_Imp1.Progress_1.Value+1
   DO EVENTS

   IF LIN>=265 .OR. PAG=0
      IF PAG<>0
         oprint:printdata(LIN+5,105,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

      LIN:=45
      oprint:printdata(LIN,aCOL[2],"Cuenta","times new roman",aCOL[1],.F.,,"L",)
      oprint:printdata(LIN,aCOL[3],"Descripcion","times new roman",aCOL[1],.F.,,"L",)
      oprint:printdata(LIN,aCOL[4],"Cierre "+STR(EJERANY-1,4),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[5],"Apertura "+STR(EJERANY,4),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[6],"Diferencia","times new roman",aCOL[1],.F.,,"R",)
      oprint:printline(LIN+4,20,LIN+4,aCOL[6],,0.5)

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN,aCOL[2],W_Imp1.GR_Fic1.Cell(N,1),"times new roman",aCOL[1],.F.,,"L",)
   oprint:printdata(LIN,aCOL[3],W_Imp1.GR_Fic1.Cell(N,2),"times new roman",aCOL[1],.F.,,"L",)
   oprint:printdata(LIN,aCOL[4],MIL(W_Imp1.GR_Fic1.Cell(N,3),15,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
   oprint:printdata(LIN,aCOL[5],MIL(W_Imp1.GR_Fic1.Cell(N,4),15,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
   oprint:printdata(LIN,aCOL[6],MIL(W_Imp1.GR_Fic1.Cell(N,5),15,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)

   LIN:=LIN+5

NEXT

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil



