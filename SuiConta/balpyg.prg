#include "minigui.ch"
#include "winprint.ch"

procedure balpyg()
   TituloImp:="Balance de perdidas y ganancias"

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 330 ;
      TITLE 'Imprimir: '+TituloImp ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @ 15,10 LABEL L_Fec1 VALUE 'Desde la Fecha' AUTOSIZE TRANSPARENT
      @ 10,110 DATEPICKER D_Fec1 WIDTH 100 VALUE DMA1

      @ 45,10 LABEL L_Fec2 VALUE 'Hasta la Fecha' AUTOSIZE TRANSPARENT
      @ 40,110 DATEPICKER D_Fec2 WIDTH 100 VALUE DMA2

      @ 75,10 LABEL L_Hoja VALUE 'Hoja inicial' AUTOSIZE TRANSPARENT
      @ 70,110 TEXTBOX T_Hoja WIDTH 100 VALUE 1 NUMERIC RIGHTALIGN

      @ 10,250 RADIOGROUP R_Fecha OPTIONS { 'Por fecha' , 'Anual' } VALUE 1 WIDTH 80 

      @130,10 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 300 HEIGHT 20 SMOOTH

draw rectangle in window W_Imp1 at 170,010 to 172,390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @185,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @180,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @215,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @210,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @210, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @240, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @270, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION balpygi()

      @270,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION balpygi()
   aNOM6:={{60,"Compras"},;
           {61,"Variación de existencias"},;
           {62,"Servicios exteriores"},;
           {63,"Tributos"},;
           {64,"Gastos de personal"},;
           {65,"Otros gastos de gestión"},;
           {66,"Gastos financieros"},;
           {67,"Perdidas inmov.y gastos exc."},;
           {68,"Dotaciones amortización"},;
           {69,"Dotaciones a las provisiones"}}
   aNOM7:={{70,"Ventas"},;
           {71,"Variación de existencias"},;
           {72,""},;
           {73,"Trabajos propios"},;
           {74,"Subvenciones a la explotación"},;
           {75,"Otros ingresos de gestión"},;
           {76,"Ingresos financieros"},;
           {77,"Beneficios inmov.e ingres.exc."},;
           {78,""},;
           {79,"Excesos y aplica.provisiones"}}

   IF W_Imp1.R_Fecha.Value=1
      balpygi_fec()
   ELSE
      balpygi_mes()
   ENDIF


STATIC FUNCTION balpygi_fec()
   local oprint

   SUMCTA:=ARRAY(2,11)
   AFILL(SUMCTA[1],0)
   AFILL(SUMCTA[2],0)
   AbrirDBF("APUNTES")
   GO TOP
   W_Imp1.P_Progres.RangeMax:=LASTREC()
   W_Imp1.P_Progres.Value:=0
   DO WHILE .NOT. EOF()
      W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
      IF FECHA>=W_Imp1.D_Fec1.Value .AND. FECHA<=W_Imp1.D_Fec2.Value
         NUM:=VAL(SUBSTR(LTRIM(STR(CODCTA)),2,1))+1
         IF LTRIM(STR(CODCTA))="6"
            SUMCTA[1,NUM]=SUMCTA[1,NUM]+DEBE-HABER
         ENDIF
         IF LTRIM(STR(CODCTA))="7"
            IF LTRIM(STR(CODCTA))<>"72" .AND. LTRIM(STR(CODCTA))<>"78"
               SUMCTA[2,NUM]=SUMCTA[2,NUM]-DEBE+HABER
            ENDIF
         ENDIF
      ENDIF
      SKIP
   ENDDO
   FOR N=1 TO 10
      SUMCTA[1,11]:=SUMCTA[1,11]+SUMCTA[1,N]
      SUMCTA[2,11]:=SUMCTA[2,11]+SUMCTA[2,N]
   NEXT


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

   PAG:=W_Imp1.T_Hoja.value
   LIN:=0

   oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
   oprint:printdata(20,190,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
   oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

   oprint:printdata(25,105,NOMEMPRESA,"times new roman",12,.F.,,"C",)
   oprint:printdata(32,105,TituloImp,"times new roman",18,.F.,,"C",)

   oprint:printdata(40,105,'Desde: '+DIA(W_Imp1.D_Fec1.Value,10)+" "+'Hasta: '+DIA(W_Imp1.D_Fec2.Value,10),"times new roman",12,.F.,,"C",)


   LIN:=60

   oprint:printdata(LIN, 20,"DEBE" ,"times new roman",18,.F.,,"L",)
   oprint:printdata(LIN,190,"HABER","times new roman",18,.F.,,"R",)
   oprint:printline(LIN+4,20,LIN+4,190,,0.5)

   FOR N1=1 TO 10
      LIN:=LIN+10
      oprint:printdata(LIN, 20,STR(aNOM6[N1,1],2)+" "+aNOM6[N1,2],"times new roman",10,.F.,,"L",)
      oprint:printdata(LIN,100,MIL(SUMCTA[1,N1],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
      DO CASE
      CASE N1=3 .AND. SUMCTA[2,N1]=0 //72
         SiImp:=0
      CASE N1=9 .AND. SUMCTA[2,N1]=0 //78
         SiImp:=0
      OTHERWISE
         SiImp:=1
      ENDCASE
      IF SiImp=1
         oprint:printdata(LIN,110,STR(aNOM7[N1,1],2)+" "+aNOM7[N1,2],"times new roman",10,.F.,,"L",)
         oprint:printdata(LIN,190,MIL(SUMCTA[2,N1],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
      ENDIF
   NEXT

   LIN:=LIN+5
   oprint:printline(LIN+4,20,LIN+4,190,,0.5)

   LIN:=LIN+10
   oprint:printdata(LIN, 20,"GANANCIAS","times new roman",12,.F.,,"L",)
   IF SUMCTA[1,11]<SUMCTA[2,11]
      oprint:printdata(LIN,100,MIL(SUMCTA[2,11]-SUMCTA[1,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,100,MIL(0,14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF
   oprint:printdata(LIN,110,"PERDIDAS","times new roman",12,.F.,,"L",)
   IF SUMCTA[1,11]>SUMCTA[2,11]
      oprint:printdata(LIN,190,MIL(SUMCTA[1,11]-SUMCTA[2,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,190,MIL(0,14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF

   LIN:=LIN+5
   oprint:printdata(LIN, 20,"(Antes de impuestos)","times new roman",12,.F.,,"L",)
   oprint:printdata(LIN,110,"(Antes de impuestos)","times new roman",12,.F.,,"L",)

   LIN:=LIN+5
   oprint:printline(LIN+4,20,LIN+4,190,,0.5)

   LIN:=LIN+10
   oprint:printdata(LIN, 20,"TOTAL "+MDA_NOM2(EJERANY),"times new roman",12,.F.,,"L",)
   IF SUMCTA[1,11]>=SUMCTA[2,11]
      oprint:printdata(LIN,100,MIL(SUMCTA[1,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,100,MIL(SUMCTA[2,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF
   oprint:printdata(LIN,110,"TOTAL "+MDA_NOM2(EJERANY),"times new roman",12,.F.,,"L",)
   IF SUMCTA[1,11]>=SUMCTA[2,11]
      oprint:printdata(LIN,190,MIL(SUMCTA[1,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,190,MIL(SUMCTA[2,11],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF

   LIN:=LIN+5
   oprint:printline(LIN+4,20,LIN+4,190,,0.5)

   ***LINEA HORIZONTAL***
   oprint:printline(55,105,LIN+4,105,,0.5)

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil




STATIC FUNCTION balpygi_mes()
   local oprint

   SUMCTA6:={}
   SUMCTA7:={}
   FOR N1=1 TO 10
      AADD(SUMCTA6,{0,0,0,0,0,0,0,0,0,0,0,0})
      AADD(SUMCTA7,{0,0,0,0,0,0,0,0,0,0,0,0})
   NEXT

   AbrirDBF("APUNTES")
   GO TOP
   W_Imp1.P_Progres.RangeMax:=LASTREC()
   W_Imp1.P_Progres.Value:=0
   DO WHILE .NOT. EOF()
      W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
      NUM:=VAL(SUBSTR(LTRIM(STR(CODCTA)),2,1))+1
      IF LTRIM(STR(CODCTA))="6"
         SUMCTA6[NUM,MONTH(FECHA)]:=SUMCTA6[NUM,MONTH(FECHA)]+DEBE-HABER
      ENDIF
      IF LTRIM(STR(CODCTA))="7"
         IF LTRIM(STR(CODCTA))<>"72" .AND. LTRIM(STR(CODCTA))<>"78"
            SUMCTA7[NUM,MONTH(FECHA)]:=SUMCTA7[NUM,MONTH(FECHA)]-DEBE+HABER
         ENDIF
      ENDIF
      SKIP
   ENDDO

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , .T. , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloImp)
   oprint:setpreviewsize(1)  // tamaño del preview
   oprint:beginpage()

   PAG:=W_Imp1.T_Hoja.value
   LIN:=0

   oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
   oprint:printdata(20,280,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
   oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

   oprint:printdata(25,150,NOMEMPRESA,"times new roman",12,.F.,,"C",)
   oprint:printdata(32,150,TituloImp,"times new roman",18,.F.,,"C",)


   LIN:=50
      oprint:printdata(LIN, 20,"DEBE","times new roman",14,.F.,,"L",)
      FOR N2=1 TO 12
         oprint:printdata(LIN,60+(N2*18),MES(N2),"times new roman",8,.F.,,"R",)
      NEXT

   oprint:printline(LIN+4,20,LIN+4,276,,0.5)
   LIN:=LIN+5

   FOR N1=1 TO LEN(SUMCTA6)
      oprint:printdata(LIN, 20,STR(aNOM6[N1,1],2)+" "+aNOM6[N1,2],"times new roman",8,.F.,,"L",)
      FOR N2=1 TO 12
         oprint:printdata(LIN,60+(N2*18),MIL(SUMCTA6[N1,N2],14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
      NEXT
      LIN:=LIN+5
   NEXT
   LIN2:=LIN

      LIN:=LIN+5
      oprint:printdata(LIN, 20,"HABER","times new roman",14,.F.,,"L",)
      FOR N2=1 TO 12
         oprint:printdata(LIN,60+(N2*18),MES(N2),"times new roman",8,.F.,,"R",)
      NEXT

   oprint:printline(LIN+4,20,LIN+4,276,,0.5)
   LIN:=LIN+5

   FOR N1=1 TO LEN(SUMCTA7)
      oprint:printdata(LIN, 20,STR(aNOM7[N1,1],2)+" "+aNOM7[N1,2],"times new roman",8,.F.,,"L",)
      FOR N2=1 TO 12
         DO CASE
         CASE N1=3 .AND. SUMCTA7[N1,N2]=0 //72
            SiImp:=0
         CASE N1=9 .AND. SUMCTA7[N1,N2]=0 //78
            SiImp:=0
         OTHERWISE
            SiImp:=1
         ENDCASE
         IF SiImp=1
            oprint:printdata(LIN,60+(N2*18),MIL(SUMCTA7[N1,N2],14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
         ENDIF
      NEXT
      LIN:=LIN+5
   NEXT

   LIN1:=LIN+5
   LIN2:=LIN+10
   oprint:printline(LIN1-1,20,LIN1-1,276,,0.5)
   oprint:printdata(LIN1, 20,"Perdidas (Antes de impuestos)","times new roman",8,.F.,,"L",)
   oprint:printdata(LIN2, 20,"Ganancias (Antes de impuestos)","times new roman",8,.F.,,"L",)
   oprint:printline(LIN2+4,20,LIN2+4,276,,0.5)
   FOR N2=1 TO 12
      IMPORTE2:=0
      FOR N1=1 TO LEN(SUMCTA6)
         IMPORTE2:=IMPORTE2-SUMCTA6[N1,N2]
      NEXT
      FOR N1=1 TO LEN(SUMCTA7)
         IMPORTE2:=IMPORTE2+SUMCTA7[N1,N2]
      NEXT
      IF IMPORTE2>0
         oprint:printdata(LIN1,60+(N2*18),MIL(       0,14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
         oprint:printdata(LIN2,60+(N2*18),MIL(IMPORTE2,14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
      ELSE
         oprint:printdata(LIN1,60+(N2*18),MIL(IMPORTE2*-1,14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
         oprint:printdata(LIN2,60+(N2*18),MIL(       0   ,14,MDA_DEC(EJERANY)) ,"times new roman",8,.F.,,"R",)
      ENDIF
   NEXT

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

   W_Imp1.release

Return Nil




