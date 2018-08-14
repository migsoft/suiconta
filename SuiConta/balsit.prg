#include "minigui.ch"
#include "winprint.ch"

procedure balsit()
   TituloImp:="Balance de Situacion"

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
               ACTION balsiti()

      @270,110 BUTTONEX B_Can CAPTION 'Cancelar' ICON icobus('cancelar') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil


STATIC FUNCTION balsiti()
   local oprint

   SUMCTA:=ARRAY(2,20)
   AFILL(SUMCTA[1],0)
   AFILL(SUMCTA[2],0)
   AbrirDBF("APUNTES")
   GO TOP
   W_Imp1.P_Progres.RangeMax:=LASTREC()
   W_Imp1.P_Progres.Value:=0
   DO WHILE .NOT. EOF()
      W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
      IF FECHA<W_Imp1.D_Fec1.Value .OR. FECHA>W_Imp1.D_Fec2.Value
         SKIP
         LOOP
      ENDIF
      CODCTA2:=STRZERO(CODCTA,8)
      DO CASE
* DEBE
      CASE CODCTA2="22" .OR. CODCTA2="23"
         SUMCTA[1,1]=SUMCTA[1,1]+DEBE-HABER
      CASE CODCTA2="21"
         SUMCTA[1,2]=SUMCTA[1,2]+DEBE-HABER
      CASE CODCTA2="24" .OR. CODCTA2="25"
         SUMCTA[1,3]=SUMCTA[1,3]+DEBE-HABER
      CASE CODCTA2="20" .OR. CODCTA2="27"
         SUMCTA[1,4]=SUMCTA[1,4]+DEBE-HABER
      CASE CODCTA2="26"
         SUMCTA[1,5]=SUMCTA[1,5]+DEBE-HABER
      CASE CODCTA2="28"
         SUMCTA[1,6]=SUMCTA[1,6]+DEBE-HABER
      CASE CODCTA2="29"
         SUMCTA[1,7]=SUMCTA[1,7]+DEBE-HABER
      CASE CODCTA2="3"
         SUMCTA[1,8]=SUMCTA[1,8]+DEBE-HABER
      CASE CODCTA2="43" .OR. CODCTA2="45"
         SUMCTA[1,9]=SUMCTA[1,9]+DEBE-HABER
      CASE CODCTA2="44" .OR. CODCTA2="46" .OR. CODCTA2="19"
         SUMCTA[1,10]=SUMCTA[1,10]+DEBE-HABER
      CASE CODCTA2="49"
         SUMCTA[1,11]=SUMCTA[1,11]+DEBE-HABER
      CASE CODCTA2="53" .OR. CODCTA2="54" .OR. CODCTA2="550" .OR. CODCTA2="558"
         SUMCTA[1,12]=SUMCTA[1,12]+DEBE-HABER
      CASE CODCTA2="57"
         SUMCTA[1,13]=SUMCTA[1,13]+DEBE-HABER
      CASE CODCTA2="565" .OR. CODCTA2="566"
         SUMCTA[1,14]=SUMCTA[1,14]+DEBE-HABER
      CASE CODCTA2="59"
         SUMCTA[1,15]=SUMCTA[1,15]+DEBE-HABER
      CASE CODCTA2="480" .OR. CODCTA2="580"
         SUMCTA[1,16]=SUMCTA[1,16]+DEBE-HABER
* HABER
      CASE CODCTA2="10"
         SUMCTA[2,1]=SUMCTA[2,1]-DEBE+HABER
      CASE CODCTA2="11" .OR. CODCTA2="120" .OR. CODCTA2="122" .OR. CODCTA2="129" .OR. CODCTA2="557"
         SUMCTA[2,2]=SUMCTA[2,2]-DEBE+HABER
      CASE CODCTA2="13"
         SUMCTA[2,3]=SUMCTA[2,3]-DEBE+HABER
      CASE CODCTA2="121"
         SUMCTA[2,4]=SUMCTA[2,4]-DEBE+HABER
      CASE CODCTA2="15" .OR. CODCTA2="16" .OR. CODCTA2="17"
         SUMCTA[2,5]=SUMCTA[2,5]-DEBE+HABER
      CASE CODCTA2="18"
         SUMCTA[2,6]=SUMCTA[2,6]-DEBE+HABER
      CASE CODCTA2="14"
         SUMCTA[2,7]=SUMCTA[2,7]-DEBE+HABER
      CASE CODCTA2="40" .OR. CODCTA2="42"
         SUMCTA[2,8]=SUMCTA[2,8]-DEBE+HABER
      CASE CODCTA2="41"
         SUMCTA[2,9]=SUMCTA[2,9]-DEBE+HABER
      CASE CODCTA2="47"
         SUMCTA[2,10]=SUMCTA[2,10]-DEBE+HABER
      CASE CODCTA2="50" .OR. CODCTA2="51" .OR. CODCTA2="52" .OR. CODCTA2="551";
      .OR. CODCTA2="552" .OR. CODCTA2="553" .OR. CODCTA2="555" .OR. CODCTA2="556"
         SUMCTA[2,11]=SUMCTA[2,11]-DEBE+HABER
      CASE CODCTA2="560" .OR. CODCTA2="561"
         SUMCTA[2,12]=SUMCTA[2,12]-DEBE+HABER
      CASE CODCTA2="485" .OR. CODCTA2="585"
         SUMCTA[2,13]=SUMCTA[2,13]-DEBE+HABER
      ENDCASE
      SKIP
   ENDDO

   FOR N=1 TO 19
      SUMCTA[1,20]=SUMCTA[1,20]+SUMCTA[1,N]
      SUMCTA[2,20]=SUMCTA[2,20]+SUMCTA[2,N]
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


   LIN1:=60
   LIN2:=60

   oprint:printdata(LIN1, 20,"ACTIVO" ,"times new roman",18,.F.,,"L",)
   oprint:printdata(LIN1,190,"PASIVO","times new roman",18,.F.,,"R",)
   oprint:printline(LIN1+4,20,LIN1+4,190,,0.5)

***ACTIVO***
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Inmovilizado material","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(22,23)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,1],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Inmovilizado inmaterial","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(21)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,2],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Inmovilizado financiero","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(24,25)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,3],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Gastos amortizables","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(20,27)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,4],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Fianzas y depos.const.largo","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(26)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,5],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Menos amortizaciones","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(28)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,6],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Menos provision inmoviliz.","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(29)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,7],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN1+4,30,LIN1+4,100,,0.5)
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 30,"INMOVILIZADO","times new roman",10,.F.,,"L",)
   NUM2=0
   FOR N=1 TO 7
      NUM2=NUM2+SUMCTA[1,N]
   NEXT
   oprint:printdata(LIN1,100,MIL(NUM2,14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN1:=LIN1+5

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Existencias","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(3)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,8],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN1+4,30,LIN1+4,100,,0.5)
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 30,"EXISTENCIAS","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,8],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN1:=LIN1+5

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Clientes","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(43,45)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,9],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Otros deudores","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(44,46,19)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,10],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Menos provis.insolven","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(49)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,11],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN1+4,30,LIN1+4,100,,0.5)
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 30,"DEUDORES","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,9]+SUMCTA[1,10]+SUMCTA[1,11],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN1:=LIN1+5

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Inversiones financieras","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(53,54,550,558)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,12],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Caja y bancos","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(57)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,13],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Fianzas y depos.const.corto","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(565,566)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,14],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Menos provis.financieras","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(59)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,15],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN1+4,30,LIN1+4,100,,0.5)
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 30,"ACTIV.FINANCIERA","times new roman",10,.F.,,"L",)
   NUM2=0
   FOR N=12 TO 15
      NUM2=NUM2+SUMCTA[1,N]
   NEXT
   oprint:printdata(LIN1,100,MIL(NUM2,14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN1:=LIN1+5

   LIN1:=LIN1+5
   oprint:printdata(LIN1, 20,"Gastos anticipados","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1, 70,"(480,580)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,16],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN1+4,30,LIN1+4,100,,0.5)
   LIN1:=LIN1+5
   oprint:printdata(LIN1, 30,"AJUSTES","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN1,100,MIL(SUMCTA[1,16],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN1:=LIN1+5


***PASIVO***
   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Capital","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(10)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,1],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Reservas y remanente","times new roman",10,.F.,,"L",)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,160,"(11,120,122,129,557)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,2],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Ingresos a distribuir","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(13)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,3],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Menos resultados anteriores","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(121)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,4],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN2+4,120,LIN2+4,190,,0.5)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,120,"FONDOS PROPIOS","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,1]+SUMCTA[2,2]+SUMCTA[2,3]+SUMCTA[2,4],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN2:=LIN2+5

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Acreedores largo plazo","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(15,16,17)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,5],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Fianzas y deposit.rec.largo","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(18)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,6],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Menos prov.riesgo y gastos","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(14)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,7],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN2+4,120,LIN2+4,190,,0.5)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,120,"ACREEDORES LARGO","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,5]+SUMCTA[2,6]+SUMCTA[2,7],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN2:=LIN2+5

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Proveedores","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(40,42)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,8],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Otros acreedores comerciales","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(41)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,9],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Deudas a entidades publicas","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(47)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,10],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Otras deudas a corto plazo","times new roman",10,.F.,,"L",)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,160,"(50,51,52,551,552,553,555,556)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,11],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Fianzas y depos.recib.corto","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(560,561)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,12],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN2+4,120,LIN2+4,190,,0.5)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,120,"ACREEDORES CORTO","times new roman",10,.F.,,"L",)
   NUM2=0
   FOR N=8 TO 12
      NUM2=NUM2+SUMCTA[2,N]
   NEXT
   oprint:printdata(LIN2,190,MIL(NUM2,14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN2:=LIN2+5


   LIN2:=LIN2+5
   oprint:printdata(LIN2,110,"Ingresos anticipados","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,160,"(485,585)","times new roman",8,.F.,,"R",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,13],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)

   oprint:printline(LIN2+4,120,LIN2+4,190,,0.5)
   LIN2:=LIN2+5
   oprint:printdata(LIN2,120,"AJUSTES","times new roman",10,.F.,,"L",)
   oprint:printdata(LIN2,190,MIL(SUMCTA[2,13],14,MDA_DEC(EJERANY)) ,"times new roman",10,.F.,,"R",)
   LIN2:=LIN2+5



   LIN:=MAX(LIN1,LIN2)
   LIN:=LIN+5
   oprint:printline(LIN+4,20,LIN+4,190,,0.5)

   LIN:=LIN+10
   oprint:printdata(LIN, 20,"PERDIDAS","times new roman",12,.F.,,"L",)
   IF SUMCTA[1,20]<SUMCTA[2,20]
      oprint:printdata(LIN,100,MIL(SUMCTA[2,20]-SUMCTA[1,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,100,MIL(0,14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF
   oprint:printdata(LIN,110,"GANANCIAS","times new roman",12,.F.,,"L",)
   IF SUMCTA[1,20]>SUMCTA[2,20]
      oprint:printdata(LIN,190,MIL(SUMCTA[1,20]-SUMCTA[2,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
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
   IF SUMCTA[1,20]>=SUMCTA[2,20]
      oprint:printdata(LIN,100,MIL(SUMCTA[1,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,100,MIL(SUMCTA[2,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ENDIF
   oprint:printdata(LIN,110,"TOTAL "+MDA_NOM2(EJERANY),"times new roman",12,.F.,,"L",)
   IF SUMCTA[1,20]>=SUMCTA[2,20]
      oprint:printdata(LIN,190,MIL(SUMCTA[1,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
   ELSE
      oprint:printdata(LIN,190,MIL(SUMCTA[2,20],14,MDA_DEC(EJERANY)) ,"times new roman",12,.F.,,"R",)
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




