#include "minigui.ch"
#include "winprint.ch"

procedure Faclvto()

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 480 ;
      TITLE 'Imprimir: Cuadro prevision de cobros' ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()


      @ 15,10 LABEL L_Tipo VALUE 'Tipo de listado' AUTOSIZE TRANSPARENT
      @ 10,110 RADIOGROUP R_Tipo OPTIONS { 'Cuadro de vencimientos' , 'Listado de vencimientos' } ;
               VALUE 1 WIDTH 160 ON CHANGE Faclvto_act() 

      @ 75,10 LABEL L_Vto1 VALUE 'Desde vencimiento' AUTOSIZE TRANSPARENT
      @ 70,120 DATEPICKER D_Vto1 WIDTH 100 VALUE DIAINIMES(DATE())
      @105,10 LABEL L_Vto2 VALUE 'Hasta vencimiento' AUTOSIZE TRANSPARENT
      @100,120 DATEPICKER D_Vto2 WIDTH 100 VALUE DIAINIMES(DATE())+365

      aFpago:={"  Todas las formas de pago"}
      FOR N=0 TO 8
         AADD(aFpago, LTRIM(STR(N))+"-"+VENCI_NC(N,30) )
      NEXT
      @135,10 LABEL L_Fpago VALUE 'Forma de Pago' AUTOSIZE TRANSPARENT
      @130,110 COMBOBOX C_Fpago WIDTH 250 ITEMS aFpago VALUE 1

      @165,10 LABEL L_SiVto VALUE 'Listar' AUTOSIZE TRANSPARENT
      @160,110 RADIOGROUP R_SiVto OPTIONS { 'Todas las facturas' , 'Solo con vencimiento' } ;
               VALUE 1 WIDTH 160

      @225,10 LABEL L_Orden VALUE 'Ordenado por' AUTOSIZE TRANSPARENT
      @220,110 RADIOGROUP R_Orden OPTIONS { 'Fecha vencimiento' , 'Nombre del cliente' } ;
               VALUE 1 WIDTH 160 

      @280,10 LABEL L_Progres VALUE 'Progreso' AUTOSIZE TRANSPARENT
      @280,110 PROGRESSBAR P_Progres RANGE 0 , 100 WIDTH 250 HEIGHT 20 SMOOTH

draw rectangle in window W_Imp1 at 320,010 to 322,390 fillcolor{255,0,0} //Rojo
      aIMP:=Impresoras(EMP_IMPRESORA)
      @335,10 LABEL L_Impresora VALUE 'Impresora' AUTOSIZE TRANSPARENT
      @330,100 COMBOBOX C_Impresora WIDTH 280 ;
            ITEMS aIMP[1] VALUE aIMP[3] NOTABSTOP

      @365,220 LABEL L_LibreriaImp VALUE 'Formato' AUTOSIZE TRANSPARENT
      @360,280 COMBOBOX C_LibreriaImp WIDTH 100 ;
            ITEMS aIMP[4] VALUE aIMP[5] NOTABSTOP

      @360, 10 CHECKBOX nImp CAPTION 'Seleccionar impresora' ;
               width 155 value .f. ;
               ON CHANGE W_Imp1.C_Impresora.Enabled:=IF(W_Imp1.nImp.Value=.T.,.F.,.T.)

      @390, 10 CHECKBOX nVer CAPTION 'Previsualizar documento' ;
               width 155 value .f.

      @420, 10 BUTTONEX B_Imp CAPTION 'Imprimir' ICON icobus('imprimir') WIDTH 90 HEIGHT 25 ;
               ACTION Faclvtoi1()

      @420,110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

      END WINDOW
      VentanaCentrar("W_Imp1","Ventana1")
      ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION Faclvto_act()
IF W_Imp1.R_Tipo.value=1
   W_Imp1.D_Vto1.value:=DIAINIMES(DATE())
   W_Imp1.D_Vto2.value:=DIAFINMES(DATE())+365
ELSE
   W_Imp1.D_Vto1.value:=DIAINIMES(DIAMESMAS(DATE(),1))
   W_Imp1.D_Vto2.value:=DIAFINMES(DIAMESMAS(DATE(),1))
ENDIF


STATIC FUNCTION Faclvtoi1()

   IF FILE("FIN.DBF")
      AbrirDBF("Fin","SIN_INDICE")
      FIN->( DBCLOSEAREA() )
      ERASE FIN.DBF
      ERASE FIN.CDX
   ENDIF

FICBASE:={{'SERFAC      ','C',         1,         0},;
          {'NFAC        ','N',         6,         0},;
          {'FFAC        ','D',        10,         0},;
          {'COD         ','N',         5,         0},;
          {'CLIENTE     ','C',        30,         0},;
          {'TFAC        ','N',        14,         3},;
          {'TVTO        ','N',        14,         3},;
          {'FVTO        ','D',        10,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'DPAGO       ','C',         4,         0},;
          {'PEND        ','L',         1,         0},;
          {'ANYFAC      ','N',         4,         0}}
DBCREATE("FIN.DBF",FICBASE)
AbrirDBF("FIN","SIN_INDICE",,"Exclusive")


*** INCLUIR EJERCICIOS ANTORIORES ***
FOR ANY2=EJERANY-1 TO EJERANY+1
   W_Imp1.L_Progres.Value:=STR(ANY2,4)
   RUTA2:=EMPRUTANY(NUMEMP,ANY2)
   IF FILE(RUTA2+"\FAC92.DBF") .AND. ;
      FILE(RUTA2+"\FAC92.CDX")
      AbrirDBF("FAC92",,,,RUTA2)
      W_Imp1.P_Progres.RangeMax:=LASTREC()
      W_Imp1.P_Progres.Value:=0
      GO TOP
      DO WHILE .NOT. EOF()
         W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
			DO EVENTS
         SFAC2:=SERFAC
         NFAC2:=NFAC
         FFAC2:=FFAC
         COD2:=COD
         CLIENTE2:=CLIENTE
         TFAC2:=TFAC
         PEND2:=PEND
         ACTA2:=ACTA
         FPAGO2:=FPAGO
         DPAGO2:=DPAGO
         IMPVTO2:=0

         NVTOS:=VAL(SUBSTR(LTRIM(STR(FPAGO)),2,1))
         DIAVTO1=VAL(LEFT(DPAGO,2))
         DIAVTO2=VAL(RIGHT(DPAGO,2))

         IF NVTOS=0 .AND. FFAC>=W_Imp1.D_Vto1.value .AND. FFAC<=W_Imp1.D_Vto2.value
               SELEC FIN
               APPEND BLANK
               REPLACE SERFAC WITH SFAC2
               REPLACE NFAC WITH NFAC2
               REPLACE FFAC WITH FFAC2
               REPLACE COD WITH COD2
               REPLACE CLIENTE WITH CLIENTE2
               REPLACE TFAC WITH TFAC2
               REPLACE TVTO WITH TFAC2-ACTA2
               REPLACE FVTO WITH FFAC2
               REPLACE FPAGO WITH FPAGO2
               REPLACE DPAGO WITH DPAGO2
               REPLACE ANYFAC WITH ANY2
               IF TFAC2>0
                  REPLACE PEND WITH IF(PEND2>0, .T. , .F. )
               ELSE
                  REPLACE PEND WITH IF(PEND2<0, .T. , .F. )
               ENDIF
               SELEC FAC92
         ENDIF
         IMPVTO2:=IMPVTO2+ACTA
         IF ACTA<>0 .AND. FFAC>=W_Imp1.D_Vto1.value .AND. FFAC<=W_Imp1.D_Vto2.value
               SELEC FIN
               APPEND BLANK
               REPLACE SERFAC WITH SFAC2
               REPLACE NFAC WITH NFAC2
               REPLACE FFAC WITH FFAC2
               REPLACE COD WITH COD2
               REPLACE CLIENTE WITH CLIENTE2
               REPLACE TFAC WITH TFAC2
               REPLACE TVTO WITH ACTA2
               REPLACE FVTO WITH FFAC2
               REPLACE FPAGO WITH 0
               REPLACE DPAGO WITH DPAGO2
               REPLACE ANYFAC WITH ANY2
               IF TFAC2>0
                  REPLACE PEND WITH IF(TFAC2-PEND2>=IMPVTO2, .F. , .T. )
               ELSE
                  REPLACE PEND WITH IF(TFAC2-PEND2<=IMPVTO2, .F. , .T. )
               ENDIF
               SELEC FAC92
         ENDIF

         FOR NVTO=1 TO NVTOS
            FVTO2:=VENCI(2,FPAGO,FFAC,,DIAVTO1,DIAVTO2,NVTO)
            TVTO2:=VENCI(3,FPAGO,,TFAC-ACTA,,,NVTO)
            IMPVTO2:=IMPVTO2+TVTO2
            IF FVTO2>=W_Imp1.D_Vto1.value  .AND. FVTO2<=W_Imp1.D_Vto2.value
               SELEC FIN
               APPEND BLANK
               REPLACE SERFAC WITH SFAC2
               REPLACE NFAC WITH NFAC2
               REPLACE FFAC WITH FFAC2
               REPLACE COD WITH COD2
               REPLACE CLIENTE WITH CLIENTE2
               REPLACE TFAC WITH TFAC2
               REPLACE TVTO WITH TVTO2
               REPLACE FVTO WITH FVTO2
               REPLACE FPAGO WITH FPAGO2
               REPLACE DPAGO WITH DPAGO2
               REPLACE ANYFAC WITH ANY2
               IF TFAC2>0
                  REPLACE PEND WITH IF(TFAC2-PEND2>=IMPVTO2, .F. , .T. )
               ELSE
                  REPLACE PEND WITH IF(TFAC2-PEND2<=IMPVTO2, .F. , .T. )
               ENDIF
               SELEC FAC92
            ENDIF
         NEXT
         SKIP
      ENDDO
   ENDIF
NEXT
*** FINAL INCLUIR EJERCICIOS ANTORIORES ***

   SELEC FIN

   IF W_Imp1.C_Fpago.value>1
      DELETE FOR VAL(LEFT(LTRIM(STR(FPAGO)),1))<>W_Imp1.C_Fpago.value-2
      PACK
   ENDIF

   IF W_Imp1.R_SiVto.value=2
      DELETE FOR VAL(SUBSTR(STR(FPAGO),2,1))=0
      PACK
   ENDIF

   IF W_Imp1.R_Orden.value=1
      INDEX ON DTOS(FVTO)+UPPER(CLIENTE)+SERFAC+STR(NFAC) TO FIN
   ELSE
      INDEX ON UPPER(CLIENTE)+DTOS(FVTO)+SERFAC+STR(NFAC) TO FIN
   ENDIF

   GO TOP
   IF LASTREC()=0
      MsgExclamation("No hay datos en los parametros introducidos","Informacion")
      RETURN
   ENDIF

   Faclvtoi2()





STATIC FUNCTION Faclvtoi2()
   local oprint

   oprint:=tprint(UPPER(W_Imp1.C_LibreriaImp.DisplayValue))
   oprint:init()
   oprint:setunits("MM",5)
   HOJAHORZ2:=IF(W_Imp1.R_Tipo.value=1,.T.,.F.)
   oprint:selprinter(W_Imp1.nImp.value , W_Imp1.nVer.value , HOJAHORZ2 , 9 , W_Imp1.C_Impresora.DisplayValue)
   if oprint:lprerror
      oprint:release()
      return nil
   endif
   oprint:begindoc(TituloqImp(W_Imp1.Title))
   oprint:setpreviewsize(W_Imp1.R_Tipo.value)  // tamaño del preview
   oprint:beginpage()

W_Imp1.P_Progres.RangeMax:=LASTREC()
W_Imp1.P_Progres.Value:=0
aFECVTO:={DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,0)) , ;
          DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,1)) , ;
          DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,2)) , ;
          DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,3)) , ;
          DIAINIMES(DIAMESMAS(W_Imp1.D_Vto1.value,4)) }
PAG:=0
LIN:=0
TOT1:=0
TOT2:=0
TOT3:=0
TOT4:=0
TOT5:=0
aCOL:={}
   IF W_Imp1.R_Tipo.value=1
AADD(aCOL, 10) //tamaño letra
AADD(aCOL,297) //ancho pagina
AADD(aCOL,210) //alto pagina
AADD(aCOL, 35)
AADD(aCOL, 37)
AADD(aCOL, 65)
AADD(aCOL, 67)
AADD(aCOL,130)
AADD(aCOL,170)
AADD(aCOL,190)
AADD(aCOL,210)
AADD(aCOL,230)
AADD(aCOL,250)
AADD(aCOL,252)
AADD(aCOL,270)
   ELSE
AADD(aCOL, 10) //tamaño letra
AADD(aCOL,210) //ancho pagina
AADD(aCOL,297) //alto pagina
AADD(aCOL, 35)
AADD(aCOL, 37)
AADD(aCOL, 65)
AADD(aCOL, 67)
AADD(aCOL,130)
AADD(aCOL,170)
AADD(aCOL, 10)
AADD(aCOL, 10)
AADD(aCOL, 10)
AADD(aCOL, 10)
AADD(aCOL,172)
AADD(aCOL,190)
   ENDIF
DO WHILE .NOT. EOF()
   W_Imp1.P_Progres.Value:=W_Imp1.P_Progres.Value+1
   DO EVENTS

   IF LIN>=aCOL[3]-35 .OR. PAG=0 .OR. SALTA2=1
      SALTA2:=0
      IF PAG<>0
         oprint:printline(LIN-1,aCOL[9]-40,LIN-1,aCOL[15],,0.5)
         oprint:printdata(LIN,aCOL[9]-40,"Sumas","times new roman",aCOL[1],.F.,,"L",)
         oprint:printdata(LIN,aCOL[9],MIL(TOT1,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
         IF W_Imp1.R_Tipo.value=1
            oprint:printdata(LIN,aCOL[10],MIL(TOT2,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[11],MIL(TOT3,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[12],MIL(TOT4,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[13],MIL(TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
            oprint:printdata(LIN,aCOL[15],MIL(TOT1+TOT2+TOT3+TOT4+TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
         ENDIF
         LIN:=LIN+10
         oprint:printdata(LIN,aCOL[2]/2,"Sigue en la hoja: "+LTRIM(STR(PAG+1)),"times new roman",10,.F.,,"C",)
         oprint:endpage()
         oprint:beginpage()
      ENDIF
      PAG=PAG+1

      oprint:printdata(20,20,NOMPROGRAMA,"times new roman",12,.F.,,"L",)
      oprint:printdata(20,aCOL[2]-20,"Hoja: "+LTRIM(STR(PAG)),"times new roman",12,.F.,,"R",)
      oprint:printdata(25,20,DIA(DATE(),10),"times new roman",12,.F.,,"L",)

      oprint:printdata(25,aCOL[2]/2,NOMEMPRESA,"times new roman",12,.F.,,"C",)
      oprint:printdata(32,aCOL[2]/2,TituloqImp(W_Imp1.Title),"times new roman",18,.F.,,"C",)

      oprint:printdata(40,20,'Desde fecha: '+DIA(W_Imp1.D_Vto1.value,10),"times new roman",12,.F.,,"L",)
      oprint:printdata(45,20,'Hasta fecha: '+DIA(W_Imp1.D_Vto2.value,10),"times new roman",12,.F.,,"L",)

      LIN:=55
      oprint:printdata(LIN,aCOL[4],"Factura","times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[5],"Fecha","times new roman",aCOL[1],.F.,,"L",)
      oprint:printdata(LIN,aCOL[6],"Codigo","times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[7],"Cliente","times new roman",aCOL[1],.F.,,"L",)
      oprint:printdata(LIN,aCOL[8],"Fecha vto.","times new roman",aCOL[1],.F.,,"L",)

      IF W_Imp1.R_Tipo.value=1
         oprint:printdata(LIN,aCOL[09],MES(MONTH(aFECVTO[1])),"times new roman",aCOL[1],.F.,,"R",)
         oprint:printdata(LIN,aCOL[10],MES(MONTH(aFECVTO[2])),"times new roman",aCOL[1],.F.,,"R",)
         oprint:printdata(LIN,aCOL[11],MES(MONTH(aFECVTO[3])),"times new roman",aCOL[1],.F.,,"R",)
         oprint:printdata(LIN,aCOL[12],MES(MONTH(aFECVTO[4])),"times new roman",aCOL[1],.F.,,"R",)
         oprint:printdata(LIN,aCOL[13],"Posterior","times new roman",aCOL[1],.F.,,"R",)
      ELSE
         oprint:printdata(LIN,aCOL[9],"Importe","times new roman",aCOL[1],.F.,,"R",)
      ENDIF

      oprint:printdata(LIN,aCOL[14],"F.Cobro","times new roman",aCOL[1],.F.,,"L",)
      oprint:printline(LIN+4,15,LIN+4,aCOL[15],,0.5)

      LIN:=LIN+5
   ENDIF

   oprint:printdata(LIN,aCOL[4],LTRIM(STR(NFAC))+"-"+SERFAC,"times new roman",aCOL[1],.F.,,"R",)
   oprint:printdata(LIN,aCOL[5],DIA(FFAC,8),"times new roman",aCOL[1],.F.,,"L",)
   oprint:printdata(LIN,aCOL[6],COD,"times new roman",aCOL[1],.F.,,"R",)
   oprint:printdata(LIN,aCOL[7],CLIENTE,"times new roman",aCOL[1],.F.,,"L",)
   oprint:printdata(LIN,aCOL[8],DIA(FVTO,8),"times new roman",aCOL[1],.F.,,"L",)
   DO CASE
   CASE FVTO<aFECVTO[2] .OR. W_Imp1.R_Tipo.value=2
      oprint:printdata(LIN,aCOL[9],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[9],IF(PEND=.T.,"","c"),"times new roman",aCOL[1],.F.,,"L",)
      TOT1=TOT1+TVTO
   CASE FVTO>=aFECVTO[2] .AND. FVTO<aFECVTO[3]
      oprint:printdata(LIN,aCOL[10],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[10],IF(PEND=.T.,"","c"),"times new roman",aCOL[1],.F.,,"L",)
      TOT2=TOT2+TVTO
   CASE FVTO>=aFECVTO[3] .AND. FVTO<aFECVTO[4]
      oprint:printdata(LIN,aCOL[11],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[11],IF(PEND=.T.,"","c"),"times new roman",aCOL[1],.F.,,"L",)
      TOT3=TOT3+TVTO
   CASE FVTO>=aFECVTO[4] .AND. FVTO<aFECVTO[5]
      oprint:printdata(LIN,aCOL[12],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[12],IF(PEND=.T.,"","c"),"times new roman",aCOL[1],.F.,,"L",)
      TOT4=TOT4+TVTO
   OTHERWISE
      oprint:printdata(LIN,aCOL[13],MIL(TVTO,11,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[13],IF(PEND=.T.,"","c"),"times new roman",aCOL[1],.F.,,"L",)
      TOT5=TOT5+TVTO
   ENDCASE

   oprint:printdata(LIN,aCOL[14],VENCI_NC( VAL(LEFT(LTRIM(STR(FPAGO)),1)) ,8),"times new roman",aCOL[1],.F.,,"L",)

   LIN:=LIN+5
   SKIP

ENDDO

   oprint:printline(LIN-1,aCOL[9]-40,LIN-1,aCOL[15],,0.5)
   oprint:printdata(LIN,aCOL[9]-40,"Total","times new roman",aCOL[1],.F.,,"L",)
   oprint:printdata(LIN,aCOL[9],MIL(TOT1,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
   IF W_Imp1.R_Tipo.value=1
      oprint:printdata(LIN,aCOL[10],MIL(TOT2,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[11],MIL(TOT3,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[12],MIL(TOT4,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[13],MIL(TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
      oprint:printdata(LIN,aCOL[15],MIL(TOT1+TOT2+TOT3+TOT4+TOT5,14,MDA_DEC(EJERANY)),"times new roman",aCOL[1],.F.,,"R",)
   ENDIF

   SELEC FIN
   FIN->( DBCLOSEAREA() )

   oprint:endpage()
   oprint:enddoc()
   oprint:RELEASE()

Return Nil




