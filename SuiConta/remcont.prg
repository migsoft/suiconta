#include "minigui.ch"

procedure RemCont()

   RUTAREMESA:=RUTAEMPRESA
   RUTAREMEMP:=RUTAPROGRAMA

   aREMESAS:=Remesa_alis(RUTAREMESA)

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 900 HEIGHT 480 ;
      TITLE 'Contabilizar remesas' ;
      MODAL      ;
      NOSIZE     ;
      ON RELEASE CloseTables()

      @10,10 GRID GR_Remesas ;
      HEIGHT 250 ;
      WIDTH 460 ;
      HEADERS {'Serie','Numero','Remesa','Fecha','Banco','Importe','Documentos','Asiento'} ;
      WIDTHS {0,0,75,80,75,100,50,50 } ;
      ITEMS aREMESAS ;
      VALUE 1 ;
      COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999','E'},{'TEXTBOX','NUMERIC'}} ;
      ON HEADCLICK {{|| Remesa_Grid_Ord(1)},{|| Remesa_Grid_Ord(2)},{|| Remesa_Grid_Ord(3)}, ;
                    {|| Remesa_Grid_Ord(4)},{|| Remesa_Grid_Ord(5)},{|| Remesa_Grid_Ord(6)}, ;
                    {|| Remesa_Grid_Ord(7)},{|| Remesa_Grid_Ord(8)}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
               BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      ON CHANGE RemesaActRec()

      @10,480 GRID GR_Recibos ;
      HEIGHT 250 ;
      WIDTH 410 ;
      HEADERS {'Factura','Cliente','Importe','Vencimiento'} ;
      WIDTHS {50,150,100,80 } ;
      ITEMS {} ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'},{'TEXTBOX','DATE'}} ;
      ON HEADCLICK {{|| Remesa_Grid_RecOrd(1)},{|| Remesa_Grid_RecOrd(2)},{|| Remesa_Grid_RecOrd(3)},{|| Remesa_Grid_RecOrd(4)}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER}

      @350,010 TEXTBOX T_ImpCtaGas WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
      @350,120 BUTTONEX Bt_CodCtaGas CAPTION 'Cuenta gasto' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
         ACTION ( br_cue1(VAL(W_Imp1.T_CodCtaGas.Value),"W_Imp1","T_CodCtaGas","L_CodCtaGas") ) NOTABSTOP
      @350,240 TEXTBOX T_CodCtaGas WIDTH 100 VALUE "66430001" MAXLENGTH 8 ;
             ON LOSTFOCUS ( W_Imp1.T_CodCtaGas.Value:=PCODCTA3(W_Imp1.T_CodCtaGas.Value) , ;
             W_Imp1.L_CodCtaGas.Value:=Codigos_NomCta(W_Imp1.T_CodCtaGas.Value) )
      @355,350 LABEL L_CodCtaGas VALUE Codigos_NomCta(W_Imp1.T_CodCtaGas.Value) AUTOSIZE TRANSPARENT



      @410, 10 BUTTONEX B_Conta CAPTION 'Contabilizar' WIDTH 90 HEIGHT 25 ;
               ACTION RemCont1()

      @410,110 BUTTONEX B_Cobros CAPTION 'Altas cobros' WIDTH 90 HEIGHT 25 ;
               ACTION RemCont2()

      @410,210 BUTTONEX B_Pagos CAPTION 'Altas pagos' WIDTH 90 HEIGHT 25 ;
               ACTION RemCont3()

      @410,310 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

   RemesaActRec()

   END WINDOW
   VentanaCentrar("W_Imp1","Ventana1")
   ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION RemCont1()
   NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
   NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
   CODBAN2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,5)
   TOTREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,6)
   REMTIT:="Remesa "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)

   IF MSGYESNO("¿Desea contabilizar la "+REMTIT+"?")=.F.
      RETURN
   ENDIF

   IF W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,8)<>0
      IF MSGYESNO("La remesa esta contabilizada"+HB_OsNewLine()+ ;
                  "¿Desea seguir con el proceso?","Atencion")=.F.
         RETURN
      ELSE
         PonerEspera("Eliminando asiento remesa "+REMTIT)
         NASI2=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,8)
         AbrirDBF("APUNTES",,,,,1)
         SEEK NASI2
         DO WHILE NASI=NASI2 .AND. .NOT. EOF()
            IF RLOCK()
               DELETE
               DBCOMMIT()
               DBUNLOCK()
               Suizo_saldocuenta(CODCTA,DEBE,HABER)
            ENDIF
            SKIP
         ENDDO
         QuitarEspera()
      ENDIF
   ELSE
      AbrirDBF("APUNTES")
      GO BOTT
      NASI2=IF(NASI<2,2,NASI+1)
   ENDIF
   NAPU2:=1

PonerEspera("Contabilizando remesa "+REMTIT)

   aAsiento:={}


   AbrirDBF("REMESA",,,,RUTAREMESA)
   SEEK (NSER2*100000)+NREM2
   DEBE2 :=IF(NSER2>=5,0,TOTREM2-W_Imp1.T_ImpCtaGas.Value)
   HABER2:=IF(NSER2>=5,TOTREM2+W_Imp1.T_ImpCtaGas.Value,0)

   ***CONTABILIZAR POR FECHA DE VENCIMIENTO***
   FAPU2:=FREM
   DO WHILE NREM=NREM2 .AND. SERIE=NSER2 .AND. .NOT. EOF()
      FAPU2:=MAX(FAPU2,FVTO)
      SKIP
   ENDDO
   ***FIN CONTABILIZAR POR FECHA DE VENCIMIENTO***

   SEEK (NSER2*100000)+NREM2
   AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,FAPU2,CODBAN2,REMTIT,DEBE2,HABER2})
   IF W_Imp1.T_ImpCtaGas.Value<>0
      IF VAL(W_Imp1.T_CodCtaGas.Value)=0
         W_Imp1.T_CodCtaGas.Value:="66430001"
         W_Imp1.L_CodCtaGas.Value:=Codigos_NomCta(W_Imp1.T_CodCtaGas.Value)
      ENDIF
      AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,FAPU2,VAL(W_Imp1.T_CodCtaGas.Value),REMTIT,W_Imp1.T_ImpCtaGas.Value,0})
   ENDIF
   DO WHILE NREM=NREM2 .AND. SERIE=NSER2 .AND. .NOT. EOF()
      DO CASE
      CASE NSER2>=1 .AND. NSER2<=2
         REMTIT:="Cobro giro fac."+LTRIM(NFRA)
      CASE NSER2>=3 .AND. NSER2<=4
         REMTIT:="Cobro talon fac."+LTRIM(NFRA)
      CASE NSER2>=5 .AND. NSER2<=6
         REMTIT:="Pago transf.fac."+LTRIM(NFRA)
      ENDCASE
      DEBE2 :=IF(NSER2>=5,IMPORTE,0)
      HABER2:=IF(NSER2>=5,0,IMPORTE)
      AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,FAPU2,CODCTA,REMTIT,DEBE2,HABER2})
      IF RLOCK()
         REPLACE NASI WITH NASI2
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
      SKIP
   ENDDO

   AbrirDBF("CUENTAS")
   AbrirDBF("APUNTES")
   FOR N=1 TO LEN(aAsiento)
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI   WITH aAsiento[N,1]
         REPLACE APU    WITH aAsiento[N,2]
         REPLACE NEMP   WITH aAsiento[N,3]
         REPLACE FECHA  WITH aAsiento[N,4]
         REPLACE CODCTA WITH aAsiento[N,5]
         REPLACE NOMAPU WITH aAsiento[N,6]
         REPLACE DEBE   WITH aAsiento[N,7]
         REPLACE HABER  WITH aAsiento[N,8]
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   NEXT

   W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,8):=NASI2

QuitarEspera()
MsgInfo("La remesa se ha contabilizado correctamente"+HB_OsNewLine()+"Asiento "+LTRIM(STR(NASI2)))




STATIC FUNCTION RemCont2()
   NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
   NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
   CODBAN2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,5)
   REMTIT:="Remesa "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)
   ALTCOB2:=0

   IF MSGYESNO("¿Desea dar de alta los cobros de la "+REMTIT+"?")=.F.
      RETURN
   ENDIF

PonerEspera("Contabilizando cobros facturas remesa "+REMTIT)

   AbrirDBF("REMESA",,,,RUTAREMESA)
   SEEK (NSER2*100000)+NREM2
   DO WHILE NREM=NREM2 .AND. SERIE=NSER2 .AND. .NOT. EOF()
      CODCTA2:=CODCTA
      FRACOB2:=RTRIM(NFRA)
      IMP2:=IMPORTE
      FCOB2:=FREM
      FVTO2:=FVTO
      NASI2:=NASI
      NEMPREG2:=NEMPREG
      NREG2:=NREG

      NFAC2:=""
      SFAC2:=""
      RUTA1:={RUTAEMPRESA}
      IF NEMPREG2<>0 .AND. NREG2<>0
         NFAC2:=NREG2
         FOR N=1 TO LEN(RTRIM(FRACOB2))
            IF ISALPHA(SUBSTR(FRACOB2,N,1))=.T.
               SFAC2:=SUBSTR(FRACOB2,N,1)
               EXIT
            ENDIF
         NEXT
         IF NUMEMP<>NEMPREG2
            RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,NEMPREG2,"SUICONTA")
         ENDIF
         AbrirDBF("FAC92",,,,RUTA1[1],1)
         SEEK SERIE(SFAC2,NFAC2)
      ELSE
         ***SEPARACION "," Y "->"***
         DO CASE
         CASE RAT(",",FRACOB2)>RAT(">",FRACOB2)
            FRACOB2:=SUBSTR(FRACOB2,RAT(",",FRACOB2)+1,LEN(FRACOB2))
         CASE RAT(",",FRACOB2)<RAT(">",FRACOB2)
            FRACOB2:=SUBSTR(FRACOB2,RAT(">",FRACOB2)+1,LEN(FRACOB2))
         ENDCASE
         ***FIN SEPARACION "," Y "->"***

         IF AT("/",FRACOB2)<>0
            FRACOB2:=SUBSTR(FRACOB2,1,RAT("/",FRACOB2)-1)
         ENDIF

         NFAC2:=""
         FOR N=1 TO LEN(FRACOB2)
            IF ASC(SUBSTR(FRACOB2,N,1))>=48 .AND. ASC(SUBSTR(FRACOB2,N,1))<=57
               NFAC2:=NFAC2+SUBSTR(FRACOB2,N,1)
            ELSE
               EXIT
            ENDIF
         NEXT
         NFAC2:=VAL(NFAC2)
         IF NFAC2=0
            SKIP
            LOOP
         ENDIF

         SFAC2:=" "
         FOR N=LEN(FRACOB2) TO 1 STEP -1
            IF ASC(SUBSTR(FRACOB2,N,1))>=65 .AND. ASC(SUBSTR(FRACOB2,N,1))<=90
               SFAC2:=SUBSTR(FRACOB2,N,1)
               EXIT
            ENDIF
         NEXT
         IF SFAC2=" "
            SFAC2:="A"
         ENDIF

         AbrirDBF("FAC92")
         SEEK SERIE(SFAC2,NFAC2)
      ENDIF

      IF CODCTA=CODCTA2 .AND. .NOT. EOF()
         IF RLOCK()
            REPLACE PEND WITH PEND-IMP2
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
         FFAC2:=FFAC

         AbrirDBF("COBROS",,,,RUTA1[1],1)
         ALTCOB2++
         APPEND BLANK
         IF RLOCK()
            REPLACE SERFAC WITH SFAC2
            REPLACE NFAC WITH NFAC2
            REPLACE FFAC WITH FFAC2
            REPLACE IMPORTE WITH IMP2
            REPLACE FCOB WITH FCOB2
            REPLACE DESCRIP WITH REMTIT
            REPLACE FVTO WITH FVTO2
            REPLACE BANCO WITH CODBAN2
            REPLACE NEMP WITH NEMPREG2
            REPLACE NASI WITH NASI2
            REPLACE FASI WITH FCOB2
            REPLACE NEMPASI WITH NUMEMP
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
      ENDIF

      AbrirDBF("REMESA",,,,RUTAREMESA)
      SKIP

   ENDDO


QuitarEspera()
MsgInfo("Los cobros han sido dados de alta"+HB_OsNewLine()+"Total cobros "+LTRIM(STR(ALTCOB2)))




STATIC FUNCTION RemCont3()
   NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
   NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
   CODBAN2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,5)
   REMTIT:="Remesa "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)
   ALTPAG2:=0

   IF MSGYESNO("¿Desea dar de alta los pagos de la "+REMTIT+"?")=.F.
      RETURN
   ENDIF

PonerEspera("Contabilizando pagos facturas remesa "+REMTIT)

   AbrirDBF("REMESA",,,,RUTAREMESA)
   SEEK (NSER2*100000)+NREM2
   DO WHILE NREM=NREM2 .AND. SERIE=NSER2 .AND. .NOT. EOF()
      CODCTA2:=CODCTA
      FRAPAG2:=RTRIM(NFRA)
      IMP2:=IMPORTE
      FPAG2:=FREM
      FVTO2:=FVTO
      NASI2:=NASI
      NEMPREG2:=NEMPREG
      NREG2:=NREG

      RUTA1:={RUTAEMPRESA}
      IF NEMPREG2<>0 .AND. NREG2<>0
         IF NUMEMP<>NEMPREG2
            RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,NEMPREG2,"SUICONTA")
         ENDIF
         AbrirDBF("FACREB",,,,RUTA1[1],1)
         SEEK NREG2
      ELSE
         AbrirDBF("FACREB")
         GO TOP
         NREG2:=0
         DO WHILE .NOT. EOF()
            IF CODIGO=CODCTA2 .AND. REF=FRAPAG2
               NREG2:=NREG
               EXIT
            ENDIF
            SKIP
         ENDDO
      ENDIF

      IF CODIGO=CODCTA2 .AND. .NOT. EOF()
         IF RLOCK()
            REPLACE PEND WITH PEND-IMP2
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
         FREG2:=FREG

         AbrirDBF("PAGOS",,,,RUTA1[1],1)
         ALTPAG2++
         APPEND BLANK
         IF RLOCK()
            REPLACE NREG WITH NREG2
            REPLACE FREG WITH FREG2
            REPLACE IMPORTE WITH IMP2
            REPLACE FPAG WITH FPAG2
            REPLACE DESCRIP WITH REMTIT
            REPLACE FVTO WITH FVTO2
            REPLACE BANCO WITH CODBAN2
            REPLACE NEMP WITH NEMPREG2
            REPLACE NASI WITH NASI2
            REPLACE FASI WITH FPAG2
            REPLACE NEMPASI WITH NUMEMP
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
      ENDIF

      AbrirDBF("REMESA",,,,RUTAREMESA)
      SKIP

   ENDDO


QuitarEspera()
MsgInfo("Los cobros han sido dados de alta"+HB_OsNewLine()+"Total cobros "+LTRIM(STR(ALTPAG2)))

