#include "minigui.ch"

Function Nomina_Conta()

   IF IsWindowDefined(WinNomCont1)=.T.
      DECLARE WINDOW WinNomCont1
      WinNomCont1.restore
      WinNomCont1.Show
      WinNomCont1.SetFocus
      RETURN
   ENDIF

IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
   RETURN
ENDIF


   aBANCO:={}
IF UPPER(PROGRAMA())="SUICONTA"
   AADD(aBANCO,{NOMEMPRESA,RUTAEMPRESA,LTRIM(STR(NUMEMP))+"-"+NOMEMPRESA} )
ELSE
   IF FILE("BANCOT.DBF") .AND. FILE("BANCOT.CDX")
      AbrirDBF("bancot")
      GO TOP
      DO WHILE .NOT. EOF()
         IF BANCOCAJA=1
            AADD(aBANCO,{LTRIM(STR(CODBAN))+"-"+RTRIM(NOMBAN)+" "+RTRIM(NOMTAL),RTRIM(RUTCONTA),EMPCONTA} )
         ENDIF
         SKIP
      ENDDO
   ENDIF
ENDIF
   IF LEN(aBANCO)=0
      AADD(aBANCO,{" "," "," "})
   ENDIF

   DEFINE WINDOW WinNomCont1 ;
      AT 50,0     ;
      WIDTH 1024  ;
      HEIGHT 470 ;
      TITLE "Apuntes" ;
      CHILD NOMAXIMIZE ;
      NOSIZE BACKCOLOR MiColor("ARENA") ;
      ON RELEASE CloseTables()

      @010,110 COMBOBOX C_Mes WIDTH 100 ITEMS MES(,1) VALUE MONTH(DATE()-30)
      @010,10 BUTTONEX Bt_BuscarMes ;
         CAPTION 'Nominas mes' ICON icobus('buscar') ;
         ACTION Actualiz_Nomi1("ASIENTO") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @015,230 LABEL L_Fec1 VALUE 'Nomina' AUTOSIZE TRANSPARENT
      @010,285 DATEPICKER D_Fec1 WIDTH 90 VALUE DATE()
      @015,395 LABEL L_Fec2 VALUE 'Pago' AUTOSIZE TRANSPARENT
      @010,430 DATEPICKER D_Fec2 WIDTH 90 VALUE DATE()

      IF DAY(DATE())<=15
         WinNomCont1.D_Fec1.Value:=DIAINIMES(DATE())-1
      ELSE
         WinNomCont1.D_Fec1.Value:=DIAFINMES(DATE())
      ENDIF


/*

      @040,010 GRID GR_Nominas ;
      HEIGHT 250 ;
      WIDTH 1000 ;
      HEADERS {'Cuenta','Descripcion','Nomina','Devengado','Seg.Social','Retencion','Varios','Liquido','A cuenta', ;
               'Asiento','Fecha','Banco','CIF'} ;
      WIDTHS {75,150,75,90,90,90,90,90,90,50,90,200,100 } ;
      ITEMS {} ;
      EDIT INPLACE {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'COMBOBOX',{'Nomina','Paga ext.','Finiquito'}}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC'}, ;
         {'DATEPICKER','DROPDOWN'}, ;
         {'TEXTBOX','CHARACTER','9999-9999-99-9999999999'},{'TEXTBOX','CHARACTER'}} ;
      COLUMNVALID { ,,,{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},,} ;
      COLUMNWHEN  { { || BusCta_Nomi1()},,,,,,,,,,,, } ;
      ON HEADCLICK {{|| Grid1_Ord(1)},{|| Grid1_Ord(2)},{|| Grid1_Ord(3)},{|| Grid1_Ord(4)},{|| Grid1_Ord(5)}, ;
                    {|| Grid1_Ord(6)},{|| Grid1_Ord(7)},{|| Grid1_Ord(8)},{|| Grid1_Ord(9)},{|| Grid1_Ord(10)}, ;
                    {|| Grid1_Ord(11)},{|| Grid1_Ord(12)},{|| Grid1_Ord(13)}} ;
      JUSTIFY {GRID_JTFY_RIGHT,GRID_JTFY_LEFT,GRID_JTFY_LEFT, ;
               GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT, ;
               GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_CENTER,GRID_JTFY_LEFT,GRID_JTFY_LEFT}
*/

      @300,310 TEXTBOX T_Sumas1 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
      @300,400 TEXTBOX T_Sumas2 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
      @300,490 TEXTBOX T_Sumas3 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
      @300,580 TEXTBOX T_Sumas4 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
      @300,670 TEXTBOX T_Sumas5 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN
      @300,760 TEXTBOX T_Sumas6 WIDTH 90 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN

      @040,010 GRID GR_Nominas ;
      HEIGHT 250 ;
      WIDTH 1000 ;
      HEADERS {'Cuenta','Descripcion','Nomina','Devengado','Seg.Social','Retencion','Varios','Liquido','A cuenta', ;
               'Asiento','Fecha','Banco','CIF'} ;
      WIDTHS {75,150,75,90,90,90,90,90,90,50,90,200,100 } ;
      ITEMS {} ;
      EDIT INPLACE ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'COMBOBOX',{'Nomina','Paga ext.','Finiquito'}}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC'}, ;
         {'DATEPICKER','DROPDOWN'}, ;
         {'TEXTBOX','CHARACTER','9999-9999-99-9999999999'},{'TEXTBOX','CHARACTER'}} ;
      VALID { ,,,{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},{ || Sumas_Nomi1()},,} ;
      COLUMNWHEN  { { || BusCta_Nomi1()},,,,,,,,,,,, } ;
      ON HEADCLICK {{|| Grid1_Ord(1)},{|| Grid1_Ord(2)},{|| Grid1_Ord(3)},{|| Grid1_Ord(4)},{|| Grid1_Ord(5)}, ;
                    {|| Grid1_Ord(6)},{|| Grid1_Ord(7)},{|| Grid1_Ord(8)},{|| Grid1_Ord(9)},{|| Grid1_Ord(10)}, ;
                    {|| Grid1_Ord(11)},{|| Grid1_Ord(12)},{|| Grid1_Ord(13)}} ;
      JUSTIFY {GRID_JTFY_RIGHT,GRID_JTFY_LEFT,GRID_JTFY_LEFT, ;
               GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT, ;
               GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_RIGHT,GRID_JTFY_CENTER,GRID_JTFY_LEFT,GRID_JTFY_LEFT}


      DEFINE CONTEXT MENU CONTROL GR_Nominas OF WinNomCont1
         ITEM "Añadir"   ACTION Menu_SobreImp("Nuevo")
         ITEM "Duplicar" ACTION Duplicar_Nomi1()
         ITEM "Eliminar" ACTION Eliminar_Nomi1()
         SEPARATOR
         ITEM "Copiar registro" ACTION Menu_SobreImp("CopiarRegistro")
         ITEM "Copiar tabla"    ACTION Menu_SobreImp("CopiarTabla")
         SEPARATOR
         ITEM "Ver asiento" ACTION Menu_SobreImp("Verasiento")
      END MENU

      @010,540 COMBOBOX C_RutaBanco WIDTH 270 ON CHANGE Actualiz_Nomi1("CUENTAS")
      @010,540 COMBOBOX C_CodBanco  WIDTH 270 INVISIBLE
         FOR N=1 TO LEN(aBANCO)
            WinNomCont1.C_RutaBanco.AddItem(aBANCO[N,2])
            WinNomCont1.C_CodBanco.AddItem( LEFT(aBANCO[N,3],AT("-",aBANCO[N,3])-1) )
         NEXT
      WinNomCont1.C_RutaBanco.Value:=1

      @300,210 BUTTON Bt_Sumar CAPTION 'Sumar' ;
         ACTION Sumas_Nomi1("TODO") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP


      @300,010 BUTTONEX Bt_BuscarTrab ;
         CAPTION 'Nominas trabajador' ICON icobus('buscar') ;
         ACTION Actualiz_Nomi1("TRABAJADOR") ;
         WIDTH 150 HEIGHT 25 NOTABSTOP

      @340,010 BUTTONEX Bt_BuscarCue ;
         CAPTION 'Banco' ICON icobus('buscar') ;
         ACTION br_suizocue(WinNomCont1.C_RutaBanco.DisplayValue,PCODCTA1(WinNomCont1.T_CodBan.Value),"WinNomCont1","T_CodBan") ;
         WIDTH 90 HEIGHT 25 NOTABSTOP
      @340,110 TEXTBOX T_CodBan WIDTH 100 VALUE "57200001" MAXLENGTH 8 ;
               ON CHANGE Actualiz_Banco() ;
               ON LOSTFOCUS WinNomCont1.T_CodBan.Value:=PCODCTA3(WinNomCont1.T_CodBan.Value)
      @340,220 TEXTBOX T_NomBan WIDTH 250 VALUE Codigos_NomCta(WinNomCont1.T_CodBan.Value) READONLY


      @340,610 BUTTON Bt_SegSoc CAPTION 'Contabilizar seguridad social' ;
         ACTION Contabilizar_SegSoc1() ;
         WIDTH 200 HEIGHT 25 NOTABSTOP


      @400,010 BUTTONEX Bt_Duplicar CAPTION 'Duplicar' ICON icobus('nuevo') ;
         ACTION Duplicar_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,110 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
         ACTION Eliminar_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,210 BUTTON Bt_Conta CAPTION 'Contabilizar' ;
         ACTION Contabilizar_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,310 BUTTON Bt_Remesar CAPTION 'Remesar' ;
         ACTION Remesar_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,410 BUTTON Bt_ContaPag CAPTION 'Conta. Pago' ;
         ACTION ContaPago_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,510 BUTTON Bt_Norma34 CAPTION 'Norma 34' ;
         ACTION Norma34_Nomi1() ;
         WIDTH 90 HEIGHT 25 NOTABSTOP

      @400,610 BUTTONEX Bt_Calc CAPTION 'Hoja Calc' ICON icobus('calc') ;
         ACTION Calc_Nomi1() ;
         WIDTH 80 HEIGHT 25 NOTABSTOP

      @400,700 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
         ACTION WinNomCont1.Release ;
         WIDTH 80 HEIGHT 25 NOTABSTOP

   Actualiz_Nomi1("CUENTAS")

   END WINDOW
   VentanaCentrar("WinNomCont1","Ventana1","Alinear")
   ACTIVATE WINDOW WinNomCont1


STATIC FUNCTION Menu_SobreImp(LLAMADA)
   IF WinNomCont1.GR_Nominas.Value<=0 .AND. LLAMADA<>"Nuevo"
      MSGSTOP("No hay ningun registro seleccionado")
      RETURN
   ENDIF
   DO CASE
   CASE LLAMADA="Nuevo"
      IF MSGYESNO("Desea añadir un registro en blanco")=.T.
         WinNomCont1.GR_Nominas.AddItem({"","",1,0,0,0,0,0,0,0,WinNomCont1.D_Fec1.Value,"",""})
         WinNomCont1.GR_Nominas.Value:=WinNomCont1.GR_Nominas.ItemCount
         WinNomCont1.GR_Nominas.Refresh
      ENDIF
   CASE LLAMADA="CopiarRegistro"
      IF MSGYESNO("Desea copiar el registro activo al portapapeles")=.T.
         TEXTO2:=""
         aCopiar:=WinNomCont1.GR_Nominas.Item(WinNomCont1.GR_Nominas.Value)
         FOR N1=1 TO LEN(aCopiar)
            DO CASE
            CASE VALTYPE(aCopiar[N1])="C"
               TEXTO2:=TEXTO2+aCopiar[N1]
            CASE VALTYPE(aCopiar[N1])="N"
               TEXTO2:=TEXTO2+STR(aCopiar[N1])
            CASE VALTYPE(aCopiar[N1])="D"
               TEXTO2:=TEXTO2+DTOC(aCopiar[N1])
            CASE VALTYPE(aCopiar[N1])="L"
               TEXTO2:=TEXTO2+IF(aCopiar[N1]=.T.,"Si","No")
            ENDCASE
            IF N1<LEN(aCopiar)
               TEXTO2:=TEXTO2+CHR(9)
            ENDIF
         NEXT
         //*** CopyToClipboard(TEXTO2)
         SetClipboardText(TEXTO2)
      ENDIF
   CASE LLAMADA="CopiarTabla"
      IF MSGYESNO("Desea copiar la tabla al portapapeles")=.T.
         TEXTO2:=""
         FOR N2=1 TO WinNomCont1.GR_Nominas.ItemCount
         aCopiar:=WinNomCont1.GR_Nominas.Item(N2)
            FOR N1=1 TO LEN(aCopiar)
               DO CASE
               CASE VALTYPE(aCopiar[N1])="C"
                  TEXTO2:=TEXTO2+aCopiar[N1]
               CASE VALTYPE(aCopiar[N1])="N"
                  TEXTO2:=TEXTO2+STR(aCopiar[N1])
               CASE VALTYPE(aCopiar[N1])="D"
                  TEXTO2:=TEXTO2+DTOC(aCopiar[N1])
               CASE VALTYPE(aCopiar[N1])="L"
                  TEXTO2:=TEXTO2+IF(aCopiar[N1]=.T.,"Si","No")
               ENDCASE
               IF N1<LEN(aCopiar)
                  TEXTO2:=TEXTO2+CHR(9)
               ENDIF
            NEXT
            TEXTO2:=TEXTO2+HB_OsNewLine()
         NEXT
         //*** CopyToClipboard(TEXTO2)
         SetClipboardText(TEXTO2)
      ENDIF
   CASE LLAMADA="Verasiento"
      IF WinNomCont1.GR_Nominas.Cell(WinNomCont1.GR_Nominas.Value,10)=0
         MSGSTOP("La nomina no esta contabilizada")
      ELSE
         br_suizoasi(WinNomCont1.C_RutaBanco.DisplayValue,WinNomCont1.GR_Nominas.Cell(WinNomCont1.GR_Nominas.Value,10),NUMEMP)
      ENDIF
   ENDCASE


STATIC FUNCTION Grid1_Ord(LLAMADA)
IF WinNomCont1.GR_Nominas.ItemCount=0
   MSGSTOP("No hay registros para ordenar")
   RETURN
ENDIF
IF VALTYPE(WinNomCont1.GR_Nominas.Cell(1,LLAMADA))<>"C" .AND. ;
   VALTYPE(WinNomCont1.GR_Nominas.Cell(1,LLAMADA))<>"N" .AND. ;
   VALTYPE(WinNomCont1.GR_Nominas.Cell(1,LLAMADA))<>"D"
   MSGSTOP("Esta columna no se puede ordenar")
   RETURN
ENDIF
PonerEspera("Ordenando la tabla por "+WinNomCont1.GR_Nominas.Header(LLAMADA) )
   ESTOY:=IF(WinNomCont1.GR_Nominas.Value=0,1,WinNomCont1.GR_Nominas.Value)
   aOrdenar2:={}
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      AADD(aOrdenar2,WinNomCont1.GR_Nominas.Item(N))
   NEXT
   DO CASE
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="C"
      ASORT(aOrdenar2,,, { |x, y| UPPER(x[LLAMADA]) < UPPER(y[LLAMADA]) })
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="N"
      ASORT(aOrdenar2,,, { |x, y| x[LLAMADA] < y[LLAMADA] })
   CASE VALTYPE(aOrdenar2[ESTOY,LLAMADA])="D"
      ASORT(aOrdenar2,,, { |x, y| DTOS(x[LLAMADA]) < DTOS(y[LLAMADA]) })
   ENDCASE
   WinNomCont1.GR_Nominas.DeleteAllItems
   FOR N=1 TO LEN(aOrdenar2)
      WinNomCont1.GR_Nominas.AddItem(aOrdenar2[N])
   NEXT
   WinNomCont1.GR_Nominas.Value:=ESTOY
QuitarEspera()


STATIC FUNCTION Actualiz_Nomi1(LLAMADA)
IF LLAMADA="CUENTAS"
   PonerEspera("Recopilando datos....")
   WinNomCont1.GR_Nominas.DeleteAllItems
   IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
      IF !FILE(WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.DBF") .AND. !FILE(WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.CDX")
         MSGSTOP("No se han localizado el fichero"+HB_OsNewLine()+WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.DBF","error")
         RETURN
      ENDIF
      AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      SET SOFTSEEK ON
      SEEK 46500000
      SET SOFTSEEK OFF
      DO WHILE STR(CODCTA,8)="465"
         IF NOMCTA="***"
            SKIP
            LOOP
         ENDIF
         CODCTA2:=STRZERO(CODCTA,8)
         NOMCTA2:=RTRIM(NOMCTA)
         BANCTA2:=BANCTA
         CODCTA3:="460"+RIGHT(CODCTA2,5)
         CIFCTA2:=RTRIM(CIF)
         ESTOY:=RECNO()
         SEEK VAL(CODCTA3)
         IMPACTA2:=IF(EOF(),0,SALDO)
         WinNomCont1.GR_Nominas.AddItem({CODCTA2,NOMCTA2,1,0,0,0,0,0,IMPACTA2,0,WinNomCont1.D_Fec1.Value,BANCTA2,CIFCTA2})
         GO ESTOY
         SKIP
      ENDDO
   ELSE
      AbrirDBF("SCTAB" ,,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      SET SOFTSEEK ON
      SEEK "46500000"
      SET SOFTSEEK OFF
      DO WHILE COD="465"
         IF TITULO="***"
            SKIP
            LOOP
         ENDIF
         SELEC SCTAB
         SEEK SUBCTA->COD
         BANCTA2:=IF(EOF(),"",CODBANCO+"-"+CODOFICINA+"-"+CODDC+"-"+CODCUENTA)
         SELEC SUBCTA
         CODCTA2:=STRZERO(VAL(COD),8)
         NOMCTA2:=RTRIM(TITULO)
         CODCTA3:="460"+RIGHT(CODCTA2,5)
         CIFCTA2:=RTRIM(NIF)
         ESTOY:=RECNO()
         SEEK CODCTA3
         IMPACTA2:=IF(EOF(),0,SUMADBEU-SUMAHBEU)
         WinNomCont1.GR_Nominas.AddItem({CODCTA2,NOMCTA2,1,0,0,0,0,0,IMPACTA2,0,WinNomCont1.D_Fec1.Value,BANCTA2,CIFCTA2})
         GO ESTOY
         SKIP
      ENDDO
   ENDIF
   IF WinNomCont1.GR_Nominas.ItemCount>0
      WinNomCont1.GR_Nominas.Value:=1
   ENDIF
ELSE
   CODCTA2:=WinNomCont1.GR_Nominas.Cell(WinNomCont1.GR_Nominas.Value,1)
   IF MSGYESNO("¿Desea borrar las lineas actuales?")=.T.
      WinNomCont1.GR_Nominas.DeleteAllItems
   ENDIF
   PonerEspera("Recopilando datos....")

   aNOMINA2:={}
   IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
      AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      GO TOP
      IF WinNomCont1.C_Mes.Value>MONTH(DMA1)
         WinNomCont1.D_Fec1.Value:=DIAFINMES(CTOD( "01-"+STRZERO(WinNomCont1.C_Mes.Value,2)+"-"+STRZERO(YEAR(DMA1),4) ))
      ELSE
         WinNomCont1.D_Fec1.Value:=DIAFINMES(CTOD( "01-"+STRZERO(WinNomCont1.C_Mes.Value,2)+"-"+STRZERO(YEAR(DMA2),4) ))
      ENDIF
      DO WHILE .NOT. EOF()
         IF LLAMADA="ASIENTO" .AND. MONTH(FECHA)=WinNomCont1.C_Mes.Value .OR. ;
            LLAMADA="TRABAJADOR" .AND. STR(CODCTA)=CODCTA2
            IF UPPER(NOMAPU)="NOMINA" .OR. UPPER(NOMAPU)="PAGA EXT" .OR. UPPER(NOMAPU)="FINIQUITO"
               NASI2:=NASI
               SEEK NASI2
               aNOMINA1:={"","",1,0,0,0,0,0,0,NASI,FECHA,"",""}
               DO WHILE .T.
                  DO CASE
                  CASE STR(CODCTA,8)="640"
                     aNOMINA1[4]:=DEBE
                  CASE STR(CODCTA,8)="476"
                     aNOMINA1[5]:=HABER
                  CASE STR(CODCTA,8)="4751"
                     aNOMINA1[6]:=HABER
                  CASE STR(CODCTA,8)="460"
                     aNOMINA1[9]:=HABER
                     aNOMINA1[8]:=aNOMINA1[8]+HABER
                  CASE STR(CODCTA,8)="465"
                     aNOMINA1[8]:=aNOMINA1[8]+HABER
                     aNOMINA1[1]:=STR(CODCTA)
                     aNOMINA1[3]:=IF(UPPER(NOMAPU)="PAGA"     ,2,aNOMINA1[3])
                     aNOMINA1[3]:=IF(UPPER(NOMAPU)="FINIQUITO",3,aNOMINA1[3])
                     AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
                     SEEK VAL(aNOMINA1[1])
                     IF .NOT. EOF()
                        aNOMINA1[2]:=RTRIM(NOMCTA)
                        aNOMINA1[12]:=BANCTA
                        aNOMINA1[13]:=RTRIM(CIF)
                     ENDIF
                     AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
                  ENDCASE
                  SKIP
                  IF NASI<>NASI2 .OR. EOF()
                     AADD(aNOMINA2,aNOMINA1)
                     SKIP -1
                     EXIT
                  ENDIF
               ENDDO
            ENDIF
         ENDIF
         SKIP
      ENDDO
   ELSE
      AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
      GO TOP
      IF WinNomCont1.C_Mes.Value>MONTH(DMA1)
         WinNomCont1.D_Fec1.Value:=DIAFINMES(CTOD( "01-"+STRZERO(WinNomCont1.C_Mes.Value,2)+"-"+STRZERO(YEAR(DMA1),4) ))
      ELSE
         WinNomCont1.D_Fec1.Value:=DIAFINMES(CTOD( "01-"+STRZERO(WinNomCont1.C_Mes.Value,2)+"-"+STRZERO(YEAR(DMA2),4) ))
      ENDIF
      DO WHILE .NOT. EOF()
         IF LLAMADA="ASIENTO" .AND. MONTH(FECHA)=WinNomCont1.C_Mes.Value .OR. ;
            LLAMADA="TRABAJADOR" .AND. SUBCTA=CODCTA2
            IF UPPER(NOMAPU)="NOMINA" .OR. UPPER(NOMAPU)="PAGA EXT" .OR. UPPER(NOMAPU)="FINIQUITO"
               NASI2:=ASIEN
               SEEK NASI2
               aNOMINA1:={"","",1,0,0,0,0,0,0,ASIEN,FECHA,"",""}
               DO WHILE .T.
                  DO CASE
                  CASE SUBCTA="640"
                     aNOMINA1[4]:=EURODEBE
                  CASE SUBCTA="476"
                     aNOMINA1[5]:=EUROHABER
                  CASE SUBCTA="4751"
                     aNOMINA1[6]:=EUROHABER
                  CASE SUBCTA="460"
                     aNOMINA1[9]:=EUROHABER
                     aNOMINA1[8]:=aNOMINA1[8]+EUROHABER
                  CASE SUBCTA="465"
                     aNOMINA1[8]:=aNOMINA1[8]+EUROHABER
                     aNOMINA1[1]:=SUBCTA
                     aNOMINA1[3]:=IF(UPPER(CONCEPTO)="PAGA"     ,2,aNOMINA1[3])
                     aNOMINA1[3]:=IF(UPPER(CONCEPTO)="FINIQUITO",3,aNOMINA1[3])
                     AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
                     SEEK aNOMINA1[1]
                     IF .NOT. EOF()
                        aNOMINA1[2]:=RTRIM(TITULO)
                        aNOMINA1[13]:=RTRIM(NIF)
                        AbrirDBF("SCTAB" ,,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
                        SEEK SUBCTA->COD
                        BANCTA2:=IF(EOF(),"",CODBANCO+"-"+CODOFICINA+"-"+CODDC+"-"+CODCUENTA)
                        aNOMINA1[12]:=BANCTA2
                     ENDIF
                     AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
                  ENDCASE
                  SKIP
                  IF ASIEN<>NASI2 .OR. EOF()
                     AADD(aNOMINA2,aNOMINA1)
                     SKIP -1
                     EXIT
                  ENDIF
               ENDDO
            ENDIF
         ENDIF
         SKIP
      ENDDO
   ENDIF
   ASORT(aNOMINA2,,,{|X,Y| X[1]<Y[1]})
   FOR N=1 TO LEN(aNOMINA2)
       WinNomCont1.GR_Nominas.AddItem(aNOMINA2[N])
   NEXT
ENDIF
   Sumas_Nomi1("TODO")
WinNomCont1.GR_Nominas.Refresh
WinNomCont1.GR_Nominas.SetFocus
QuitarEspera()
RETURN


STATIC FUNCTION Duplicar_Nomi1()
   IF MsgYesNo("¿Desea duplicar el registro activo?")=.F.
      RETURN
   ENDIF
   nNomina:=WinNomCont1.GR_Nominas.Value
   WinNomCont1.T_Sumas1.Value:=WinNomCont1.T_Sumas1.Value+WinNomCont1.GR_Nominas.Cell(nNomina,4)
   WinNomCont1.T_Sumas2.Value:=WinNomCont1.T_Sumas2.Value+WinNomCont1.GR_Nominas.Cell(nNomina,5)
   WinNomCont1.T_Sumas3.Value:=WinNomCont1.T_Sumas3.Value+WinNomCont1.GR_Nominas.Cell(nNomina,6)
   WinNomCont1.T_Sumas4.Value:=WinNomCont1.T_Sumas4.Value+WinNomCont1.GR_Nominas.Cell(nNomina,7)
   WinNomCont1.T_Sumas5.Value:=WinNomCont1.T_Sumas5.Value+WinNomCont1.GR_Nominas.Cell(nNomina,8)
   WinNomCont1.T_Sumas6.Value:=WinNomCont1.T_Sumas6.Value+WinNomCont1.GR_Nominas.Cell(nNomina,9)
   aNomina:={}
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      AADD(aNomina,WinNomCont1.GR_Nominas.Item(N))
      IF N=nNomina
         AADD(aNomina,WinNomCont1.GR_Nominas.Item(N))
      ENDIF
   NEXT
   WinNomCont1.GR_Nominas.DeleteAllItems
   FOR N=1 TO LEN(aNomina)
      WinNomCont1.GR_Nominas.AddItem(aNomina[N])
   NEXT
   WinNomCont1.GR_Nominas.Value:=nNomina
   WinNomCont1.GR_Nominas.Refresh


STATIC FUNCTION Eliminar_Nomi1()
   IF MsgYesNo("¿Desea eliminar el registro activo?")=.F.
      RETURN
   ENDIF
   nNomina:=WinNomCont1.GR_Nominas.Value
   WinNomCont1.T_Sumas1.Value:=WinNomCont1.T_Sumas1.Value-WinNomCont1.GR_Nominas.Cell(nNomina,4)
   WinNomCont1.T_Sumas2.Value:=WinNomCont1.T_Sumas2.Value-WinNomCont1.GR_Nominas.Cell(nNomina,5)
   WinNomCont1.T_Sumas3.Value:=WinNomCont1.T_Sumas3.Value-WinNomCont1.GR_Nominas.Cell(nNomina,6)
   WinNomCont1.T_Sumas4.Value:=WinNomCont1.T_Sumas4.Value-WinNomCont1.GR_Nominas.Cell(nNomina,7)
   WinNomCont1.T_Sumas5.Value:=WinNomCont1.T_Sumas5.Value-WinNomCont1.GR_Nominas.Cell(nNomina,8)
   WinNomCont1.T_Sumas6.Value:=WinNomCont1.T_Sumas6.Value-WinNomCont1.GR_Nominas.Cell(nNomina,9)
   WinNomCont1.GR_Nominas.DeleteItem(WinNomCont1.GR_Nominas.Value)
   IF WinNomCont1.GR_Nominas.ItemCount>0
      IF nNomina>WinNomCont1.GR_Nominas.ItemCount
         WinNomCont1.GR_Nominas.Value:=WinNomCont1.GR_Nominas.ItemCount
      ELSE
         WinNomCont1.GR_Nominas.Value:=nNomina
      ENDIF
   ENDIF
   WinNomCont1.GR_Nominas.Refresh


STATIC FUNCTION Sumas_Nomi1(LLAMADA)
local nVal := This.CellValue
local nCol := This.CellColIndex
local nRow := This.CellRowIndex
LLAMADA := IF(LLAMADA=NIL," ",LLAMADA)
IF LLAMADA="TODO"
   WinNomCont1.T_Sumas1.Value:=0
   WinNomCont1.T_Sumas2.Value:=0
   WinNomCont1.T_Sumas3.Value:=0
   WinNomCont1.T_Sumas4.Value:=0
   WinNomCont1.T_Sumas5.Value:=0
   WinNomCont1.T_Sumas6.Value:=0
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      WinNomCont1.T_Sumas1.Value:=WinNomCont1.T_Sumas1.Value+WinNomCont1.GR_Nominas.Cell(N,4)
      WinNomCont1.T_Sumas2.Value:=WinNomCont1.T_Sumas2.Value+WinNomCont1.GR_Nominas.Cell(N,5)
      WinNomCont1.T_Sumas3.Value:=WinNomCont1.T_Sumas3.Value+WinNomCont1.GR_Nominas.Cell(N,6)
      WinNomCont1.T_Sumas4.Value:=WinNomCont1.T_Sumas4.Value+WinNomCont1.GR_Nominas.Cell(N,7)
      WinNomCont1.GR_Nominas.Cell(N,8):= ;
         WinNomCont1.GR_Nominas.Cell(N,4)-WinNomCont1.GR_Nominas.Cell(N,5)- ;
         WinNomCont1.GR_Nominas.Cell(N,6)+WinNomCont1.GR_Nominas.Cell(N,7)
      WinNomCont1.T_Sumas5.Value:=WinNomCont1.T_Sumas5.Value+WinNomCont1.GR_Nominas.Cell(N,8)
      WinNomCont1.T_Sumas6.Value:=WinNomCont1.T_Sumas6.Value+WinNomCont1.GR_Nominas.Cell(N,9)
   NEXT
ELSE
   IF nCol>=4 .AND. nCol<=7
      SetProperty("WinNomCont1","T_Sumas"+LTRIM(STR(nCol-3)),"Value", ;
      GetProperty("WinNomCont1","T_Sumas"+LTRIM(STR(nCol-3)),"Value")-WinNomCont1.GR_Nominas.Cell(nRow,nCol)+nVal)
      IF nCol=4 .OR. nCol=7
         nVal2:=nVal*-1
      ELSE
         nVal2:=nVal
      ENDIF
      nTot1:=WinNomCont1.GR_Nominas.Cell(nRow,4)-WinNomCont1.GR_Nominas.Cell(nRow,5)- ;
             WinNomCont1.GR_Nominas.Cell(nRow,6)+WinNomCont1.GR_Nominas.Cell(nRow,7)+ ;
             WinNomCont1.GR_Nominas.Cell(nRow,nCol)-nVal2
      WinNomCont1.T_Sumas5.Value:=WinNomCont1.T_Sumas5.Value-WinNomCont1.GR_Nominas.Cell(nRow,8)+nTot1
      WinNomCont1.GR_Nominas.Cell(nRow,8):=nTot1
   ENDIF
   IF nCol=9
      WinNomCont1.T_Sumas6.Value:=WinNomCont1.T_Sumas6.Value-WinNomCont1.GR_Nominas.Cell(N,9)+nVal
   ENDIF
ENDIF
WinNomCont1.GR_Nominas.Refresh


STATIC FUNCTION Contabilizar_Nomi1()
   IF WinNomCont1.D_Fec1.Value<DMA1 .OR. WinNomCont1.D_Fec1.Value>DMA2
      MsgStop("La fecha esta fuera del ejercicio"+HB_OsNewLine()+DIA(WinNomCont1.D_Fec1.Value,10))
      RETURN
   ENDIF
   IF DAY(WinNomCont1.D_Fec1.Value)<15
      IF MsgYesNo("¿La fecha de las nominas es correcta?"+HB_OsNewLine()+ ;
                  "Nominas mes: "+MES(WinNomCont1.D_Fec1.Value)+HB_OsNewLine()+ ;
                  "Fecha: "+DIA(WinNomCont1.D_Fec1.Value,10) )=.F.
         RETURN
      ENDIF
   ENDIF
   IF MsgYesNo("¿Desea contabilizar las nominas?")=.F.
      RETURN
   ENDIF

PonerEspera("Contabilizando....")

IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
   AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
      DO EVENTS

      FECAPU2:=WinNomCont1.GR_Nominas.Cell(N,11)
      DO CASE
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=2
         NOMAPU2:="Paga ext."
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=3
         NOMAPU2:="Finiquito"
      OTHERWISE
         NOMAPU2:="Nomina"
      ENDCASE
      NOMAPU2:=NOMAPU2+" "+LEFT(MES(FECAPU2),3)+ ;
               +"-"+STR(YEAR(FECAPU2),4)+";"+PCODCTA4(WinNomCont1.GR_Nominas.Cell(N,1))

      CODCTA465:=VAL(PCODCTA3(WinNomCont1.GR_Nominas.Cell(N,1)))
      CODCTA460:=VAL("460"+RIGHT(STRZERO(CODCTA465,8),5))
      CODCTA640:=VAL("640"+RIGHT(STRZERO(CODCTA465,8),5))

      AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)

      IF WinNomCont1.GR_Nominas.Cell(N,10)<>0
         NASI2:=WinNomCont1.GR_Nominas.Cell(N,10)
         SEEK NASI2
         DO WHILE NASI=NASI2 .AND. .NOT. EOF()
            IF RLOCK()
               Suizo_saldocuenta(CODCTA,DEBE*-1,HABER*-1)
               DELETE
               DBCOMMIT()
               DBUNLOCK()
            ENDIF
            SKIP
         ENDDO
      ELSE
         GO BOTT
         NASI2:=NASI+1
      ENDIF
      NAPU2:=1

      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH CODCTA640
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE DEBE WITH WinNomCont1.GR_Nominas.Cell(N,4)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,5)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH 47600000
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE HABER WITH WinNomCont1.GR_Nominas.Cell(N,5)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,6)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH 47510000
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE HABER WITH WinNomCont1.GR_Nominas.Cell(N,6)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,7)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH CODCTA640
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE DEBE WITH WinNomCont1.GR_Nominas.Cell(N,7)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   ENDIF
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH CODCTA465
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE HABER WITH WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,9)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH FECAPU2
         REPLACE CODCTA WITH CODCTA460
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE HABER WITH WinNomCont1.GR_Nominas.Cell(N,9)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   ENDIF

   NEXT

ELSE

   AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
      DO EVENTS

      FECAPU2:=WinNomCont1.GR_Nominas.Cell(N,11)
      DO CASE
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=2
         NOMAPU2:="Paga ext."
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=3
         NOMAPU2:="Finiquito"
      OTHERWISE
         NOMAPU2:="Nomina"
      ENDCASE
      NOMAPU2:=NOMAPU2+" "+LEFT(MES(FECAPU2),3)+ ;
               +"-"+STR(YEAR(FECAPU2),4)+";"+PCODCTA4(WinNomCont1.GR_Nominas.Cell(N,1))

   IF UPPER(RUTAPROGRAMA)=UPPER("C:\Suizowin.ter") //BELEN TERRA
      NOMAPU2:=IF(WinNomCont1.GR_Nominas.Cell(N,3)=1,"Nomina","Finiquito")+" "+LEFT(MES(FECAPU2),3)
      NOMAPU2:=UPPER(NOMAPU2)
   ENDIF
      CODCTA465:=VAL(PCODCTA3(WinNomCont1.GR_Nominas.Cell(N,1)))
      CODCTA460:=VAL("460"+RIGHT(STRZERO(CODCTA465,8),5))
      CODCTA640:=VAL("640"+RIGHT(STRZERO(CODCTA465,8),5))

      AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
      GO BOTT
      NASI2:=ASIEN+1
      NAPU2:=1

      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(CODCTA640,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,4),YEAR(FECAPU2))
         REPLACE PTAHABER WITH 0
         REPLACE EURODEBE  WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,4),YEAR(FECAPU2))
         REPLACE EUROHABER WITH 0
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,5)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(47600000,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,5),YEAR(FECAPU2))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,5),YEAR(FECAPU2))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,6)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(47510000,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,6),YEAR(FECAPU2))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,6),YEAR(FECAPU2))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,7)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(CODCTA640,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,7),YEAR(FECAPU2))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,7),YEAR(FECAPU2))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   ENDIF
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(CODCTA465,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9),YEAR(FECAPU2))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9),YEAR(FECAPU2))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   IF WinNomCont1.GR_Nominas.Cell(N,9)<>0
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH FECAPU2
         REPLACE SUBCTA WITH STRZERO(CODCTA460,8)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,9),YEAR(FECAPU2))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,9),YEAR(FECAPU2))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   ENDIF

   NEXT

ENDIF
QuitarEspera()
MsgInfo('Los datos han sido guardados','Datos guardados')



STATIC FUNCTION ContaPago_Nomi1()
   IF WinNomCont1.D_Fec1.Value<DMA1 .OR. WinNomCont1.D_Fec1.Value>DMA2
      MsgStop("La fecha esta fuera del ejercicio"+HB_OsNewLine()+DIA(WinNomCont1.D_Fec1.Value,10))
      RETURN
   ENDIF
   IF MsgYesNo("¿Desea contabilizar el pago de las nominas?")=.F.
      RETURN
   ENDIF

PonerEspera("Contabilizando pago....")

IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
   AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)

      AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      GO BOTT
      NASI2:=NASI+1
      NAPU2:=1

   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
      DO EVENTS

      NOMAPU2:="Pago Nomina "+LEFT(MES(WinNomCont1.GR_Nominas.Cell(N,11)),3)+ ;
               +"-"+STR(YEAR(WinNomCont1.GR_Nominas.Cell(N,11)),4)

      CODCTA465:=VAL(PCODCTA3(WinNomCont1.GR_Nominas.Cell(N,1)))

      DO CASE
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=2
         NOMAPU3:=STRTRAN(NOMAPU2,"Nomina","Paga ext.")
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=3
         NOMAPU3:=STRTRAN(NOMAPU2,"Nomina","Finiquito")
      OTHERWISE
         NOMAPU3:=NOMAPU2
      ENDCASE

      AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH WinNomCont1.D_Fec2.Value
         REPLACE CODCTA WITH CODCTA465
         REPLACE NOMAPU WITH NOMAPU3+";"+PCODCTA4(WinNomCont1.GR_Nominas.Cell(N,1))
         REPLACE DEBE WITH WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9)
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF
   NEXT

      APPEND BLANK
      IF RLOCK()
         REPLACE NASI WITH NASI2
         REPLACE APU WITH NAPU2++
         REPLACE NEMP WITH NUMEMP
         REPLACE FECHA WITH WinNomCont1.D_Fec2.Value
         REPLACE CODCTA WITH PCODCTA1(WinNomCont1.T_CodBan.Value)
         REPLACE NOMAPU WITH NOMAPU2
         REPLACE HABER WITH WinNomCont1.T_Sumas5.Value-WinNomCont1.T_Sumas6.Value
         DBCOMMIT()
         DBUNLOCK()
         Suizo_saldocuenta(CODCTA,DEBE,HABER)
      ENDIF

ELSE

   AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)

      AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
      GO BOTT
      NASI2:=ASIEN+1
      NAPU2:=1

   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
      DO EVENTS

      NOMAPU2:="Pago Nomina "+LEFT(MES(WinNomCont1.GR_Nominas.Cell(N,11)),3)+ ;
               +"-"+STR(YEAR(WinNomCont1.GR_Nominas.Cell(N,11)),4)
   IF UPPER(RUTAPROGRAMA)=UPPER("C:\Suizowin.ter") //BELEN TERRA
      NOMAPU2:="PAGO TRANSF.RC "
   ENDIF

      CODCTA465:=VAL(PCODCTA3(WinNomCont1.GR_Nominas.Cell(N,1)))

      DO CASE
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=2
         NOMAPU3:=STRTRAN(NOMAPU2,"Nomina","Paga ext.")
      CASE WinNomCont1.GR_Nominas.Cell(N,3)=3
         NOMAPU3:=STRTRAN(NOMAPU2,"Nomina","Finiquito")
      OTHERWISE
         NOMAPU3:=NOMAPU2
      ENDCASE

      AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH WinNomCont1.D_Fec2.Value
         REPLACE SUBCTA WITH CODCTA465
         REPLACE CONCEPTO WITH NOMAPU3+WinNomCont1.GR_Nominas.Cell(N,2)
         REPLACE PTADEBE  WITH 0
         REPLACE PTAHABER WITH MDA_PTA(WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9),YEAR(WinNomCont1.D_Fec2.Value))
         REPLACE EURODEBE  WITH 0
         REPLACE EUROHABER WITH MDA_EURO(WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9),YEAR(WinNomCont1.D_Fec2.Value))
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF
   NEXT

      APPEND BLANK
      IF RLOCK()
         REPLACE ASIEN WITH NASI2
         REPLACE FECHA WITH WinNomCont1.D_Fec2.Value
         REPLACE SUBCTA WITH PCODCTA3(WinNomCont1.T_CodBan.Value)
         REPLACE CONCEPTO WITH NOMAPU2
         REPLACE PTADEBE  WITH MDA_PTA(WinNomCont1.T_Sumas5.Value-WinNomCont1.T_Sumas6.Value,YEAR(WinNomCont1.D_Fec2.Value))
         REPLACE PTAHABER WITH 0
         REPLACE EURODEBE  WITH MDA_EURO(WinNomCont1.T_Sumas5.Value-WinNomCont1.T_Sumas6.Value,YEAR(WinNomCont1.D_Fec2.Value))
         REPLACE EUROHABER WITH 0
         REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
         REPLACE NOCONV WITH .F.
         DBCOMMIT()
         DBUNLOCK()
         GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
      ENDIF

ENDIF
QuitarEspera()
MsgInfo('Los datos han sido guardados','Datos guardados')





STATIC FUNCTION Remesar_Nomi1()
   IF MSGYESNO("¿Desea remesar las nominas?")=.F.
      RETURN
   ENDIF

PonerEspera("Remesando....")

   AbrirDBF("REMESA")
   GO BOTT
   NREM2:=NREM+1
   LREM2:=1
   FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
      NOMAPU2:=IF(WinNomCont1.GR_Nominas.Cell(N,3)=1,"Nomina","Finiq.")+" "+;
            LEFT(MES(WinNomCont1.GR_Nominas.Cell(N,11)),3)+"-"+STR(YEAR(WinNomCont1.GR_Nominas.Cell(N,11)),4)

      AbrirDBF("REMESA")
      APPEND BLANK
      IF RLOCK()
         REPLACE SERIE WITH 6
         REPLACE NREM WITH NREM2
         REPLACE FREM WITH WinNomCont1.D_Fec2.Value
         REPLACE CODCTA WITH VAL(PCODCTA3(WinNomCont1.GR_Nominas.Cell(N,1)))
         REPLACE NOMCTA WITH WinNomCont1.GR_Nominas.Cell(N,2)
         REPLACE BANCTA WITH CTA_BAN_SUIZO( WinNomCont1.GR_Nominas.Cell(N,12) ,30)
         REPLACE NFRA WITH NOMAPU2
         REPLACE IMPORTE WITH WinNomCont1.GR_Nominas.Cell(N,8)
         REPLACE FVTO WITH WinNomCont1.D_Fec2.Value
         REPLACE CODBAN WITH VAL(PCODCTA3(WinNomCont1.T_CodBan.Value))
         REPLACE TRANOMI WITH 1
         REPLACE LREM WITH LREM2++
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
   NEXT
QuitarEspera()
MsgInfo('La remesa ha sido creada '+LTRIM(STR(NREM2)),'Datos guardados')


STATIC FUNCTION Actualiz_Banco()
   NOMCTA2:=""
IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
      IF !FILE(WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.DBF") .AND. !FILE(WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.CDX")
         MSGSTOP("No se han localizado el fichero"+HB_OsNewLine()+WinNomCont1.C_RutaBanco.DisplayValue+"\CUENTAS.DBF","error")
         RETURN
      ENDIF
   AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   SEEK PCODCTA1(WinNomCont1.T_CodBan.Value)
   IF .NOT. EOF()
      NOMCTA2:=RTRIM(NOMCTA)
   ENDIF
ELSE
   AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   CODCTA2:=PCODCTA3(WinNomCont1.T_CodBan.Value)
   SEEK CODCTA2
   IF .NOT. EOF()
      NOMCTA2:=RTRIM(TITULO)
   ENDIF
ENDIF
   WinNomCont1.T_NomBan.Value:=RTRIM(NOMCTA2)

STATIC FUNCTION BusCta_Nomi1()
local nVal := PCODCTA3(This.CellValue)
local nCol := This.CellColIndex
local nRow := This.CellRowIndex
   CODCTA2:=br_suizocue(WinNomCont1.C_RutaBanco.DisplayValue,VAL(nVal))
   AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   SEEK VAL(CODCTA2)
   IF .NOT. EOF()
      WinNomCont1.GR_Nominas.Cell(nRow,1):=CODCTA2
      WinNomCont1.GR_Nominas.Cell(nRow,2):=RTRIM(NOMCTA)
      WinNomCont1.GR_Nominas.Cell(nRow,12):=BANCTA
      WinNomCont1.GR_Nominas.Cell(nRow,13):=RTRIM(CIF)
   ELSE
      WinNomCont1.GR_Nominas.Cell(nRow,1):=""
      WinNomCont1.GR_Nominas.Cell(nRow,2):=""
      WinNomCont1.GR_Nominas.Cell(nRow,12):=""
      WinNomCont1.GR_Nominas.Cell(nRow,13):=""
   ENDIF
RETURN .F.


STATIC FUNCTION Calc_Nomi1()
   nOption:=InputWindow("Elija el tipo de salida",{"Programa"},{1},{{"Hoja Calculo","Hoja Excel"}},300,300,.T.,{"Aceptar","Cancelar"})
   DO CASE
   CASE nOption[1] == Nil
      RETURN
   CASE nOption[1]=1
      Calc_Nomi2()
   CASE nOption[1]=2
      Excel_Nomi1()
   ENDCASE

STATIC FUNCTION Calc_Nomi2()
   local oServiceManager,oDesktop,oDocument,oSchedule,oSheet,oCell,oColums,oColumn

   PonerEspera("Procesando datos del listado")

   // inicializa
   oServiceManager := TOleAuto():New("com.sun.star.ServiceManager")
   oDesktop := oServiceManager:createInstance("com.sun.star.frame.Desktop")
   IF oDesktop = NIL
      QuitarEspera()
      MsgStop('OpenOffice Calc no esta disponible','error')
      RETURN Nil
   ENDIF
   oDocument := oDesktop:loadComponentFromURL("private:factory/scalc","_blank", 0, {})

   // tomar hoja
   oSchedule := oDocument:GetSheets()

   // tomar primera hoja por nombre oSheet := oSchedule:GetByName("Hoja1")
   // o por indice
   oSheet := oSchedule:GetByIndex(0)

   // escribir en las celdas
   aMiColor:=MiColor("AMARILLOPALIDO")
   FOR N1=1 TO WinNomCont1.GR_Nominas.ItemCount
      FOR N2=1 TO 13
         IF N1=1
            oCell:=oSheet:getCellByPosition(N2-1,0) // columna,linea
            oCell:SetString(WinNomCont1.GR_Nominas.Header(N2))
            oCell:CharWeight:=150 //NEGRITA
            oCell:CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
         ENDIF
         oCell:=oSheet:getCellByPosition(N2-1,N1) // columna,linea
         IF VALTYPE(WinNomCont1.GR_Nominas.Cell(N1,N2))="N"
            oCell:SetValue(WinNomCont1.GR_Nominas.Cell(N1,N2))
         ELSE
            oCell:SetString(WinNomCont1.GR_Nominas.Cell(N1,N2))
         ENDIF
      NEXT
      oSheet:getCellRangeByPosition(3,N1,8,N1):NumberFormat:=4 //#.##0,00
   NEXT

   LIN1:=1+1
   LIN:=WinNomCont1.GR_Nominas.ItemCount+1
   oSheet:getCellByPosition(2,LIN):SetString("TOTAL") 
   FOR N1=3 TO 8
      oSheet:getCellByPosition(N1,LIN):SetFormula("=SUM("+LetraExcel(N1+1)+LTRIM(STR(LIN1))+":"+LetraExcel(N1+1)+LTRIM(STR(LIN))+")")
   NEXT
   oSheet:getCellRangeByPosition(3,LIN,8,LIN):NumberFormat:=4 //#.##0,00
   oSheet:getCellRangeByPosition(2,LIN,8,LIN):CharWeight:=150 //NEGRITA
   aMiColor:=MiColor("VERDEPALIDO")
   oSheet:getCellRangeByPosition(2,LIN,8,LIN):CellBackColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])

   oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)

   QuitarEspera()

return nil 



STATIC FUNCTION Excel_Nomi1()
   LOCAL oExcel, oHoja

   PonerEspera("Procesando datos del listado")

   oExcel := TOleAuto():New( "Excel.Application" )
   IF Ole2TxtError() != 'S_OK'
      QuitarEspera()
      MsgStop('Excel no esta disponible','error')
      RETURN Nil
   ENDIF
   oExcel:WorkBooks:Add()
   oExcel:Sheets("Hoja1"):Name := "Nominas"
   oHoja := oExcel:Get( "ActiveSheet" )
   oHoja:Cells:Font:Name := "Arial"
   oHoja:Cells:Font:Size := 10

   FOR N1=1 TO WinNomCont1.GR_Nominas.ItemCount
      FOR N2=1 TO 13
         IF N1=1
            oHoja:Cells( N1, N2 ):Value := WinNomCont1.GR_Nominas.Header(N2)
            oHoja:Cells( N1, N2 ):Interior:ColorIndex := 36 //sombrear celdas  amarillo claro
            oHoja:Cells( N1, N2 ):Font:Bold := .T.
         ENDIF
         oHoja:Cells( N1+1, N2 ):Value := WinNomCont1.GR_Nominas.Cell(N1,N2)
      NEXT
      oHoja:Range(LetraExcel(4)+LTRIM(STR(N1+1))+":"+LetraExcel(9)+LTRIM(STR(N1+1))):Set( "NumberFormat", "#.##0,00" )
   NEXT

   LIN1:=1+1
   LIN:=WinNomCont1.GR_Nominas.ItemCount+2
   oHoja:Cells(LIN,3):Value:="TOTAL"
   FOR N1=4 TO 9
      oHoja:Cells(LIN,N1):Value:="=SUMA("+LetraExcel(N1)+LTRIM(STR(LIN1))+":"+LetraExcel(N1)+LTRIM(STR(LIN-1))+")"
   NEXT
   oHoja:Range(LetraExcel(4)+LTRIM(STR(LIN))+":"+LetraExcel(9)+LTRIM(STR(LIN))):Set( "NumberFormat", "#.##0,00" )
   oHoja:Range(LetraExcel(3)+LTRIM(STR(LIN))+":"+LetraExcel(9)+LTRIM(STR(LIN))):Font:Bold := .T.
   oHoja:Range(LetraExcel(3)+LTRIM(STR(LIN))+":"+LetraExcel(9)+LTRIM(STR(LIN))):Interior:ColorIndex := 35 //sombrear celdas verde claro

   FOR nCol:=1 TO FCOUNT()
      oHoja:Columns( nCol ):AutoFit()
   NEXT

   oHoja:Cells( 1, 1 ):Select()
   oExcel:Visible := .T.

   IF AT("XHARBOUR",UPPER(VERSION()))=1
      oHoja:end()
      oExcel:end()
   ENDIF

   QuitarEspera()

Return Nil



STATIC FUNCTION Norma34_Nomi1()
LOCAL PAG:=0,LINR:=0,TOT1:=0,EURO2:=1,NSER2:=6
CODBAN2:=VAL(RIGHT(PCODCTA3(WinNomCont1.T_CodBan.Value),4))

PonerEspera("Creando fichero de normas")

NEMP2:=WinNomCont1.C_CodBanco.Item(WinNomCont1.C_RutaBanco.Value)
IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
   RUTA2:=LEFT(WinNomCont1.C_RutaBanco.DisplayValue,RAT("\",WinNomCont1.C_RutaBanco.DisplayValue)-1)
   IF !FILE(RUTA2+"\EMPRESA.DBF") .AND. !FILE(RUTA2+"\EMPRESA.CDX")
      MSGSTOP("No se han localizado el fichero"+HB_OsNewLine()+RUTA2+"\EMPRESA.DBF","error")
      RETURN
   ENDIF
   AbrirDBF("empresa",,,,RUTA2,1)
   SEEK VAL(NEMP2)
   NORNIFEMP:=IMPNIF(CIF)
   NORNOMEMP:=RTRIM(EMP)
   NORDIREMP:=RTRIM(DIRECCION)
   NORPOBEMP:=RTRIM(CODPOS+"-"+POBLACION)
   AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
   SEEK PCODCTA1(WinNomCont1.T_CodBan.Value)
   NORCTAEMP:=CTA_BAN_SUIZO(BANCTA,20)
ELSE
***modificar***
   AbrirDBF("empresa",,,,RUTAPROGRAMA)
   GO RECNOEMPRESA
   NORNIFEMP:=IMPNIF(CIF)
   NORNOMEMP:=RTRIM(EMP)
   NORDIREMP:=RTRIM(DIRECCION)
   NORPOBEMP:=RTRIM(CODPOS+"-"+POBLACION)
   NORCTAEMP:=CTA_BAN_SUIZO(BANCTA,20)
ENDIF

SET PRINTER TO "Rem_n34.txt"

FOR N=1 TO WinNomCont1.GR_Nominas.ItemCount
      IF WinNomCont1.GR_Nominas.Cell(N,8)=0
         LOOP
      ENDIF
   SET PRINT ON
   IF LINR=0
*** Registros de cabecera de Ordenante
      ?? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "001" //E
      ?? DIA(DATE(),-6) //F1
      ?? DIA(DATE(),-6) //F2
      ?? PADL(ALLTRIM(SUBSTR(NORCTAEMP,1,4)),4,"0")    //F3
      ?? PADL(ALLTRIM(SUBSTR(NORCTAEMP,5,4)),4,"0")    //F4
      ?? PADL(ALLTRIM(SUBSTR(NORCTAEMP,9,2)),2,"*")    //F5
      ?? PADL(ALLTRIM(SUBSTR(NORCTAEMP,11,10)),10,"0") //F6
      ?? "0" //F7
      ?? SPACE(8) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "002" //E
      ?? PADR(NORNOMEMP,36," ") //F
      ?? SPACE(5) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "003" //E
      ?? PADR(NORDIREMP,36," ") //F
      ?? SPACE(5) //G

      ? "03" //A
      ?? IF(EURO2=0,"06","62") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? "004" //E
      ?? PADR(NORPOBEMP,36," ") //F
      ?? SPACE(5) //G

*** Registro de Cabecera
      ? "04" //A
      ?? IF(EURO2=0,"06","56") //B
      ?? PADL(NORNIFEMP,9,"0") //C1
      ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
      ?? SPACE(12) //D
      ?? SPACE(03) //E
      ?? SPACE(41) //F

   ENDIF
*** Registros de Detalle

      NOMAPU2:=IF(WinNomCont1.GR_Nominas.Cell(N,3)=1,"Nomina","Finiq.")+" "+;
            LEFT(MES(WinNomCont1.GR_Nominas.Cell(N,11)),3)+"-"+STR(YEAR(WinNomCont1.GR_Nominas.Cell(N,11)),4)

   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(ALLTRIM(WinNomCont1.GR_Nominas.Cell(N,13)),12,"0") //D
   ?? "010" //E
   ?? FTONUMNOR( WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9) ,12,2) //F1
   ?? PADL(ALLTRIM(SUBSTR(WinNomCont1.GR_Nominas.Cell(N,12),1,4)),4,"0")       //F2
   ?? PADL(ALLTRIM(SUBSTR(WinNomCont1.GR_Nominas.Cell(N,12),6,4)),4,"0")       //F3
   ?? PADL(ALLTRIM(SUBSTR(WinNomCont1.GR_Nominas.Cell(N,12),11,2)),2,"*")      //F4
   ?? PADL(ALLTRIM(SUBSTR(WinNomCont1.GR_Nominas.Cell(N,12),14,10)),10,"0")    //F5
   ?? "1"      //F6
   ?? "1"      //F7 1-NOMINA 8-PENSION 9-OTROS
   ?? "1"      //F8
   ?? SPACE(6) //G

   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(ALLTRIM(WinNomCont1.GR_Nominas.Cell(N,13)),12,"0") //D
   ?? "011" //E
   ?? PADR(WinNomCont1.GR_Nominas.Cell(N,2),36," ") //F
   ?? SPACE(5) //G

   ? "06" //A
   IF NSER2=5
      ?? IF(EURO2=0,"07","57") //B
   ELSE
      ?? IF(EURO2=0,"06","56") //B
   ENDIF
   ?? PADL(NORNIFEMP,9,"0") //C1
   ?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
   ?? PADL(ALLTRIM(WinNomCont1.GR_Nominas.Cell(N,13)),12,"0") //D
   ?? "016" //E
   ?? PADR(NOMAPU2,36," ") //F
   ?? SPACE(5) //G

   SET PRINT OFF
   TOT1:=TOT1+WinNomCont1.GR_Nominas.Cell(N,8)-WinNomCont1.GR_Nominas.Cell(N,9)
   LINR++
   SELEC 1
   SKIP
NEXT

*** Registro de totales
SET PRINT ON
? "08" //A
??  IF(EURO2=0,"06","56") //B
?? PADL(NORNIFEMP,9,"0") //C1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
?? SPACE(12) //D
?? SPACE( 3) //E
?? FTONUMNOR(TOT1,12,2) //F1
?? PADL(LTRIM(STR(LINR)),8,"0") //F2
?? PADL(LTRIM(STR((LINR*3)+2)),10,"0") //F3  lineas 0362 y 0962 no se cuentan
?? SPACE(6) //F4
?? SPACE(5) //G

*** Registro de total general
? "09" //A
??  IF(EURO2=0,"06","62") //B
?? PADL(NORNIFEMP,9,"0") //C1
?? PADL(LTRIM(STR(CODBAN2)),3,"0") //C2
?? SPACE(12) //D
?? SPACE( 3) //E
?? FTONUMNOR(TOT1,12,2) //F1
?? PADL(LTRIM(STR(LINR)),8,"0") //F2
?? PADL(LTRIM(STR((LINR*3)+7)),10,"0") //F3
?? SPACE(6) //F4
?? SPACE(5) //G

?
SET PRINT OFF
SET PRINTER TO

QuitarEspera()
MsgInfo('El fichero de norma 34-1 se ha creado en:'+HB_OsNewLine()+RUTAEMPRESA+"\Rem_n34.txt",'Datos guardados')
//*** CopyToClipboard(RUTAEMPRESA+"\Rem_n34.txt")
SetClipboardText(RUTAEMPRESA+"\Rem_n34.txt")

STATIC FUNCTION Contabilizar_SegSoc1()
   DEFINE WINDOW W_NomSS ;
      AT 10,10 ;
      WIDTH 400 HEIGHT 330 ;
      TITLE 'Contabilizar seguridad social' ;
      MODAL      ;
      NOSIZE


      @010,010 CHECKBOX SiSS CAPTION 'Contabilizar Seg.Social' WIDTH 150 VALUE .T.
      @010,200 DATEPICKER D_SS WIDTH 90 VALUE WinNomCont1.D_Fec1.Value
      @045,010 LABEL   L_SSEmp VALUE 'Seg.Social empresa' AUTOSIZE TRANSPARENT
      @040,150 TEXTBOX T_SSEmp WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN ;
         ON LOSTFOCUS W_NomSS.T_Pago.Value:=W_NomSS.T_SSEmp.Value+W_NomSS.T_SSTra.Value-W_NomSS.T_PagDe.Value-W_NomSS.T_Bonif.Value
      @075,010 LABEL   L_SSTra VALUE 'Seg.Social trabajador' AUTOSIZE TRANSPARENT
      @070,150 TEXTBOX T_SSTra WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN ;
         ON LOSTFOCUS W_NomSS.T_Pago.Value:=W_NomSS.T_SSEmp.Value+W_NomSS.T_SSTra.Value-W_NomSS.T_PagDe.Value-W_NomSS.T_Bonif.Value
      @105,010 LABEL   L_PagDe VALUE 'Pago delegado' AUTOSIZE TRANSPARENT
      @100,150 TEXTBOX T_PagDe WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN ;
         ON LOSTFOCUS W_NomSS.T_Pago.Value:=W_NomSS.T_SSEmp.Value+W_NomSS.T_SSTra.Value-W_NomSS.T_PagDe.Value-W_NomSS.T_Bonif.Value
      @135,010 LABEL   L_Bonif VALUE 'Bonificacion' AUTOSIZE TRANSPARENT
      @130,150 TEXTBOX T_Bonif WIDTH 120 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN ;
         ON LOSTFOCUS W_NomSS.T_Pago.Value:=W_NomSS.T_SSEmp.Value+W_NomSS.T_SSTra.Value-W_NomSS.T_PagDe.Value-W_NomSS.T_Bonif.Value
      @170,010 CHECKBOX SiPago CAPTION 'Contabilizar el pago' WIDTH 150 VALUE .T.
      @170,150 TEXTBOX T_Pago WIDTH 120 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN
      @170,290 DATEPICKER D_Pago WIDTH 90 VALUE DIAFINMES(DIAMESMAS(WinNomCont1.D_Fec1.Value))-1

      MiColor_Enabled(.F.,"W_NomSS","T_Pago")
      W_NomSS.T_SSTra.Value:=WinNomCont1.T_Sumas2.Value

      @260,210 BUTTON Bt_Conta ;
         CAPTION 'Contabilizar' ;
         WIDTH 90 HEIGHT 25 ;
         ACTION Contabilizar_SegSoc2() ;
         NOTABSTOP

      @260,310 BUTTONEX Bt_Salir ;
         CAPTION 'Salir' ;
         ICON icobus('salir') ;
         ACTION W_NomSS.Release ;
         WIDTH 80 HEIGHT 25 ;
         NOTABSTOP

      END WINDOW
      VentanaCentrar("W_NomSS","Ventana1")
      ACTIVATE WINDOW W_NomSS

Return Nil


STATIC FUNCTION Contabilizar_SegSoc2()
   IF W_NomSS.D_SS.Value<DMA1 .OR. W_NomSS.D_SS.Value>DMA2
      MsgStop("La fecha esta fuera del ejercicio"+HB_OsNewLine()+DIA(W_NomSS.D_SS.Value,10))
      RETURN
   ENDIF

   IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
      AbrirDBF("CUENTAS",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      AbrirDBF("APUNTES",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      GO BOTT
      NASI2:=NASI
   ELSE
      AbrirDBF("SUBCTA",,,,WinNomCont1.C_RutaBanco.DisplayValue,1)
      AbrirDBF("DIARIO",,,,WinNomCont1.C_RutaBanco.DisplayValue,3)
      GO BOTT
      NASI2:=ASIEN+1
   ENDIF

   aAsiento:={}
   IF W_NomSS.SiSS.Value=.T.
      IF MsgYesNo("¿Desea contabilizar la seguridad social?")=.T.
         NASI2++
         NAPU2:=1
         NOMAPU2:="Seg.Social "+MES(W_NomSS.D_SS.Value)+" "+STRZERO(YEAR(W_NomSS.D_SS.Value),4)
         AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_SS.Value,'64200000',NOMAPU2,W_NomSS.T_SSEmp.Value,0})
         IF W_NomSS.T_PagDe.Value<>0
            AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_SS.Value,'64200000',NOMAPU2,W_NomSS.T_PagDe.Value*-1,0})
         ENDIF
         IF W_NomSS.T_Bonif.Value<>0
            AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_SS.Value,'64200000',NOMAPU2,W_NomSS.T_Bonif.Value*-1,0})
         ENDIF
         AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_SS.Value,'47600000',NOMAPU2,0,W_NomSS.T_SSEmp.Value-W_NomSS.T_PagDe.Value-W_NomSS.T_Bonif.Value})
      ENDIF
   ENDIF
   IF W_NomSS.SiPago.Value=.T.
      IF MsgYesNo("¿Desea contabilizar el pago de la seguridad social?")=.T.
         NASI2++
         NAPU2:=1
         NOMAPU2:="Pago Seg.Soc."+MES(W_NomSS.D_SS.Value)+" "+STRZERO(YEAR(W_NomSS.D_SS.Value),4)
         AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_Pago.Value,'47600000',NOMAPU2,W_NomSS.T_Pago.Value,0})
         AADD(aAsiento,{NASI2,NAPU2++,NUMEMP,W_NomSS.D_Pago.Value,PCODCTA3(WinNomCont1.T_CodBan.Value),NOMAPU2,0,W_NomSS.T_Pago.Value})
      ENDIF
   ENDIF

   FOR N=1 TO LEN(aAsiento)
      IF AT("GRUPOSP",UPPER(WinNomCont1.C_RutaBanco.DisplayValue))=0
         APPEND BLANK
         IF RLOCK()
            REPLACE NASI   WITH aAsiento[N,1]
            REPLACE APU    WITH aAsiento[N,2]
            REPLACE NEMP   WITH aAsiento[N,3]
            REPLACE FECHA  WITH aAsiento[N,4]
            REPLACE CODCTA WITH VAL(aAsiento[N,5])
            REPLACE NOMAPU WITH aAsiento[N,6]
            REPLACE DEBE   WITH aAsiento[N,7]
            REPLACE HABER  WITH aAsiento[N,8]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
         ENDIF
      ELSE
         APPEND BLANK
         IF RLOCK()
            REPLACE ASIEN     WITH aAsiento[N,1]
            REPLACE FECHA     WITH aAsiento[N,4]
            REPLACE SUBCTA    WITH aAsiento[N,5]
            REPLACE CONCEPTO  WITH aAsiento[N,6]
            REPLACE PTADEBE   WITH MDA_PTA(aAsiento[N,7],YEAR(aAsiento[N,4]))
            REPLACE PTAHABER  WITH MDA_PTA(aAsiento[N,8],YEAR(aAsiento[N,4]))
            REPLACE EURODEBE  WITH MDA_EURO(aAsiento[N,7],YEAR(aAsiento[N,4]))
            REPLACE EUROHABER WITH MDA_EURO(aAsiento[N,8],YEAR(aAsiento[N,4]))
            REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
            REPLACE NOCONV    WITH .F.
            DBCOMMIT()
            DBUNLOCK()
            GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
         ENDIF
      ENDIF
   NEXT
MsgInfo('Los datos han sido guardados','Datos guardados')


Return Nil


*
STATIC FUNCTION FTONUMNOR(IMP2,LARGO2,DEC2)
   LARGO2:=IF(DEC2=0,LARGO2,LARGO2+1)
   NOM2:=STRZERO(IMP2,LARGO2,DEC2)
   NOM2:=STRTRAN(NOM2,".","")
   RETURN(NOM2)







