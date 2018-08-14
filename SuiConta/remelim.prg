#include "minigui.ch"

procedure RemElim()
   TituloImp:="Eliminar remesa"

   DO CASE
   CASE UPPER(PROGRAMA())="SUICONTA"
      RUTAREMESA:=RUTAEMPRESA
      RUTAREMEMP:=RUTAPROGRAMA
   CASE FILE(RUTACONTABLE+"\REMESA.DBF")=.T.
      RUTAREMESA:=RUTACONTABLE
      RUTAREMEMP:=LEFT(RUTACONTABLE,RAT("\",RUTACONTABLE)-1)
   CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)+"\SUICONTA.EXE") = .T.
      RUTAREMEMP:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","SUICONTA",,1)
      AbrirDBF("empresa",,,,RUTAREMEMP,1)
      GO BOTT
      RUTAREMESA:=RUTAREMEMP+"\"+RTRIM(RUTA)
   CASE FILE(STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)+"\REMESA.EXE") = .T.
      RUTAREMESA:=STRTRAN(UPPER(RUTAPROGRAMA),"ESCOLA","REMESA",,1)
      RUTAREMEMP:=RUTAREMESA
   OTHERWISE
      RUTAREMESA:=RUTAEMPRESA
      RUTAREMEMP:=RUTAPROGRAMA
   ENDCASE

   aREMESAS:=Remesa_alis(RUTAREMESA)

   DEFINE WINDOW W_Imp1 ;
      AT 10,10 ;
      WIDTH 900 HEIGHT 380 ;
      TITLE 'Imprimir: '+TituloImp ;
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

      @310, 10 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') WIDTH 90 HEIGHT 25 ;
               ACTION RemElim2()

      @310,110 BUTTONEX B_Salir CAPTION 'Salir' ICON icobus('salir') WIDTH 90 HEIGHT 25 ;
               ACTION W_Imp1.release

   RemesaActRec()

   END WINDOW
   VentanaCentrar("W_Imp1","Ventana1")
   ACTIVATE WINDOW W_Imp1

Return Nil



STATIC FUNCTION RemElim2()
   NSER2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
   NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
   FREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,4)

   IF MSGYESNO("¿Desea eliminar la remesa?"+HB_OsNewLine()+"Remesa: "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)+" "+DIA(FREM2,10))=.F.
      RETURN
   ENDIF

   AbrirDBF("REMESA",,,,RUTAREMESA)
   SEEK (NSER2*100000)+NREM2
   DO WHILE NSER2=SERIE .AND. NREM2=NREM .AND. .NOT. EOF()
      IF RLOCK()
         DELETE
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
      SKIP
   ENDDO

   AbrirDBF("REMLIN",,,,RUTAREMESA)
   SEEK (NSER2*100000)+NREM2
   DO WHILE NSER2=SERIE .AND. NREM2=NREM .AND. .NOT. EOF()
      IF RLOCK()
         DELETE
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
      SKIP
   ENDDO

   W_Imp1.GR_Remesas.DeleteItem(W_Imp1.GR_Remesas.Value)

   MSGINFO("La remesa se ha eliminado correctamente"+HB_OsNewLine()+"Remesa: "+LTRIM(STR(NREM2))+"-"+REM_NOM1(NSER2)+" "+DIA(FREM2,10))









*********************importar datos de remesa******************
FUNCTION RemImpBase()
RUTAREMESA:=UPPER(RUTAPROGRAMA)
RUTAREMESA:=STRTRAN(RUTAREMESA,"SUIZOWIN","SUIZO")
RUTAREMESA:=STRTRAN(RUTAREMESA,"SUICONTA","REMESA")
RUTAREMESA:=STRTRAN(RUTAREMESA,"ESCOLA","REMESA")
IF MSGYESNO("Desea importar los datos de la base:"+HB_OsNewLine()+ ;
            RUTAREMESA)=.F.
   RETURN
ENDIF
IF .NOT. FILE(RUTAREMESA+"\REM"+STR(EJERANY,4)+".DBF") .OR. ;
   .NOT. FILE(RUTAREMESA+"\REM"+STR(EJERANY,4)+".CDX") .OR. ;
   .NOT. FILE(RUTAREMESA+"\LIN"+STR(EJERANY,4)+".DBF") .OR. ;
   .NOT. FILE(RUTAREMESA+"\LIN"+STR(EJERANY,4)+".CDX")
   MSGSTOP("No se han localizados los ficheros:"+HB_OsNewLine()+ ;
         RUTAREMESA+"\REM"+STR(EJERANY,4)+".DBF"+HB_OsNewLine()+ ;
         RUTAREMESA+"\REM"+STR(EJERANY,4)+".CDX"+HB_OsNewLine()+ ;
         RUTAREMESA+"\LIN"+STR(EJERANY,4)+".DBF"+HB_OsNewLine()+ ;
         RUTAREMESA+"\LIN"+STR(EJERANY,4)+".CDX")
   RETURN
ENDIF

IF .NOT. FILE("REMESA.DBF")
   MSGSTOP("El fichero de remesas no existe","error")
   RETURN
ENDIF

   AbrirDBF("REMESA")

IF LASTREC()>0
   IF MSGYESNO("El fichero de remesas no esta vacio"+HB_OsNewLine()+ ;
               "¿Desea seguir con la importacion?")=.F.
      RETURN
   ENDIF
ENDIF

   PonerEspera("Importando datos de: "+RUTAREMESA+"\REM"+STR(EJERANY,4)+".DBF")

   AbrirDBF("REMESA")
   AbrirDBF("REM"+STR(EJERANY,4)+".DBF",,,,RUTAREMESA)
   GO TOP
   DO WHILE .NOT. EOF()
      aDATOS:={SERIE,NREM,LREM,FREM,NFRA,CODCTA,CLIENTE,IMPORTE,BANCO,PLAZA,FVTO,CODBAN,TRANOMI}
      AbrirDBF("REMESA")
      APPEND BLANK
      IF RLOCK()
         REPLACE SERIE   WITH aDATOS[1]
         REPLACE NREM    WITH aDATOS[2]
         REPLACE LREM    WITH aDATOS[3]
         REPLACE FREM    WITH aDATOS[4]
         REPLACE NFRA    WITH HB_OEMtoANSI(RTRIM(aDATOS[5]))
         REPLACE CODCTA  WITH aDATOS[6]
         REPLACE NOMCTA  WITH HB_OEMtoANSI(RTRIM(aDATOS[7]))
         REPLACE IMPORTE WITH aDATOS[8]
         REPLACE BANCTA  WITH aDATOS[9]
         REPLACE TALON   WITH aDATOS[10]
         REPLACE FVTO    WITH aDATOS[11]
         REPLACE CODBAN  WITH aDATOS[12]
         REPLACE TRANOMI WITH aDATOS[13]
         DBCOMMIT()
         DBUNLOCK()
      ENDIF
      AbrirDBF("REM"+STR(EJERANY,4)+".DBF",,,,RUTAREMESA)
      SKIP
   ENDDO

   AbrirDBF("REMLIN")
   FICHEROLIN:=RUTAREMESA+"\LIN"+STR(EJERANY,4)+".DBF"
   APPEND FROM &FICHEROLIN

   IF AT("ESCOLA",UPPER(PROGRAMA()))<>0
      AbrirDBF("BANCOS",,,,RUTAREMESA)
      DO WHILE .NOT. EOF()
         aDATOS:={COD,NOMBRE,CODCTA}
         AbrirDBF("CUENTAS")
         APPEND BLANK
         IF RLOCK()
            REPLACE CODCTA WITH aDATOS[1]+57200000
            REPLACE NOMCTA WITH HB_OEMtoANSI(RTRIM(aDATOS[2]))
            REPLACE BANCTA WITH CTA_BAN_SUIZO(aDATOS[3],24)
            DBCOMMIT()
            DBUNLOCK()
         ENDIF
         AbrirDBF("BANCOS",,,,RUTAREMESA)
         SKIP
      ENDDO
   ENDIF

   QuitarEspera()
   MSGINFO("La importacion se ha realizado con exito")
*********************FIN importar datos de remesa******************


