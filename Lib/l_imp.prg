#include "minigui.ch"
#include "winprint.ch"

procedure Impresoras(MI_IMPRESORA)
   MI_IMPRESORA:=IF(MI_IMPRESORA=NIL,0,MI_IMPRESORA)
   INIT PRINTSYS
*   aIMP1:=aclone(hbprn:printers)
   GET PRINTERS TO aIMP1
   ASORT(aIMP1,,, { |x, y| UPPER(x) < UPPER(y) })
   aIMP3:=0
   IF VALTYPE(MI_IMPRESORA)="C"
         aIMP2:=MI_IMPRESORA
         aIMP3:=ASCAN(aIMP1, {|aVal| PADR(aVal,LEN(aIMP2)," ") == aIMP2})
   ENDIF
   IF aIMP3=0
      GET DEFAULT PRINTER TO aIMP2
      aIMP3:=ASCAN(aIMP1, {|aVal| aVal == aIMP2})
   ENDIF
   GET PORTS TO aIMP6
   ASORT(aIMP6,,, { |x, y| UPPER(x) < UPPER(y) })
   AADD(aIMP6,"Fichero")
   RELEASE PRINTSYS

   aIMP4:={"HbPrinter","MiniPrint","PDFPrint","CalcPrint","ExcelPrint","HTMLPrint","CSVPrint","RTFPrint","DosPrint"}
   aIMP5:=1

   aIMP7:={"LPT1","LPT2","COM1","COM2","COM3","COM4","Fichero"}
   aIMP8:=GuardarPuertoMS2("LEER")

   aIMP:={aIMP1,aIMP2,aIMP3,aIMP4,aIMP5,aIMP6,aIMP7,aIMP8}

RETURN(aIMP)

*aIMP1:=array de impresoras
*aIMP2:=nombre de la impresora por defecto
*aIMP3:=numero de la impresora por defecto
*aIMP4:=array de librerias
*aIMP5:=libreria seleccionada por defecto
*aIMP6:=array de puertos programa
*aIMP7:=array de puertos JM
*aIMP8:=numero puerto guardado en INI


FUNCTION GuardarPuertoMS2(LLAMADA,GuaPue2)
   DO CASE
   CASE LLAMADA="LEER"
      BEGIN INI FILE "Escola.ini"
         GuaPue2:=1
         GET GuaPue2 SECTION "Impresora" ENTRY "Puerto_MS2" DEFAULT GuaPue2 //LEER
      END INI
   CASE LLAMADA="GRABAR"
      IF MSGYESNO("¿Desea guardar el puerto de impresora MS2?")=.T.
         BEGIN INI FILE "Escola.ini"
            SET SECTION "Impresora" ENTRY "Puerto_MS2" TO GuaPue2 //GRABAR
         END INI
      ENDIF
   ENDCASE
RETURN(GuaPue2)


*  start doc
*  hbprn:startdoc(TituloImp)

*  set preview scale 2
*  hbprn:PREVIEWSCALE:=2

*para saber la orientacion del papel
*  if hbprn:maxrow>hbprn:maxcol
*     mipageorient:="vertical"
*  else
*     mipageorient:="horizontal"
*  endif
*  msgbox(str(hbprn:maxrow)+HB_OsNewLine()+str(hbprn:maxcol)+HB_OsNewLine()+mipageorient)
