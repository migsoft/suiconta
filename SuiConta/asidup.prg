#include "minigui.ch"

Function Asi_Duplicar(RUTA1,NASIDUP1,NUMEMP1)
   NUMEMP1:=IF(VALTYPE(NUMEMP1)="N",STR(NUMEMP1),NUMEMP1)
   aAPUNTE:={}
   NOMEMP1:=""
   CUADRE2:=0
   NASIDUP2:=0
   DO CASE
   CASE AT("SUICONTA",UPPER(RUTA1))<>0
      IF AT("SUIZO",UPPER(RUTA1))<>0 .AND. VAL(NUMEMP1)<>0
         RUTA2:=LEFT(RUTA1, RAT("SUIZO",UPPER(RUTA1))-2 )
      ELSE
         RUTA2:=RUTA1
      ENDIF
      IF FILE(RUTA2+"\EMPRESA.DBF") .AND. ;
         FILE(RUTA2+"\EMPRESA.CDX")
         AbrirDBF("EMPRESA",,,,RUTA2)
         SEEK VAL(NUMEMP1)
         RUTA2:=RUTA2+"\"+RTRIM(RUTA)
         NOMEMP1:=RTRIM(EMP)+" "+LTRIM(STR(EJERCICIO))
         EMPRESA->( DBCLOSEAREA() )
      ENDIF
      IF FILE(RUTA2+"\APUNTES.DBF") .AND. ;
         FILE(RUTA2+"\APUNTES.CDX")
         AbrirDBF("APUNTES",,,,RUTA2)
         SEEK NASIDUP1
         DO WHILE NASI=NASIDUP1 .AND. .NOT. EOF()
            AADD(aAPUNTE,{STR(NASI),DIA(FECHA,10),STR(CODCTA),"",RTRIM(NOMAPU),DEBE,HABER})
            CUADRE2:=CUADRE2+DEBE-HABER
            SKIP
         ENDDO
         APUNTES->( DBCLOSEAREA() )
         IF FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
            FILE(RUTA2+"\CUENTAS.CDX")
            AbrirDBF("CUENTAS",,,,RUTA2)
            FOR NA=1 TO LEN(aAPUNTE)
               SEEK VAL(aAPUNTE[NA,3])
               IF .NOT. EOF()
                  aAPUNTE[NA,4]:=RTRIM(NOMCTA)
               ENDIF
            NEXT
            CUENTAS->( DBCLOSEAREA() )
         ENDIF
      ELSE
         RETURN(0)
      ENDIF
   CASE AT("GRUPOSP",UPPER(RUTA1))<>0
      RUTA2:=LEFT(RUTA1,RAT("EMP",UPPER(RUTA1))-2)
      RUTA2:=RUTA2+"\EMP"
      IF FILE(RUTA2+"\EMPRESA.DBF") .AND. ;
         FILE(RUTA2+"\EMPRESA.CDX")
         AbrirDBF("EMPRESA",,,,RUTA2)
         DBSETORDER(1)
         NUMEMP1:=PADR(NUMEMP1,LEN(COD)," ")
         SEEK NUMEMP1
         IF EOF() .AND. VAL(NUMEMP1)<>0
            LOCATE FOR VAL(COD)=VAL(NUMEMP1)
            NUMEMP1:=COD
         ENDIF
         NOMEMP1:=RTRIM(NOMBRE)+" "+LTRIM(EJERCICIO)
         EMPRESA->( DBCLOSEAREA() )
      ENDIF
      RUTA2:=RUTA2+NUMEMP1
      IF FILE(RUTA2+"\DIARIO.DBF") .AND. ;
         FILE(RUTA2+"\DIARIO.CDX")
         AbrirDBF("DIARIO",,,,RUTA2)
         DBSETORDER(3)
         SEEK NASIDUP1
         DO WHILE ASIEN=NASIDUP1 .AND. .NOT. EOF()
            IF DELETE()=.F.
               AADD(aAPUNTE,{LTRIM(STR(ASIEN)),DIA(FECHA,10),SUBCTA,"",RTRIM(CONCEPTO),EURODEBE,EUROHABER})
               CUADRE2:=CUADRE2+EURODEBE-EUROHABER
            ENDIF
            SKIP
         ENDDO
         DIARIO->( DBCLOSEAREA() )
         IF FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
            FILE(RUTA2+"\SUBCTA.CDX")
            AbrirDBF("SUBCTA",,,,RUTA2)
            DBSETORDER(1)
            FOR NA=1 TO LEN(aAPUNTE)
               SEEK PADR(LTRIM(aAPUNTE[NA,3]),LEN(COD)," ")
               IF .NOT. EOF()
                  aAPUNTE[NA,4]:=RTRIM(TITULO)
               ENDIF
            NEXT
            SUBCTA->( DBCLOSEAREA() )
         ENDIF
      ELSE
         RETURN(0)
      ENDIF
   OTHERWISE
      RETURN(0)
   ENDCASE

   IF LEN(aAPUNTE)=0
      RETURN(0)
   ENDIF



   DEFINE WINDOW WinAsiDup ;
      AT 0,0     ;
      WIDTH 750  ;
      HEIGHT 450 ;
      TITLE "Duplicar asiento" ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("ROJOCLARO")

   @ 10,10 LABEL L_Empresa VALUE NOMEMP1 AUTOSIZE TRANSPARENT
   @ 10,300 LABEL L_Ruta VALUE RUTA2 AUTOSIZE TRANSPARENT

   @ 45,10 LABEL L_Asiento VALUE "Asiento" AUTOSIZE TRANSPARENT
   @ 40,60 TEXTBOX T_Asiento WIDTH 100 VALUE VAL(aAPUNTE[1,1]) ;
           NUMERIC INPUTMASK '999,999,999,999' FORMAT 'E' RIGHTALIGN

   @ 45,250 LABEL L_Fecha VALUE "Fecha" AUTOSIZE TRANSPARENT
   @ 40,300 DATEPICKER D_Fecha WIDTH 100 VALUE CTOD(aAPUNTE[1,2]) NOTABSTOP

   @ 45,500 LABEL L_Cuadre VALUE "Cuadre" AUTOSIZE TRANSPARENT
   @ 40,550 TEXTBOX T_Cuadre WIDTH 110 VALUE 0 READONLY NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN

/*
      @ 70,10 GRID G_Asiento ;
      HEIGHT 300 ;
      WIDTH 710 ;
      HEADERS {,,'Cuenta','Descripcion cuenta','Descripcion apunte','Debe','Haber'} ;
      WIDTHS {0,0,80,150,250,100,100 } ;
      ITEMS aAPUNTE ;
      JUSTIFY {,,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE 1 ;
      EDIT INPLACE {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
      COLUMNVALID { ,, { || ActAsiDup("Cuenta")},,{ || ActAsiDup("Descripcion")},{ || ActAsiDup("Importe")},{ || ActAsiDup("Importe")}} ;
 */

      @ 70,10 GRID G_Asiento ;
      HEIGHT 300 ;
      WIDTH 710 ;
      HEADERS {,,'Cuenta','Descripcion cuenta','Descripcion apunte','Debe','Haber'} ;
      WIDTHS {0,0,80,150,250,100,100 } ;
      ITEMS aAPUNTE ;
      JUSTIFY {,,BROWSE_JTFY_CENTER,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT} ;
      VALUE 1 ;
      INPLACE ;
      COLUMNCONTROLS {{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}} ;
      VALID { ,, { || ActAsiDup("Cuenta")},,{ || ActAsiDup("Descripcion")},{ || ActAsiDup("Importe")},{ || ActAsiDup("Importe")}} ;

      @380,530 BUTTONEX Bt_Duplicar ;
         CAPTION 'Duplicar' ;
         ICON icobus('guardar') ;
         ACTION Asi_DuplicarG(RUTA2,NASIDUP1,NUMEMP1) ;
         WIDTH 90 HEIGHT 25 ;
         NOTABSTOP

      @380,630 BUTTONEX Bt_Salir ;
         CAPTION 'Salir' ;
         ICON icobus('salir') ;
         ACTION WinAsiDup.Release ;
         WIDTH 90 HEIGHT 25 ;
         NOTABSTOP

   WinAsiDup.T_Asiento.Enabled:= .F.
   WinAsiDup.G_Asiento.SetFocus

   END WINDOW
   VentanaCentrar("WinAsiDup","Ventana1")
   ACTIVATE WINDOW WinAsiDup

RETURN(NASIDUP2)


STATIC FUNCTION ActAsiDup(LLAMADA)
local nVal := This.CellValue
local nCol := This.CellColIndex
local nRow := This.CellRowIndex
DO CASE
CASE LLAMADA="Cuenta"
   IF FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
      FILE(RUTA2+"\CUENTAS.CDX")
      AbrirDBF("CUENTAS",,,,RUTA2)
      SEEK VAL(PCODCTA3(nVal))
      IF .NOT. EOF()
         WinAsiDup.G_Asiento.Cell(nRow,4):=RTRIM(NOMCTA)
      ELSE
         WinAsiDup.G_Asiento.Cell(nRow,4):=""
      ENDIF
   ENDIF
   IF FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
      FILE(RUTA2+"\SUBCTA.CDX")
      AbrirDBF("SUBCTA",,,,RUTA2)
      DBSETORDER(3)
      SEEK VAL(PCODCTA3(nVal))
      IF .NOT. EOF()
         WinAsiDup.G_Asiento.Cell(nRow,4):=RTRIM(TITULO)
      ELSE
         WinAsiDup.G_Asiento.Cell(nRow,4):=""
      ENDIF
   ENDIF
CASE LLAMADA="Descripcion"
   IF MSGYESNO("¿Desea copiar la descripion en todo el asiento?")
      FOR N=1 TO WinAsiDup.G_Asiento.ItemCount
         WinAsiDup.G_Asiento.Cell(N,5):=nVal
      NEXT
   ENDIF
CASE LLAMADA="Importe"
   IF WinAsiDup.G_Asiento.Cell(nRow,6)<>0 .AND. WinAsiDup.G_Asiento.Cell(nRow,7)<>0
      IF nCol=6
         WinAsiDup.G_Asiento.Cell(nRow,7):=0
      ELSE
         WinAsiDup.G_Asiento.Cell(nRow,6):=0
      ENDIF
   ENDIF
   SALDO2:=0
   FOR N=1 TO WinAsiDup.G_Asiento.ItemCount
      IF N=nRow
         WinAsiDup.G_Asiento.Cell(N,6):=0
         WinAsiDup.G_Asiento.Cell(N,7):=0
         SALDO2:=SALDO2+IF(nCol=6,nVal,nVal*-1)
      ELSE
         SALDO2:=SALDO2+WinAsiDup.G_Asiento.Cell(N,6)-WinAsiDup.G_Asiento.Cell(N,7)
      ENDIF
   NEXT
   WinAsiDup.T_Cuadre.Value:=SALDO2
   IF SALDO2=0
      WinAsiDup.T_Cuadre.BackColor:=MICOLOR("VERDECLARO")
   ELSE
      WinAsiDup.T_Cuadre.BackColor:=MICOLOR("ROJOCLARO")
   ENDIF
ENDCASE


STATIC FUNCTION Asi_DuplicarG(RUTA2b,NASIDUP1,NUMEMP1)
   NUMEMP1:=IF(VALTYPE(NUMEMP1)="N",STR(NUMEMP1),NUMEMP1)
   RUTA2:=RUTA2b
   IF MSGYESNO("¿Desea duplicar el asiento?")=.F.
      RETURN
   ENDIF

   PonerEspera("Duplicando asiento....")

   IF FILE(RUTA2+"\APUNTES.DBF") .AND. ;
      FILE(RUTA2+"\APUNTES.CDX") .AND. ;
      FILE(RUTA2+"\CUENTAS.DBF") .AND. ;
      FILE(RUTA2+"\CUENTAS.CDX")
      AbrirDBF("CUENTAS",,,,RUTA2)
      AbrirDBF("APUNTES",,,,RUTA2)
      GO BOTT
      IF .NOT. EOF()
         NASIDUP2:=NASI+1
      ELSE
         NASIDUP2:=2
      ENDIF
      FOR N=1 TO WinAsiDup.G_Asiento.ItemCount
         APPEND BLANK
         IF RLOCK()
            REPLACE NASI WITH NASIDUP2
            REPLACE APU WITH N
            REPLACE NEMP WITH VAL(NUMEMP1)
            REPLACE FECHA WITH WinAsiDup.D_Fecha.Value
            REPLACE CODCTA WITH VAL(PCODCTA3(WinAsiDup.G_Asiento.Cell(N,3)))
            REPLACE NOMAPU WITH WinAsiDup.G_Asiento.Cell(N,5)
            REPLACE DEBE WITH WinAsiDup.G_Asiento.Cell(N,6)
            REPLACE HABER WITH WinAsiDup.G_Asiento.Cell(N,7)
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
         ENDIF
      NEXT
      APUNTES->( DBCLOSEAREA() )
   ENDIF
   IF FILE(RUTA2+"\DIARIO.DBF") .AND. ;
      FILE(RUTA2+"\DIARIO.CDX") .AND. ;
      FILE(RUTA2+"\SUBCTA.DBF") .AND. ;
      FILE(RUTA2+"\SUBCTA.CDX")
      AbrirDBF("SUBCTA",,,,RUTA2)
      AbrirDBF("DIARIO",,,,RUTA2)
      DBSETORDER(3)
      IF .NOT. EOF()
         NASIDUP2:=NASI+1
      ELSE
         NASIDUP2:=2
      ENDIF
      FOR N=1 TO WinAsiDup.G_Asiento.ItemCount
         APPEND BLANK
         IF RLOCK()
            REPLACE ASIEN WITH NASIDUP2
            REPLACE FECHA WITH WinAsiDup.D_Fecha.Value
            REPLACE SUBCTA WITH PCODCTA3(WinAsiDup.G_Asiento.Cell(N,3))
            REPLACE CONCEPTO WITH WinAsiDup.G_Asiento.Cell(N,5)
            REPLACE PTADEBE  WITH MDA_PTA(WinAsiDup.G_Asiento.Cell(N,6),YEAR(FECHA))
            REPLACE PTAHABER WITH MDA_PTA(WinAsiDup.G_Asiento.Cell(N,7),YEAR(FECHA))
            REPLACE EURODEBE  WITH MDA_EURO(WinAsiDup.G_Asiento.Cell(N,6),YEAR(FECHA))
            REPLACE EUROHABER WITH MDA_EURO(WinAsiDup.G_Asiento.Cell(N,7),YEAR(FECHA))
            REPLACE MONEDAUSO WITH IF(YEAR(FECHA)>=2002,"2","1")
            REPLACE NOCONV WITH .F.
            DBCOMMIT()
            DBUNLOCK()
            GrupoSP_saldocuenta(SUBCTA,EURODEBE,EUROHABER,FECHA)
         ENDIF
      NEXT
      DIARIO->( DBCLOSEAREA() )
   ENDIF

   QuitarEspera()

MSGBOX("Asiento duplicado"+HB_OsNewLine()+"Asiento: "+LTRIM(STR(NASIDUP2))+HB_OsNewLine()+"Fecha: "+DIA(WinAsiDup.D_Fecha.Value))


