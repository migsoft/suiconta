#include "minigui.ch"

Function RutaProgRed()
   RUTAINI2:=Getfile( {{'Programas EXE','*.exe'}} ,'Buscar programa', RUTAINI , .f. , .t. )
   IF RAT("\",RUTAINI2)<>0
      RUTAINI2:=SUBSTR(RUTAINI2,RAT("\",RUTAINI2)+1,LEN(RUTAINI2))
   ENDIF
   NOMFICINI:=Fichero_Programa("RUTAFICHEROEXE")
   NOMFICINI:=LEFT(NOMFICINI,LEN(NOMFICINI)-4)+".ini"

   IF LEN(RTRIM(RUTAINI2))=0
      RUTAINI2:=RUTAINI
   ENDIF

   IF RUTAINI2==RUTAINI
      MSGSTOP("No se ha modificado la ruta de red"+HB_OsNewLine()+ ;
         "Ruta anterior: "+RUTAINI+HB_OsNewLine()+ ;
         "Nueva ruta: "+RUTAINI2 )
      RETURN Nil
   ELSE
      IF MSGYESNO("¿Desea cambiar la ruta de red?"+HB_OsNewLine()+ ;
         "Ruta anterior: "+RUTAINI+HB_OsNewLine()+ ;
         "Nueva ruta: "+RUTAINI2 )=.F.
         RETURN Nil
      ENDIF
   ENDIF

   BEGIN INI FILE NOMFICINI
      SET SECTION "Rutas" ENTRY "RutaPrograma" TO RUTAINI2 //GRABAR
   END INI

   MSGINFO("Se va ha cerrar el programa"+HB_OsNewLine()+ ;
           "cuando ejecute de nuevo el programa, se actualizara la ruta de red")

   RELEASE WINDOW ALL

Return Nil
