FUNCTION LIM_CAM(LLAMADA)
LOCAL TECLA,NUM2
LLAMADA:=IF(LLAMADA=NIL,"DOS","WINDOWS")
ELEMENU1:={" #  ALMUADILLA  # ", ;
           " ; PUNTO Y COMA ; ", ;
           " ,     COMA     , ", ;
           " *  ASTERISCO   * ", ;
           "     ESPACIO      ", ;
           "    TABULACION    "}
IF LLAMADA="DOS"
   PANLCAM:=SAVESCREEN(0,0,24,79)
   SETCOLOR(COLORCU2)
   @ 14,25 CLEAR TO 22,55
   @ 14,25 TO 22,55 DOUBLE
   @ 15,30 SAY "LIMITADOR DE CAMPO:"
   RESP1:=ACHOICE(16,30,21,40,ELEMENU1)
   RESTSCREEN(0,0,24,79,PANLCAM)
ELSE
   TIT2:="Limitador de campo"
   aLabels    :={'Limitador'}
   aInitValues:={1}
   aFormats   :={ELEMENU1}
   aBOTON:={"Aceptar","Cancelar"}
   AltaPag2:=InputWindow(TIT2, aLabels , aInitValues , aFormats, 260,410 , .T. , aBOTON) 
   IF AltaPag2[1] == Nil
      RESP1:=1
   ELSE
      RESP1:=AltaPag2[1]
   ENDIF

ENDIF
   DO CASE
   CASE RESP1<=1
      NUM2:=35 //ALMUADILLA
   CASE RESP1=2
      NUM2:=59 //PUNTO Y COMA
   CASE RESP1=3
      NUM2:=44 //COMA
   CASE RESP1=4
      NUM2:=42 //ASTERISCO
   CASE RESP1=5
      NUM2:=32 //ESPACIO
   CASE RESP1=6
      NUM2:=9  //TABULACION
   ENDCASE
   SETCOLOR(MICOLOR)
   RETURN(NUM2)

