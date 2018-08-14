FUNCTION MES(V_MES1,NIVEL)
LOCAL V_MES2
NIVEL:=IF(NIVEL=NIL,0,NIVEL)
IF VALTYPE(V_MES1)="D"
   V_MES1:=MONTH(V_MES1)
ENDIF

DO CASE
CASE NIVEL=0 .OR. NIVEL=3 //NOMBRE DEL MES
   V_MES3:=V_MES1
   IF V_MES1>12
*      V_MES3:=V_MES1-(INT(V_MES1/12)*12)
      V_MES3:=V_MES1%12
      IF V_MES3=0
         V_MES3=12
      ENDIF
   ENDIF
   DO CASE
   CASE V_MES3=1
   V_MES2="Enero"
   CASE V_MES3=2
   V_MES2="Febrero"
   CASE V_MES3=3
   V_MES2="Marzo"
   CASE V_MES3=4
   V_MES2="Abril"
   CASE V_MES3=5
   V_MES2="Mayo"
   CASE V_MES3=6
   V_MES2="Junio"
   CASE V_MES3=7
   V_MES2="Julio"
   CASE V_MES3=8
   V_MES2="Agosto"
   CASE V_MES3=9
   V_MES2="Septiembre"
   CASE V_MES3=10
   V_MES2="Octubre"
   CASE V_MES3=11
   V_MES2="Noviembre"
   CASE V_MES3=12
   V_MES2="Diciembre"
   OTHERWISE
   V_MES2="ERROR"
   ENDCASE

CASE NIVEL=1 //ARRAY CON NOMBRES DE LOS MESES
   V_MES2={"Enero",;
           "Febrero",;
           "Marzo",;
           "Abril",;
           "Mayo",;
           "Junio",;
           "Julio",;
           "Agosto",;
           "Septiembre",;
           "Octubre",;
           "Noviembre",;
           "Diciembre"}
CASE NIVEL=2 //ARRAY CON NOMBRES DE LOS MESES CORTOS
   V_MES2={"Enero",;
           "Febrero",;
           "Marzo",;
           "Abril",;
           "Mayo",;
           "Junio",;
           "Julio",;
           "Agosto",;
           "Sept.",;
           "Octubre",;
           "Novbre.",;
           "Dicbre."}
ENDCASE

IF NIVEL=3 //NOMBRE DEL MES CON ANCHO FIJO
   V_MES2:=PADR(V_MES2,10," ")
ENDIF

RETURN(V_MES2)
