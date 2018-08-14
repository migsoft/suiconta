FUNCTION Ejer_Cerrado(EJERCE2,NIVEL)
   IF VALTYPE(EJERCE2)="N"
      IF EJERCE2=1
         EJERCE2:=.T.
      ELSE
         EJERCE2:=.F.
      ENDIF
   ENDIF
   IF EJERCE2=.T.
      IF NIVEL="VER"
         RESPEJER:=MSGYESNO("EL EJERCICIO ESTA CERRADO"+HB_OsNewLine()+ ;
            "¿Desea continuar?","¡ATENCION!")
      ELSE //"GUARDAR"
         RESPEJER:=MSGYESNO("EL EJERCICIO ESTA CERRADO"+HB_OsNewLine()+ ;
            "¿Desea guardar los datos?","¡ATENCION!")
      ENDIF
      IF RESPEJER=.F.
         RETURN .F.
      ELSE
         RETURN .T.
      ENDIF
   ELSE
      RETURN .T.
   ENDIF
