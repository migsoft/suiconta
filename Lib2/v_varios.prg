FUNCTION INSERTAR()
   IF READINSERT()
      READINSERT(.F.)
      SETCURSOR(3)
      @ 0,64 SAY "   " COLOR COLORCAB
   ELSE
      READINSERT(.T.)
      SETCURSOR(1)
      @ 0,64 SAY "Ins" COLOR COLORCAB
   ENDIF
   RETURN
*
FUNCTION MICURSOR(VER)
   IF VER=1
      IF(READINSERT(),SETCURSOR(1),SETCURSOR(3))
   ELSE
      SETCURSOR(0)
   ENDIF
   RETURN
*
FUNCTION NOLISTADO()
   PANTNL:=SAVESCREEN(0,0,24,79)
   COLORNL:=SETCOLOR()
   SETCOLOR(COLORERROR)
   @ 20,20 CLEAR TO 23,61
   @ 20,20 TO 23,61
   @ 21,28 SAY "NO SE HAN LOCALIZADO DATOS" COLOR COLORERROR+"*"
   INKEY(0) //CONTINUA(1,22,24)
   SETCOLOR(COLORNL)
   RESTSCREEN(0,0,24,79,PANTNL)
   RETURN
*
FUNCTION MENUCERO()
   KEYBOARD CHR(27)
*
FUNCTION CAMPO(NOMBRE)
   IF FIELDPOS(NOMBRE)<>0
      RETURN(.T.)
   ELSE
      RETURN(.F.)
   ENDIF
*
FUNCTION MIPADL(MIPAD2)
   MIPAD2:=PADL(RTRIM(MIPAD2),LEN(MIPAD2)," ")
   RETURN(MIPAD2)
*
FUNCTION MIPADR(MIPAD2)
   MIPAD2:=PADR(LTRIM(MIPAD2),LEN(MIPAD2)," ")
   RETURN(MIPAD2)
*
FUNCTION ESNUMERO(ESNUM2)
   IF VALTYPE(ESNUM2)="N"
      RETURN(.T.)
   ENDIF
   IF VALTYPE(ESNUM2)<>"C"
      RETURN(.F.)
   ENDIF
   ESNUM2:=ALLTRIM(ESNUM2)
   FOR ESN2=1 TO LEN(ESNUM2)
      IF .NOT. ISDIGIT(SUBSTR(ESNUM2,ESN2,1)) ;
         .AND. SUBSTR(ESNUM2,ESN2,1)<>"." .AND. SUBSTR(ESNUM2,ESN2,1)<>"," ;
         .AND. SUBSTR(ESNUM2,ESN2,1)<>"-" .AND. SUBSTR(ESNUM2,ESN2,1)<>"+"
         RETURN(.F.)
      ENDIF
   NEXT
   RETURN(.T.)
*
