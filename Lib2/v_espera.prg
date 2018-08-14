FUNCTION ESPERA(ESP1,ESP2,OTRATECLA2)
   OTRATECLA2:=IF(OTRATECLA2=NIL,0,OTRATECLA2)
   PAGACT:=ESP1
@ 24,0  SAY SPACE(80) COLOR COLORRAT
@ 24,1  SAY "<Re.P g> RETRO.PAG." COLOR COLORRAT
@ 24,28 SAY "<Av.P g> AVANZAR PAG." COLOR COLORRAT
@ 24,55 SAY "<0> FIN" COLOR COLORRAT
@ 24,69 SAY PADL("HOJA: "+LTRIM(STR(ESP1)),10," ") COLOR COLORRAT
IF ESP1=1
   PUBLIC ESPERA:=ARRAY(1)
   ESPERA[1]:=SAVESCREEN(0,0,24,79)
ELSE
   IF ESP1>150
      ESPERA[ESP1-(150*INT(ESP1/150))+1]:=SAVESCREEN(0,0,24,79)
   ELSE
      AADD(ESPERA,SAVESCREEN(0,0,24,79))
   ENDIF
ENDIF
DO WHILE .T.
   TECLAESP:=INKEY(0)
   DO CASE
      CASE TECLAESP=OTRATECLA2 .AND. OTRATECLA2<>0
         EXIT
      CASE TECLAESP=48 .OR. TECLAESP=27 .OR. ;
           TECLAESP=1003 .AND. MROW()=24 .AND. MCOL()>=55 .AND. MCOL()<=61
         EXIT
      CASE TECLAESP=3 .OR. ;
           TECLAESP=1003 .AND. MROW()=24 .AND. MCOL()>=21 .AND. MCOL()<=41
         IF ESP1=ESP2
            IF ESP1<PAGACT
               ESP1=ESP1+1
               IF ESP1>150
                  RESTSCREEN(0,0,24,79,ESPERA[ESP1-(150*INT(ESP1/150))+1])
               ELSE
                  RESTSCREEN(0,0,24,79,ESPERA[ESP1])
               ENDIF
            ENDIF
         ELSE
            IF ESP1=PAGACT
               EXIT
            ELSE
               ESP1=ESP1+1
               IF ESP1>150
                  RESTSCREEN(0,0,24,79,ESPERA[ESP1-(150*INT(ESP1/150))+1])
               ELSE
                  RESTSCREEN(0,0,24,79,ESPERA[ESP1])
               ENDIF
            ENDIF
         ENDIF
      CASE TECLAESP=18 .AND. ESP1>1 .OR. ESP1>1 .AND. ;
           TECLAESP=1003 .AND. MROW()=24 .AND. MCOL()>=1  .AND. MCOL()<=19
         IF PAGACT-ESP1=148
            SETCOLOR(COLORERROR)
            @ 10,20 CLEAR TO 14,60
            @ 10,20 TO 14,60 DOUBLE
            @ 11,25 SAY "! ! !  A T E N C I O N  ! ! !"
            @ 12,25 SAY "  SOLO SE PUEDEN CONSULTAR"
            @ 13,25 SAY "  LAS ULTIMAS 150 PAGINAS"
            INKEY(0)
            @ 24,79 SAY ""
            SETCOLOR(MICOLOR)
         ELSE
            ESP1=ESP1-1
         ENDIF
         IF ESP1>150
            RESTSCREEN(0,0,24,79,ESPERA[ESP1-(150*INT(ESP1/150))+1])
         ELSE
            RESTSCREEN(0,0,24,79,ESPERA[ESP1])
         ENDIF
   ENDCASE
ENDDO
RETURN
