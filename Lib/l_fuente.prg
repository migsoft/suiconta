#ifdef __XHARBOUR__
#  xtranslate GetExeFilename() => ExeName()
#else
#  xtranslate GetExeFilename() => hb_ProgName()
#endif

FUNCTION NomFuente(NomFuente)
   IF AT("MEDIA",UPPER(GetExeFilename()))<>0
      DO CASE
      CASE AT(UPPER('Arial'),UPPER(NomFuente))<>0
         NomFuente:='Free Sans'
      CASE AT(UPPER('Courier'),UPPER(NomFuente))<>0
         NomFuente:='DejaVu Sans Mono'
      CASE AT(UPPER('Times New Roman'),UPPER(NomFuente))<>0
         NomFuente:='FreeSerif'
      ENDCASE
   ENDIF
RETURN(NomFuente)
