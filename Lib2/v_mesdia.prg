FUNCTION MESDIA(V_MES1,V_DIA1,V_DIA2)
IF V_DIA1<1 .OR. V_DIA1>31
IF V_DIA2<1 .OR. V_DIA2>31
   RETURN(V_MES1)
ENDIF
ENDIF
DO WHILE DAY(V_MES1)<>V_DIA1 .AND. DAY(V_MES1)<>V_DIA2
   VD1:=DAY(V_MES1)
   VD2:=MONTH(V_MES1)
   V_MES1=V_MES1+1
   IF MONTH(V_MES1)<>VD2 .AND. VD1<V_DIA1 .OR. ;
      MONTH(V_MES1)<>VD2 .AND. VD1<V_DIA2
      V_MES1=V_MES1-1
      EXIT
   ENDIF
ENDDO
RETURN(V_MES1)
