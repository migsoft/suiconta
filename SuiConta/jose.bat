call \MisProg\brmake\brmake SUICONTA

IF EXIST \SUIZOWIN.RED\SUICONTA\SUICONTA.EXE (
   COPY SUICONTA.EXE \SUIZOWIN.RED\SUICONTA
   ) 
IF EXIST \SUIZOWIN.ET\SUICONTA\SUICONTA.EXE (
   COPY SUICONTA.EXE \SUIZOWIN.ET\SUICONTA
   )

DEL SUICONTA.EXE
