#include "minigui.ch"

FUNCTION Menu_Grid(VENTANA,CONTROL,LLAMADA,gORDEN,aOPCIONES,TIPORDEN)
   DEFAULT aOPCIONES TO {"TODO"}
   DEFAULT TIPORDEN TO "ASCENDENTE"
   LLAMADA:=UPPER(LLAMADA)

   IF IsWIndowDefined(&VENTANA)=.F. .OR. IsControlDefined(&CONTROL,&VENTANA)=.F.
      RETURN
   ENDIF

	aRow := GetProperty(VENTANA,CONTROL,"Value")
	IF VALTYPE(aRow)="A"
		nRow := aRow[1]
		nCol := aRow[2]
	ELSE
		nRow := GetProperty(VENTANA,CONTROL,"Value")
		IF VALTYPE(gORDEN)="N"
			nCol := gORDEN
		ELSE
			nCol := 1
		ENDIF
	ENDIF

   IF UPPER(LLAMADA)="MENU"
      DEFINE CONTEXT MENU CONTROL &(CONTROL) OF &(VENTANA)
         SISEPARADOR:=0
         IF ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "DUPLICAR"})<>0 .OR. ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "TODO"})<>0
            ITEM "Duplicar" ACTION Menu_Grid(VENTANA,CONTROL,"Duplicar")
            SISEPARADOR:=1
         ENDIF
         IF ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "ELIMINAR"})<>0 .OR. ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "TODO"})<>0
            ITEM "Eliminar" ACTION Menu_Grid(VENTANA,CONTROL,"Eliminar")
            SISEPARADOR:=1
         ENDIF
         IF SISEPARADOR=1
            SEPARATOR
         ENDIF
         IF ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "COPCELPP"})<>0 .OR. ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "TODO"})<>0
            ITEM "Copiar celda"    ACTION Menu_Grid(VENTANA,CONTROL,"CopiarCelda",1)
         ENDIF
         IF ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "COPREGPP"})<>0 .OR. ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "TODO"})<>0
            ITEM "Copiar registro" ACTION Menu_Grid(VENTANA,CONTROL,"CopiarRegistro")
         ENDIF
         IF ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "COPTABPP"})<>0 .OR. ASCAN(aOPCIONES, {|aVal| UPPER(aVal) = "TODO"})<>0
            ITEM "Copiar tabla"    ACTION Menu_Grid(VENTANA,CONTROL,"CopiarTabla")
         ENDIF
      END MENU
      RETURN
   ENDIF

   IF GetProperty(VENTANA,CONTROL,"ItemCount")=0
      MSGSTOP("La tabla no tiene regitros")
      RETURN
   ENDIF
   IF nRow=0
      MSGSTOP("No se ha seleccionado ningun registro")
      RETURN
   ENDIF


   DO CASE
   CASE UPPER(LLAMADA)="NUEVO" //PENDIENTE DE HACER
      IF MSGYESNO("Desea añadir un registro en blanco")=.T.
**         SetProperty(VENTANA,CONTROL,"AddItem",{"Nuevo","","","","","",""})
         SetProperty(VENTANA,CONTROL,"Value",GetProperty(VENTANA,CONTROL,"ItemCount"))
         DoMethod(VENTANA,CONTROL,"Refresh")
      ENDIF
   CASE UPPER(LLAMADA)="DUPLICAR"
      DO CASE
      CASE nRow=0
         MsgStop("No se ha seleccionado ningun registro","error")
      CASE MSGYESNO("Desea duplicar el registro activo")=.T.
         DoMethod(VENTANA,CONTROL,"AddItem", GetProperty(VENTANA,CONTROL,'Item', nRow ) )
         SetProperty(VENTANA,CONTROL,"Value",GetProperty(VENTANA,CONTROL,"ItemCount"))
         DoMethod(VENTANA,CONTROL,"Refresh")
      ENDCASE
   CASE UPPER(LLAMADA)="ELIMINAR"
      DO CASE
      CASE nRow=0
         MsgStop("No se ha seleccionado ningun registro","error")
      CASE MSGYESNO("Desea eliminar el registro activo")=.T.
         DoMethod(VENTANA,CONTROL,"DeleteItem",nRow)
         DoMethod(VENTANA,CONTROL,"Refresh")
      ENDCASE
   CASE UPPER(LLAMADA)="COPIARCELDA"
      DO CASE
      CASE nRow=0
         MsgStop("No se ha seleccionado ningun registro","error")
      OTHERWISE
         TEXTO2:=""
         aCopiar:=GetProperty(VENTANA,CONTROL,'Item', nRow )
			IF VALTYPE(aRow)="A"
				SelCelda2:={nCol}
			ELSE
	         aTITULOS:={}
	         FOR N=1 TO LEN(aCopiar)
	            AADD(aTITULOS,GetProperty(VENTANA,CONTROL,'Header', N ) )
	         NEXT
	         aLabels    :={'Celda'}
	         aFormats   :={aTITULOS}
	         TIT2:="Copiar celda al portapapeles"
	         aInitValues:={gORDEN}
	         aBOTON:={"Copiar","Cancelar"}
	         SelCelda2:=InputWindow(TIT2, aLabels , aInitValues , aFormats,GetProperty(VENTANA,"Row")+200,GetProperty(VENTANA,"Col")+200 , .T. , aBOTON)
			ENDIF
         IF SelCelda2[1] == Nil
            TEXTO2:=""
         ELSE
            DO CASE
            CASE VALTYPE(aCopiar[SelCelda2[1]])="C"
               TEXTO2:=aCopiar[SelCelda2[1]]
            CASE VALTYPE(aCopiar[SelCelda2[1]])="N"
               TEXTO2:=STR(aCopiar[SelCelda2[1]])
            CASE VALTYPE(aCopiar[SelCelda2[1]])="D"
               TEXTO2:=DTOC(aCopiar[SelCelda2[1]])
            CASE VALTYPE(aCopiar[SelCelda2[1]])="L"
               TEXTO2:=IF(aCopiar[SelCelda2[1]]=.T.,"Si","No")
            OTHERWISE
               MsgStop("Formato desconocido","error")
               TEXTO2:=""
            ENDCASE
         ENDIF
         CopyToClipboard(TEXTO2)
      ENDCASE
   CASE UPPER(LLAMADA)="COPIARREGISTRO"
      DO CASE
      CASE nRow=0
         MsgStop("No se ha seleccionado ningun registro","error")
      CASE MSGYESNO("Desea copiar el registro activo al portapapeles")=.T.
         TEXTO2:=""
         aCopiar:=GetProperty(VENTANA,CONTROL,'Item', nRow )
         TEXTO2:=TextoArray(aCopiar)
         CopyToClipboard(TEXTO2)
      ENDCASE
   CASE UPPER(LLAMADA)="COPIARTABLA"
      IF MSGYESNO("Desea copiar la tabla al portapapeles")=.T.
         PonerEspera("Copiando tabla al portapapeles")
         aCopiar2:=GetProperty(VENTANA,CONTROL,'Item', 1 )
			aCopiar:={}
         TEXTO2:=""
			FOR N=1 TO LEN(aCopiar2)
				AADD(aCopiar, GetProperty(VENTANA,CONTROL,"Header",N))
			NEXT
         TEXTO2:=TEXTO2+TextoArray(aCopiar)+HB_OsNewLine()
         FOR N2=1 TO GetProperty(VENTANA,CONTROL,"ItemCount")
            DO EVENTS
            aCopiar:=GetProperty(VENTANA,CONTROL,'Item', N2 )
            TEXTO2:=TEXTO2+TextoArray(aCopiar)+HB_OsNewLine()
         NEXT
         CopyToClipboard(TEXTO2)
         QuitarEspera()
      ENDIF
   CASE UPPER(LLAMADA)="ORDENAR"
      aOrdenar1:=GetProperty(VENTANA,CONTROL,'Item', nRow )
      IF VALTYPE(aOrdenar1[gORDEN])<>"C" .AND. ;
         VALTYPE(aOrdenar1[gORDEN])<>"N" .AND. ;
         VALTYPE(aOrdenar1[gORDEN])<>"D"
         MSGSTOP("Esta columna no se puede ordenar")
         RETURN
      ENDIF
      PonerEspera("Ordenando la tabla por "+GetProperty(VENTANA,CONTROL,"Header",gORDEN ) )

      TESTOY:=TextoArray(aOrdenar1)
      NESTOY:=0
      aOrdenar2:={}
      FOR N=1 TO GetProperty(VENTANA,CONTROL,"ItemCount")
         AADD(aOrdenar2,GetProperty(VENTANA,CONTROL,'Item',N) )
      NEXT
      DO CASE
      CASE VALTYPE(aOrdenar2[1,gORDEN])="C"
			IF TIPORDEN="ASC"
	         ASORT(aOrdenar2,,, { |x, y| UPPER(x[gORDEN]) < UPPER(y[gORDEN]) })
			ELSE
	         ASORT(aOrdenar2,,, { |x, y| UPPER(x[gORDEN]) > UPPER(y[gORDEN]) })
			ENDIF
      CASE VALTYPE(aOrdenar2[1,gORDEN])="N"
			IF TIPORDEN="ASC"
	         ASORT(aOrdenar2,,, { |x, y| x[gORDEN] < y[gORDEN] })
			ELSE
	         ASORT(aOrdenar2,,, { |x, y| x[gORDEN] > y[gORDEN] })
			ENDIF
      CASE VALTYPE(aOrdenar2[1,gORDEN])="D"
			IF TIPORDEN="ASC"
	         ASORT(aOrdenar2,,, { |x, y| DTOS(x[gORDEN]) < DTOS(y[gORDEN]) })
			ELSE
	         ASORT(aOrdenar2,,, { |x, y| DTOS(x[gORDEN]) > DTOS(y[gORDEN]) })
			ENDIF
      ENDCASE
      DoMethod(VENTANA,CONTROL,"DeleteAllItems")
      FOR N=1 TO LEN(aOrdenar2)
         DoMethod(VENTANA,CONTROL,"AddItem", aOrdenar2[N] )
         IF NESTOY=0
            IF TESTOY=TextoArray(aOrdenar2[N])
               NESTOY:=N
            ENDIF
         ENDIF
      NEXT
      SetProperty(VENTANA,CONTROL,"Value",IF(NESTOY=0,1,NESTOY))
      DoMethod(VENTANA,CONTROL,"Refresh")
      QuitarEspera()


   ENDCASE



FUNCTION TextoArray(aTextoArray)
   TTextoArray:=""
   FOR N1=1 TO LEN(aTextoArray)
      DO CASE
      CASE VALTYPE(aTextoArray[N1])="C"
         TTextoArray:=TTextoArray+aTextoArray[N1]
      CASE VALTYPE(aTextoArray[N1])="N"
         TTextoArray:=TTextoArray+STRTRAN(STR(aTextoArray[N1]),".",",")
      CASE VALTYPE(aTextoArray[N1])="D"
         TTextoArray:=TTextoArray+DTOC(aTextoArray[N1])
      CASE VALTYPE(aTextoArray[N1])="L"
         TTextoArray:=TTextoArray+IF(aTextoArray[N1]=.T.,"Si","No")
      ENDCASE
      IF N1<LEN(aTextoArray)
         TTextoArray:=TTextoArray+CHR(9)
      ENDIF
   NEXT
RETURN(TTextoArray)



***29-01-2009 Grigory Filatov***
Function GetGridColumnsCount(ControlName,ParentForm)
   Local i, nColumn

   i := GetControlIndex(ControlName,ParentForm)
   //*** nColumn := Len( _HMG_aControlCaption[i] )
   nColumn := DoMethod( ParentForm, ControlName, 'ColumnCount' )
   // nColumn := oGrid:ColumnCount()

   Return(nColumn)


FUNCTION GetControlIndex ( ControlName, ParentForm )
*-----------------------------------------------------------------------------*
   LOCAL mVar := '_' + ParentForm + '_' + ControlName

   IF __mvExist ( mVar )
      RETURN ( &mVar )
   ENDIF

RETURN 0


