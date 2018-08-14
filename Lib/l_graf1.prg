/*
 * MINIGUI - Harbour Win32 GUI library Demo
*/

#include "minigui.ch"

Function Graf1(aSer,aTit,NomSer,NomTit,aColor)
    Local Serie

    IF aColor=NIL
        aColor:={}
        FOR nCOLOR=1 TO LEN(aNomSer)
            IF VAL(aNomSer[nCOLOR])<>0
                nCOLOR2:=VAL(aNomSer[nCOLOR])
            ELSE
                nCOLOR2:=nCOLOR
            ENDIF
            DO CASE
            CASE nCOLOR2%10=1
                AADD(aColor,MICOLOR("ROJO"))
            CASE nCOLOR2%10=2
                AADD(aColor,MICOLOR("VERDE"))
            CASE nCOLOR2%10=3
                AADD(aColor,MICOLOR("AZUL"))
            CASE nCOLOR2%10=4
                AADD(aColor,MICOLOR("AMARILLO"))
            CASE nCOLOR2%10=5
                AADD(aColor,MICOLOR("VIOLETA"))
            CASE nCOLOR2%10=6
                AADD(aColor,MICOLOR("NARANJA"))
            CASE nCOLOR2%10=7
                AADD(aColor,MICOLOR("ROSA"))
            CASE nCOLOR2%10=8
                AADD(aColor,MICOLOR("MARRONCLARO"))
            CASE nCOLOR2%10=9
                AADD(aColor,MICOLOR("VERDEBOSQUE"))
            CASE nCOLOR2%10=0
                AADD(aColor,MICOLOR("CIAN"))
            ENDCASE
        NEXT
    ENDIF


    For each serie in aSer
        For i:=1 To Len(serie)
            serie[i] := aSer[1][i] / 1000
        Next
    Next

    Define Window GraphTest ;
        At 0,0 OBJ oWinGraph;
        Width 800 ;
        Height 600 ;
        Title NomTit ;
        CHILD        ;                              //MAIN-CHILD, MODAL, SPLITCHILD or TOPMOST
        NOSIZE


        @535,10 RADIOGROUP R_Tipo ;
            OPTIONS { 'Barras' , 'Lineas' , 'Puntos' } ;
            VALUE 1 ;
            WIDTH 75 ;
            TOOLTIP 'Tipo de grafico' ;
            ON CHANGE DrawTipoGraph( aSer,aTit,NomSer,NomTit,aColor ) ;
            HORIZONTAL

        @535,410 BUTTON B_Imp CAPTION 'Imprimir' ;
            ACTION DrawTipoGraphi( aSer,aTit,NomSer,NomTit,aColor ) ;
            WIDTH 100 HEIGHT 25 ;

        @535,610 BUTTON B_Can CAPTION 'Salir' ACTION GraphTest.release ;
            WIDTH 100 HEIGHT 25 ;


        DrawTipoGraph( aSer,aTit,NomSer,NomTit,aColor )

    END WINDOW
    VentanaCentrar("GraphTest","Ventana1")
    CENTER WINDOW GraphTest
    ACTIVATE WINDOW GraphTest

Return

Procedure DrawTipoGraph( aSer,aTit,NomSer,NomTit,aColor, lPrn )
    DO CASE
    CASE GraphTest.R_Tipo.Value=1
         MakePic( "GraphTest","BARS",aSer,aTit,NomSer,NomTit,aColor )
    CASE GraphTest.R_Tipo.Value=2
         MakePic( "GraphTest","LINES",aSer,aTit,NomSer,NomTit,aColor )
    CASE GraphTest.R_Tipo.Value=3
         MakePic( "GraphTest","POINTS",aSer,aTit,NomSer,NomTit,aColor )
    ENDCASE
Return

Procedure DrawTipoGraphi( aSer,aTit,NomSer,NomTit,aColor )
    DO CASE
    CASE GraphTest.R_Tipo.Value=1
         MakeWin( "BARS", aSer,aTit,NomSer,NomTit,aColor )
    CASE GraphTest.R_Tipo.Value=2
         MakeWin( "LINES", aSer,aTit,NomSer,NomTit,aColor )
    CASE GraphTest.R_Tipo.Value=3
         MakeWin( "POINTS", aSer,aTit,NomSer,NomTit,aColor )
    ENDCASE
Return

Procedure MakePic( cForm,cType,aSer,aTit,NomSer,NomTit,aColor )

        ERASE WINDOW &cForm

        DRAW GRAPH           ;
            IN WINDOW &cForm ;
            AT 20,20     ;
            TO 520,770   ;
            TITLE NomTit ;
            TYPE &cType   ;
            SERIES aSer  ;
            YVALUES aTit ;
            DEPTH 15     ;
            BARWIDTH 15  ;
            HVALUES LEN(aSer[1]) ;
            SERIENAMES NomSer    ;
            COLORS aColor ;
            3DVIEW      ;
            SHOWGRID    ;
            SHOWXVALUES ;
            SHOWYVALUES ;
            SHOWLEGENDS ;
            NOBORDER

Return

Procedure MakeWin( cType, aSer,aTit,NomSer,NomTit,aColor )
    Local o_Graph, cForm := "GraphTest"

        Define Window _Graph OBJ o_Graph;
            At GetProperty( cForm, 'Row' ), GetProperty( cForm, 'Col' );
            Width GetProperty( cForm, 'Width' );
            Height GetProperty( cForm, 'Height' );
            Child ;
            BackColor WHITE ;
            ON INIT ( o_Graph:Print() , o_Graph:Release )

            MakePic( "_Graph",cType,aSer,aTit,NomSer,NomTit,aColor )

        End Window

        Activate Window _Graph

Return

