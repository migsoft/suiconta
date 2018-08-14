#include "minigui.ch"

Function br_rem1(Srem1,Nrem1,Ventana1,ControlSrem,ControlNrem)

   RUTAREMESA:=RUTAEMPRESA
   RUTAREMEMP:=RUTAPROGRAMA

   aREMESAS:=Remesa_alis(RUTAREMESA)
   IRNUM2:=ASCAN(aREMESAS,{|AVAL| AVAL[1]=Srem1 .AND. AVAL[2]=Nrem1})
   IRNUM2:=IF(IRNUM2=0,1,IRNUM2)

   DEFINE WINDOW W_Imp1 ;
      AT 0,0     ;
      WIDTH 480 HEIGHT 400 ;
      TITLE 'Remesa' ;
      MODAL      ;
      NOSIZE BACKCOLOR MiColor("AZULCLARO")

      @10,10 GRID GR_Remesas ;
      HEIGHT 350 ;
      WIDTH 460 ;
      HEADERS {'Serie','Numero','Remesa','Fecha','Banco','Importe','Documentos','Asiento'} ;
      WIDTHS {0,0,75,80,75,100,50,50 } ;
      ITEMS aREMESAS ;
      VALUE IRNUM2 ;
      ON DBLCLICK br_rem1_Terminar(Srem1,Nrem1,Ventana1,ControlSrem,ControlNrem) ;
      COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}, ;
         {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}, ;
         {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
         {'TEXTBOX','NUMERIC','9,999,999','E'},{'TEXTBOX','NUMERIC'}} ;
      ON HEADCLICK {{|| Remesa_Grid_Ord(1)},{|| Remesa_Grid_Ord(2)},{|| Remesa_Grid_Ord(3)}, ;
                    {|| Remesa_Grid_Ord(4)},{|| Remesa_Grid_Ord(5)},{|| Remesa_Grid_Ord(6)}, ;
                    {|| Remesa_Grid_Ord(7)},{|| Remesa_Grid_Ord(8)}} ;
      JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
               BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT}

      Menu_Grid("W_Imp1","GR_Remesas","MENU",,{"COPREGPP","COPTABPP"})

   W_Imp1.GR_Remesas.SetFocus



   END WINDOW
   VentanaCentrar("W_Imp1","Ventana1")
   ACTIVATE WINDOW W_Imp1

Return




STATIC FUNCTION br_rem1_Terminar(Srem1,Nrem1,Ventana1,ControlSrem,ControlNrem)
   SREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,1)
   NREM2:=W_Imp1.GR_Remesas.Cell(W_Imp1.GR_Remesas.Value,2)
   SetProperty(Ventana1,ControlSrem,"Value",SREM2)
   SetProperty(Ventana1,ControlNrem,"Value",NREM2)
   W_Imp1.Release

