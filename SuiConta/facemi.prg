#include "minigui.ch"

Function FacEmi()
    IF IsWindowDefined(WinFacEmi1)=.T.
        DECLARE WINDOW WinFacEmi1
        WinFacEmi1.restore
        WinFacEmi1.Show
        WinFacEmi1.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinFacEmi1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Facturas emitidas" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("VERDECLARO") ;
        ON RELEASE CloseTables()

        @015,10 LABEL L_NumFac VALUE 'Factura' AUTOSIZE TRANSPARENT
        @010,100 TEXTBOX T_NumFac WIDTH 80 VALUE 0 NUMERIC RIGHTALIGN
        @010,180 TEXTBOX T_SerFac WIDTH 20 VALUE ''
        @010,225 BUTTONEX Bt_BuscarEmi ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION FacEmi_Buscar() ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @010,325 BUTTONEX Bt_CopiarDir CAPTION 'portapapeles' ICON icobus('portapapeles') ;
            ACTION CopiaPP("TEXTO","WinFacEmi1") ;
            WIDTH 100 HEIGHT 25 NOTABSTOP

        @045,010 LABEL L_CodCta VALUE 'Cuenta' AUTOSIZE TRANSPARENT
        @040,100 TEXTBOX T_CodCta WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS (WinFacEmi1.T_CodCta.Value:=PCODCTA3(WinFacEmi1.T_CodCta.Value), FacEmi_ActCta() )
        @040,225 BUTTONEX Bt_BuscarCue ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION (br_cue1(VAL(WinFacEmi1.T_CodCta.Value),"WinFacEmi1","T_CodCta") , FacEmi_ActCta() , WinFacEmi1.D_FecRec.SetFocus ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @040,325 BUTTON Bt_Mayor CAPTION 'Mayor' ;
            ACTION br_suizoext(RUTAPROGRAMA,VAL(WinFacEmi1.T_CodCta.Value),DMA1,DMA2,STR(NUMEMP)) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @075,010 LABEL L_NomCta VALUE 'Descripcion' AUTOSIZE TRANSPARENT
        @070,100 TEXTBOX T_NomCta WIDTH 300 MAXLENGTH 40 NOTABSTOP

        @105,010 LABEL L_FecRec VALUE 'Fecha' AUTOSIZE TRANSPARENT
        @100,100 DATEPICKER D_FecRec WIDTH 100 VALUE DATE()

        @135,010 LABEL L_Vendedor VALUE 'Vendedor' AUTOSIZE TRANSPARENT
        @130,100 COMBOBOX C_Vendedor WIDTH 250 ITEMS {} NOTABSTOP
        @130,360 TEXTBOX T_ComiVend WIDTH 50 VALUE 0 NUMERIC INPUTMASK '999.99' FORMAT 'E'

        @165,010 LABEL L_Asiento VALUE 'Asiento' AUTOSIZE TRANSPARENT
        @160,100 TEXTBOX T_Asiento WIDTH 75 NUMERIC RIGHTALIGN ON CHANGE FacEmi_ActBotAsi() NOTABSTOP
        @160,200 BUTTON Bt_Asiento CAPTION 'Ver asiento' ;
            ACTION FacEmi_Asiento1() ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @195,010 LABEL L_Regimen VALUE 'Regimen' AUTOSIZE TRANSPARENT
        @190,100 COMBOBOX C_Regimen WIDTH 250 ITEMS REGIVAEMI("ARRAY") VALUE 1 NOTABSTOP

        @220,010 LABEL L_Acta VALUE 'A cuenta' AUTOSIZE TRANSPARENT
        @220,100 TEXTBOX T_Acta WIDTH 100 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN

        @225,220 LABEL L_DiaVto VALUE 'Dias Vto.' AUTOSIZE TRANSPARENT
        @220,280 TEXTBOX T_DiaVto1 WIDTH 40 VALUE 0 NUMERIC INPUTMASK '99' ON LOSTFOCUS FacEmi_ActVto1() RIGHTALIGN NOTABSTOP
        @220,320 TEXTBOX T_DiaVto2 WIDTH 40 VALUE 0 NUMERIC INPUTMASK '99' ON LOSTFOCUS FacEmi_ActVto1() RIGHTALIGN NOTABSTOP

        @250,010 BUTTON Bt_FPago CAPTION 'Forma pago' ;
            ACTION ( Fpago(WinFacEmi1.T_FPagoN.Value,"WinFacEmi1","T_FPagoN","T_FPagoC") , FacEmi_ActVto1() ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @250,110 TEXTBOX T_FPagoN WIDTH 50 VALUE 0 NUMERIC INVISIBLE
        @250,110 TEXTBOX T_FPagoC WIDTH 250 NOTABSTOP

        @015,460 LABEL L_Iva1 VALUE 'Base imponible' WIDTH 110 HEIGHT 20 CENTER TRANSPARENT
        @015,580 LABEL L_Iva2 VALUE '% IVA'          WIDTH  50 HEIGHT 20 CENTER TRANSPARENT
        @015,640 LABEL L_Iva3 VALUE 'Cuota IVA'      WIDTH 120 HEIGHT 20 CENTER TRANSPARENT

        @040,460 TEXTBOX T_BImp1   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp1") RIGHTALIGN
        @040,580 TEXTBOX T_Iva1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp1") RIGHTALIGN
        @040,640 TEXTBOX T_ImpIva1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas2("BImp1") RIGHTALIGN

        @070,460 TEXTBOX T_BImp2   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp2") RIGHTALIGN
        @070,580 TEXTBOX T_Iva2    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp2") RIGHTALIGN
        @070,640 TEXTBOX T_ImpIva2 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas2("BImp2") RIGHTALIGN

        @100,460 TEXTBOX T_BImp3   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp3") RIGHTALIGN
        @100,580 TEXTBOX T_Iva3    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("BImp3") RIGHTALIGN
        @100,640 TEXTBOX T_ImpIva3 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas2("BImp3") RIGHTALIGN

        @135,460 LABEL L_Req1 VALUE 'Rec.Equiv.' AUTOSIZE TRANSPARENT
        @130,580 TEXTBOX T_Req1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("Req1") RIGHTALIGN
        @130,640 TEXTBOX T_ImpReq1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas2("Req1") RIGHTALIGN

        @165,460 LABEL L_Ret1 VALUE 'Retencion' AUTOSIZE TRANSPARENT
        @160,580 TEXTBOX T_Ret1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas1("Ret1") RIGHTALIGN
        @160,640 TEXTBOX T_ImpRet1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS FacEmi_Sumas2("Ret1") RIGHTALIGN

        @195,460 LABEL L_TFac VALUE 'Total factura' AUTOSIZE TRANSPARENT FONT 'Arial' SIZE 14
        @190,610 TEXTBOX T_TFac WIDTH 150 HEIGHT 30 VALUE 0 FONT 'Arial' SIZE 14 ;
            NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN

        @225,460 LABEL L_Pend VALUE 'Pendiente de cobro' AUTOSIZE TRANSPARENT
        @225,610 TEXTBOX T_Pend WIDTH 150 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN

        @275,460 LABEL L_Descueto VALUE '' AUTOSIZE FONT "Arial" SIZE 10 BOLD FONTCOLOR MICOLOR("ROJO") TRANSPARENT

        DEFINE TAB Tab1 AT 280,0 WIDTH 800 HEIGHT 250 VALUE 1 ON CHANGE FacEmi_ActAlb() NOTABSTOP  //*** BACKCOLOR MiColor("VERDECLARO")
        PAGE 'Vencimientos/Cobros'

        @030,010 GRID GR_Vtos ;
            HEIGHT 180 ;
            WIDTH 380 ;
            HEADERS {'Registro','Vto.','Importe','Dias','Fec.Vto.','Banco'} ;
            WIDTHS {0,50,100,50,80,75 } ;
            ITEMS {} ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC','99999'},{'TEXTBOX','NUMERIC'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','99999'},{'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT, ;
            BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT}

        @225,180 LABEL L_Vtos VALUE 'Total Vencimientos' AUTOSIZE TRANSPARENT
        @220,290 TEXTBOX T_Vtos WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN

        @030,410 GRID GR_Cobros ;
            HEIGHT 180 ;
            WIDTH 380 ;
            HEADERS {'Registro','Fecha','Importe','Fec.Vto.','Banco','Asiento','Empresa','Descripcion'} ;
            WIDTHS {0,80,90,80,75,55,0,150 } ;
            ITEMS {} ;
            ON DBLCLICK FacEmi_AltCob("GRID") ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
            BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT}

        DEFINE CONTEXT MENU CONTROL GR_Cobros OF WinFacEmi1
        ITEM "Añadir"                  ACTION FacEmi_AltCob("ALTA")
        ITEM "Modificar"               ACTION FacEmi_AltCob("MODIFICAR")
        ITEM "Eliminar"                ACTION FacEmi_AltCob("ELIMINAR")
        SEPARATOR
        ITEM "Contabilizar"            ACTION FacEmi_AltCob("CONTABILIZAR")
        ITEM "Ver asiento"             ACTION FacEmi_AltCob("VERASIENTO")
    END MENU

    @220,410 BUTTON Bt_Alta_Cobro CAPTION 'Alta' ;
        ACTION FacEmi_AltCob("ALTA") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP
    @220,460 BUTTON Bt_Modif_Cobro CAPTION 'Modif.' ;
        ACTION FacEmi_AltCob("MODIFICAR") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP
    @220,510 BUTTON Bt_Elim_Cobro CAPTION 'Elim.' ;
        ACTION FacEmi_AltCob("ELIMINAR") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP

    @225,580 LABEL L_Cobros VALUE 'Total Cobros' AUTOSIZE TRANSPARENT
    @220,690 TEXTBOX T_Cobros WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN

    END PAGE
    PAGE 'Albaranes'

    @030,010 GRID GR_Albaran ;
        HEIGHT 210 ;
        WIDTH 780 ;                                 //1       2         3         4       5         6        7       8
    HEADERS {'Albaran','Codigo','Concepto','Cajas','Unidad','Importe','Total','Coste'} ;
        WIDTHS {60,90,250,60,90,90,90,0 } ;
        ITEMS {} ;
        ON DBLCLICK MsgInfo("No se pueden modificar las lineas del los albaranes") ;
        COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'},{'TEXTBOX','CHARACTER'}, ;
        {'TEXTBOX','NUMERIC','9,999,999,999','E'}, ;
        {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
        {'TEXTBOX','NUMERIC','999,999,999.9999','E'}, ;
        {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
        {'TEXTBOX','NUMERIC','999,999,999.9999','E'} } ;
        JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT,BROWSE_JTFY_LEFT, ;
        BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT}

    END PAGE
    END TAB


    @535,010 BUTTONEX Bt_Nuevo CAPTION 'Nuevo' ICON icobus('nuevo') ;
        ACTION FacEmi_Nuevo("NUEVO") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @535,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') ;
        ACTION FacEmi_Modif(.T.) ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @535,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') ;
        ACTION FacEmi_Guardar() ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @535,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') ;
        ACTION FacEmi_Actualiz("CANCELAR") ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @535,410 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
        ACTION FacEmi_Eliminar() ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    @535,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
        ACTION WinFacEmi1.Release ;
        WIDTH 95 HEIGHT 25 NOTABSTOP

    FacEmi_Modif(.F.)

    END WINDOW
    VentanaCentrar("WinFacEmi1","Ventana1","Alinear")
    CENTER WINDOW WinFacEmi1
    ACTIVATE WINDOW WinFacEmi1

Return Nil

STATIC FUNCTION FacEmi_Buscar()
    aNFAC := br_emibus(RUTAPROGRAMA,NUMEMP,WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value,WinFacEmi1.T_NomCta.Value)
    IF aNFAC[1]<>0
        AbrirDBF("FAC92",,,,,1)
        GO aNFAC[1]
        WinFacEmi1.T_SerFac.Value:=SERFAC
        WinFacEmi1.T_NumFac.Value:=NFAC
        FacEmi_Actualiz("CODREC")
    ENDIF
Return Nil

STATIC FUNCTION FacEmi_Nuevo(LLAMADA)
    AbrirDBF("FAC92",,,,,1)
    GO BOTT
    FEC2:=FFAC

    IF LLAMADA="NUEVO"
        WinFacEmi1.T_NumFac.Value := 0
        WinFacEmi1.T_SerFac.Value := "A"
    ENDIF
    WinFacEmi1.T_CodCta.Value :=""
    WinFacEmi1.T_NomCta.Value :=""
    IF FEC2>=DMA1 .AND. FEC2<=DMA2
        WinFacEmi1.D_FecRec.Value:=FEC2
    ELSE
        WinFacEmi1.D_FecRec.Value:=IF(DATE()<DMA2,DATE(),DMA2)
    ENDIF
    WinFacEmi1.C_Vendedor.Value:=0
    WinFacEmi1.T_ComiVend.Value:=0
    WinFacEmi1.T_Asiento.Value:=0
    WinFacEmi1.C_Regimen.Value:=1
    WinFacEmi1.T_DiaVto1.Value:=0
    WinFacEmi1.T_DiaVto2.Value:=0
    WinFacEmi1.T_FpagoN.Value :=0
    WinFacEmi1.T_FpagoC.Value :=""
    WinFacEmi1.T_Acta.Value   :=0
    WinFacEmi1.T_BImp1.Value  :=0
    WinFacEmi1.T_Iva1.Value   :=0
    WinFacEmi1.T_ImpIva1.Value:=0
    WinFacEmi1.T_BImp2.Value  :=0
    WinFacEmi1.T_Iva2.Value   :=0
    WinFacEmi1.T_ImpIva2.Value:=0
    WinFacEmi1.T_BImp3.Value  :=0
    WinFacEmi1.T_Iva3.Value   :=0
    WinFacEmi1.T_ImpIva3.Value:=0
    WinFacEmi1.T_Req1.Value   :=0
    WinFacEmi1.T_ImpReq1.Value:=0
    WinFacEmi1.T_Ret1.Value   :=0
    WinFacEmi1.T_ImpRet1.Value:=0
    WinFacEmi1.T_TFac.Value   :=0
    WinFacEmi1.T_Pend.Value   :=0
    WinFacEmi1.GR_Vtos.DeleteAllItems
    WinFacEmi1.GR_Vtos.Refresh
    WinFacEmi1.T_Vtos.Value   :=0
    WinFacEmi1.GR_Cobros.DeleteAllItems
    WinFacEmi1.GR_Cobros.Refresh
    WinFacEmi1.T_Cobros.Value  :=0
    IF LLAMADA="NUEVO"
        FacEmi_Modif(.t.)
    ENDIF
    FacEmi_ActBotAsi()
    WinFacEmi1.T_CodCta.SetFocus
Return Nil


STATIC FUNCTION FacEmi_Modif(ModifRec)
    SoloVer:=IF(ModifRec=.T.,.F.,.T.)

    WinFacEmi1.T_NumFac.Enabled :=.T. //F.
    WinFacEmi1.T_SerFac.Enabled :=.T. //F.
    WinFacEmi1.T_CodCta.ReadOnly :=SoloVer
    WinFacEmi1.T_NomCta.ReadOnly :=SoloVer
    WinFacEmi1.D_FecRec.Enabled  :=ModifRec
    WinFacEmi1.C_Vendedor.Enabled:=ModifRec
    WinFacEmi1.T_ComiVend.ReadOnly :=SoloVer
    WinFacEmi1.T_Asiento.ReadOnly:=.T.
    WinFacEmi1.T_FpagoC.ReadOnly :=.T.
    WinFacEmi1.T_DiaVto1.ReadOnly:=SoloVer
    WinFacEmi1.T_DiaVto2.ReadOnly:=SoloVer
    WinFacEmi1.T_Acta.ReadOnly   :=SoloVer
    WinFacEmi1.C_Regimen.Enabled:=ModifRec
    WinFacEmi1.T_BImp1.ReadOnly  :=SoloVer
    WinFacEmi1.T_Iva1.ReadOnly   :=SoloVer
    WinFacEmi1.T_ImpIva1.ReadOnly:=SoloVer
    WinFacEmi1.T_BImp2.ReadOnly  :=SoloVer
    WinFacEmi1.T_Iva2.ReadOnly   :=SoloVer
    WinFacEmi1.T_ImpIva2.ReadOnly:=SoloVer
    WinFacEmi1.T_BImp3.ReadOnly  :=SoloVer
    WinFacEmi1.T_Iva3.ReadOnly   :=SoloVer
    WinFacEmi1.T_ImpIva3.ReadOnly:=SoloVer
    WinFacEmi1.T_Req1.ReadOnly   :=SoloVer
    WinFacEmi1.T_ImpReq1.ReadOnly:=SoloVer
    WinFacEmi1.T_Ret1.ReadOnly   :=SoloVer
    WinFacEmi1.T_ImpRet1.ReadOnly:=SoloVer
    WinFacEmi1.T_TFac.ReadOnly   :=.T.
    WinFacEmi1.T_Pend.ReadOnly   :=.T.
    WinFacEmi1.T_Vtos.ReadOnly   :=.T.
    WinFacEmi1.T_Cobros.ReadOnly  :=.T.

    WinFacEmi1.GR_Vtos.Enabled  :=SoloVer
    WinFacEmi1.GR_Cobros.Enabled:=SoloVer

    WinFacEmi1.Bt_FPago.Enabled    :=ModifRec
    WinFacEmi1.Bt_BuscarCue.Enabled:=ModifRec
    WinFacEmi1.Bt_BuscarEmi.Enabled:=IF(ModifRec=.T.,.F.,.T.)
    WinFacEmi1.Bt_Asiento.Enabled  :=IF(ModifRec=.T.,.F.,.T.)

    WinFacEmi1.Bt_Alta_Cobro.Enabled :=IF(ModifRec=.T.,.F.,.T.)
    WinFacEmi1.Bt_Modif_Cobro.Enabled:=IF(ModifRec=.T.,.F.,.T.)
    WinFacEmi1.Bt_Elim_Cobro.Enabled :=IF(ModifRec=.T.,.F.,.T.)

    IF ModifRec = .T.
        WinFacEmi1.Bt_Nuevo.Enabled    := .F.
        WinFacEmi1.Bt_Modif.Enabled    := .F.
        WinFacEmi1.Bt_Guardar.Enabled  := .T.
        WinFacEmi1.Bt_Cancelar.Enabled := .T.
        WinFacEmi1.Bt_Eliminar.Enabled := .F.
        WinFacEmi1.Bt_Salir.Enabled    := .F.
    ELSE
        WinFacEmi1.Bt_Nuevo.Enabled    := .T.
        WinFacEmi1.Bt_Modif.Enabled    := .T.
        WinFacEmi1.Bt_Guardar.Enabled  := .F.
        WinFacEmi1.Bt_Cancelar.Enabled := .F.
        WinFacEmi1.Bt_Eliminar.Enabled := .T.
        WinFacEmi1.Bt_Salir.Enabled    := .T.
    ENDIF

Return Nil


STATIC FUNCTION FacEmi_Actualiz(LLAMADA)
    AbrirDBF("FAC92",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    IF .NOT. EOF()
        WinFacEmi1.T_CodCta.Value :=PCODCTA2(CODCTA)
        WinFacEmi1.T_NomCta.Value :=CLIENTE
        WinFacEmi1.D_FecRec.Value :=FFAC
        WinFacEmi1.T_Asiento.Value:=ASIENTO
        WinFacEmi1.C_Regimen.Value:=REGIMEN+1
        WinFacEmi1.T_DiaVto1.Value:=VAL(LEFT(DPAGO,2))
        WinFacEmi1.T_DiaVto2.Value:=VAL(RIGHT(DPAGO,2))
        WinFacEmi1.T_FpagoN.Value :=FPAGO
        WinFacEmi1.T_FpagoC.Value :=VENCI_NC(FPAGO,30)
        WinFacEmi1.T_Acta.Value   :=ACTA
        WinFacEmi1.T_BImp1.Value  :=BIMP
        WinFacEmi1.T_Iva1.Value   :=IVA
        WinFacEmi1.T_ImpIva1.Value:=IMPIVA
        WinFacEmi1.T_BImp2.Value  :=BIMPT2
        WinFacEmi1.T_Iva2.Value   :=IVAT2
        WinFacEmi1.T_ImpIva2.Value:=IMPIVAT2
        WinFacEmi1.T_BImp3.Value  :=BIMPT3
        WinFacEmi1.T_Iva3.Value   :=IVAT3
        WinFacEmi1.T_ImpIva3.Value:=IMPIVAT3
        WinFacEmi1.T_Req1.Value   :=REQ
        WinFacEmi1.T_ImpReq1.Value:=IMPREQ
        WinFacEmi1.T_Ret1.Value   :=RET
        WinFacEmi1.T_ImpRet1.Value:=IMPRET
        WinFacEmi1.T_TFac.Value   :=TFAC
        WinFacEmi1.T_Pend.Value   :=PEND
        FacEmi_ActVto1()
        FacEmi_ActCobro1()
    ELSE
        FacEmi_Nuevo("ACTUALIZAR")
    ENDIF
    WinFacEmi1.GR_Albaran.DeleteAllItems
    IF WinFacEmi1.Tab1.Value=2
        FacEmi_ActAlb()
    ENDIF
    WinFacEmi1.L_Descueto.Value:=""

    IF LLAMADA="CANCELAR"
        FacEmi_Modif(.F.)
    ENDIF
    FacEmi_ActBotAsi()

Return Nil

STATIC FUNCTION FacEmi_ActBotAsi()
    IF WinFacEmi1.T_Asiento.Value=0
        WinFacEmi1.Bt_Asiento.Caption:="Contabilizar"
    ELSE
        WinFacEmi1.Bt_Asiento.Caption:="Ver asiento"
    ENDIF
Return Nil

STATIC FUNCTION FacEmi_ActVto1()
    WinFacEmi1.T_Vtos.Value:=0
    WinFacEmi1.GR_Vtos.DeleteAllItems
    aVTO:={}
    IF WinFacEmi1.T_NumFac.Value<>0
        DIAVTO1:=WinFacEmi1.T_DiaVto1.Value
        DIAVTO2:=WinFacEmi1.T_DiaVto2.Value
        IMPVTOS:=WinFacEmi1.T_TFac.Value-WinFacEmi1.T_Acta.Value
        VTOS:=VAL(SUBSTR(LTRIM(STR(WinFacEmi1.T_FpagoN.Value)),2,1))
        IF WinFacEmi1.T_Acta.Value<>0
            AADD(aVTO,{0,1,WinFacEmi1.T_Acta.Value,0,WinFacEmi1.D_FecRec.Value,0})
        ENDIF
        FOR N=1 TO VTOS
            FVTO2:=VENCI(2,WinFacEmi1.T_FpagoN.Value,WinFacEmi1.D_FecRec.Value,,DIAVTO1,DIAVTO2,N)
            AADD(aVTO,{0,2,VENCI(3,WinFacEmi1.T_FpagoN.Value,,IMPVTOS,,,N),FVTO2-WinFacEmi1.D_FecRec.Value,FVTO2,0})
        NEXT
        IF VTOS=0
            AADD(aVTO,{0,3,IMPVTOS,0,WinFacEmi1.D_FecRec.Value,0})
        ENDIF

        FOR N=1 TO LEN(aVTO)
            aVTO[N,2]:=N
            WinFacEmi1.GR_Vtos.AddItem(aVTO[N])
            WinFacEmi1.T_Vtos.Value:=WinFacEmi1.T_Vtos.Value+aVTO[N,3]
        NEXT
        WinFacEmi1.GR_Vtos.Refresh
    ENDIF

Return Nil

STATIC FUNCTION FacEmi_ActCobro1()
    WinFacEmi1.T_Cobros.Value:=0
    WinFacEmi1.GR_Cobros.DeleteAllItems
    AbrirDBF("COBROS",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    DO WHILE SERFAC=WinFacEmi1.T_SerFac.Value .AND. NFAC=WinFacEmi1.T_NumFac.Value .AND. .NOT. EOF()
        aVTO:={}
        AADD(aVTO,RECNO())
        AADD(aVTO,FCOB)
        AADD(aVTO,IMPORTE)
        AADD(aVTO,FVTO)
        AADD(aVTO,BANCO)
        AADD(aVTO,NASI)
        AADD(aVTO,NEMPASI)
        AADD(aVTO,RTRIM(DESCRIP))
        WinFacEmi1.GR_Cobros.AddItem(aVTO)
        WinFacEmi1.T_Cobros.Value:=WinFacEmi1.T_Cobros.Value+IMPORTE
        SKIP
    ENDDO
    WinFacEmi1.GR_Cobros.Refresh
Return Nil

STATIC FUNCTION FacEmi_ActAlb()
    IF WinFacEmi1.Tab1.Value<>2
        RETURN
    ENDIF
    IF WinFacEmi1.GR_Albaran.ItemCount<>0
        RETURN
    ENDIF
    IF WinFacEmi1.T_NumFac.Value=0
        RETURN
    ENDIF
    AbrirDBF("FAC92",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    IF EOF()
        RETURN
    ENDIF
    RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,EMP,0,"SUIALMA")
    RUTA2:=RUTA1[1]
    IF FILE(RUTA2+"\ALBARAN.DBF")=.T. .AND. ;
        FILE(RUTA2+"\ALBARAN.CDX")=.T.

        PonerEspera("Actualizando datos...")
        AbrirDBF("ALBARAN",,,,RUTA2,2)
        SEEK STRZERO(WinFacEmi1.T_NumFac.Value,6)+WinFacEmi1.T_SerFac.Value
        IF DESCUENTO<>0 .AND. .NOT. EOF()
            WinFacEmi1.L_Descueto.Value:="Descuento "+LTRIM(STR(DESCUENTO))+"%"
        ENDIF
        DO WHILE SERFAC=WinFacEmi1.T_SerFac.Value .AND. NFAC=WinFacEmi1.T_NumFac.Value .AND. .NOT. EOF()
            DO EVENTS
            WinFacEmi1.GR_Albaran.AddItem({NALB,RTRIM(CODART),RTRIM(NOMART),CAJA_UNI(NOMART,UNIDAD,1),UNIDAD,PRECIO,ROUND(UNIDAD*PRECIO,2),PCOSTE})
            SKIP
        ENDDO
        QuitarEspera()

    ENDIF

Return Nil

STATIC FUNCTION FacEmi_Guardar()
    IF WinFacEmi1.T_NumFac.Value=0
        MSGSTOP("No ha especificado el numero de factura")
        RETURN
    ENDIF
    IF LEN(RTRIM(WinFacEmi1.T_CodCta.Value))=0
        MSGSTOP("No ha especificado el codigo de cuenta")
        RETURN
    ENDIF
    IF WinFacEmi1.T_TFac.Value=0
        MSGSTOP("El importe de la factura es cero")
        RETURN
    ENDIF

    AbrirDBF("FAC92",,,,,1)
    IF WinFacEmi1.T_NumFac.Value=0
        GO BOTT
        IF .NOT. EOF()
            WinFacEmi1.T_NumFac.Value:=NFAC+1
            WinFacEmi1.T_SerFac.Value:=SERFAC
        ELSE
            WinFacEmi1.T_NumFac.Value:=1
            WinFacEmi1.T_SerFac.Value:="A"
        ENDIF
        APPEND BLANK
    ELSE
        SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    ENDIF

    IF RLOCK()
        REPLACE SERFAC WITH WinFacEmi1.T_SerFac.Value
        REPLACE NFAC WITH WinFacEmi1.T_NumFac.Value
        REPLACE CODCTA WITH VAL(PCODCTA3(WinFacEmi1.T_CodCta.Value))
        REPLACE CLIENTE WITH WinFacEmi1.T_NomCta.Value
        REPLACE FFAC WITH WinFacEmi1.D_FecRec.Value
        REPLACE ASIENTO WITH WinFacEmi1.T_Asiento.Value
        REPLACE REGIMEN WITH WinFacEmi1.C_Regimen.Value-1
        REPLACE FPAGO WITH WinFacEmi1.T_FpagoN.Value
        REPLACE DPAGO WITH STRZERO(WinFacEmi1.T_DiaVto1.Value,2)+STRZERO(WinFacEmi1.T_DiaVto2.Value,2)
        REPLACE ACTA WITH WinFacEmi1.T_Acta.Value
        REPLACE BIMP WITH WinFacEmi1.T_BImp1.Value
        REPLACE IVA WITH WinFacEmi1.T_Iva1.Value
        REPLACE IMPIVA WITH WinFacEmi1.T_ImpIva1.Value
        REPLACE BIMPT2 WITH WinFacEmi1.T_BImp2.Value
        REPLACE IVAT2 WITH WinFacEmi1.T_Iva2.Value
        REPLACE IMPIVAT2 WITH WinFacEmi1.T_ImpIva2.Value
        REPLACE BIMPT3 WITH WinFacEmi1.T_BImp3.Value
        REPLACE IVAT3 WITH WinFacEmi1.T_Iva3.Value
        REPLACE IMPIVAT3 WITH WinFacEmi1.T_ImpIva3.Value
        REPLACE REQ WITH WinFacEmi1.T_Req1.Value
        REPLACE IMPREQ WITH WinFacEmi1.T_ImpReq1.Value
        REPLACE RET WITH WinFacEmi1.T_Ret1.Value
        REPLACE IMPRET WITH WinFacEmi1.T_ImpRet1.Value
        REPLACE TFAC WITH WinFacEmi1.T_TFac.Value
        REPLACE PEND WITH WinFacEmi1.T_Pend.Value
        REPLACE NEMP WITH NUMEMP                    //EMPRESA SUIZO CONTABILIDAD
**      REPLACE EMP WITH NUMEMP //EMPRESA SUIZO COMERCIAL
        DBCOMMIT()
        DBUNLOCK()
    ENDIF
    RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,EMP,0,"SUIALMA")
    RUTA2:=RUTA1[1]
    IF FILE(RUTA2+"\FACTURAS.DBF")=.T. .AND. ;
        FILE(RUTA2+"\FACTURAS.CDX")=.T.
        AbrirDBF("FACTURAS",,,,RUTA2,1)
        SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
        IF .NOT. EOF()
            IF FAC92->FFAC<>FACTURAS->FFAC .OR. ;
                FAC92->ASIENTO<>FACTURAS->ASIENTO .OR. ;
                FAC92->REGIMEN<>FACTURAS->REGIMEN .OR. ;
                FAC92->FPAGO<>FACTURAS->FPAGO .OR. ;
                FAC92->DPAGO<>FACTURAS->DPAGO .OR. ;
                FAC92->ACTA<>FACTURAS->ACTA
                IF MsgYesNo("Desea guardar los datos en comercial","Atencion")=.T.
                    IF RLOCK()
                        REPLACE FFAC WITH WinFacEmi1.D_FecRec.Value
                        REPLACE ASIENTO WITH WinFacEmi1.T_Asiento.Value
                        REPLACE REGIMEN WITH WinFacEmi1.C_Regimen.Value-1
                        REPLACE FPAGO WITH WinFacEmi1.T_FpagoN.Value
                        REPLACE DPAGO WITH STRZERO(WinFacEmi1.T_DiaVto1.Value,2)+STRZERO(WinFacEmi1.T_DiaVto2.Value,2)
                        REPLACE ACTA WITH WinFacEmi1.T_Acta.Value
                        DBCOMMIT()
                        DBUNLOCK()
                    ENDIF
                ENDIF
            ENDIF
        ENDIF
    ENDIF
    MsgInfo('Los datos han sido guardados','Datos guardados')
    FacEmi_Modif(.F.)

Return Nil


STATIC FUNCTION FacEmi_Eliminar()
    IF MSGYESNO("¿Desea eliminar la factura?","Atencion")=.F.
        RETURN
    ENDIF

    IF WinFacEmi1.T_Asiento.Value<>0
        IF MSGYESNO("¡Atencion! la factura esta contabilizada"+HB_OsNewLine()+ ;
            "¿Desea eliminar la factura?","Atencion")=.F.
            RETURN
        ENDIF
    ENDIF

    AbrirDBF("FAC92",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    IF EOF()
        MSGSTOP("No existe la factura")
        RETURN
    ENDIF

    IF RLOCK()
        DELETE
        DBCOMMIT()
        DBUNLOCK()

        AbrirDBF("COBROS",,,,,1)
        SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
        DO WHILE SERFAC=WinFacEmi1.T_SerFac.Value .AND. NFAC=WinFacEmi1.T_NumFac.Value .AND. .NOT. EOF()
            IF RLOCK()
                DELETE
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
            SKIP
        ENDDO

        MsgInfo('La factura ha sido eliminada')

        FacEmi_Actualiz("BORRADO")

    ELSE
        MsgStop('No se ha eliminado el registro'+HB_OsNewLine()+ ;
            'El registro esta siendo utilizado por otro usuario'+HB_OsNewLine()+ ;
            'Por favor, intentelo mas tarde','Error')
    ENDIF

Return Nil


STATIC FUNCTION FacEmi_ActCta()
    AbrirDBF("CUENTAS",,,,,1)
    SEEK VAL(WinFacEmi1.T_CodCta.Value)
    IF .NOT. EOF()
        WinFacEmi1.T_NomCta.Value:=RTRIM(NOMCTA)
    ELSE
        WinFacEmi1.T_NomCta.Value:=""
    ENDIF

Return Nil

STATIC FUNCTION FacEmi_Sumas1(LLAMADA)
    LLAMADA:=UPPER(LLAMADA)
    DO CASE
    CASE LLAMADA="BIMP1"
        WinFacEmi1.T_ImpIva1.Value:=ROUND(WinFacEmi1.T_BImp1.Value*WinFacEmi1.T_Iva1.Value/100,2)
    CASE LLAMADA="BIMP2"
        WinFacEmi1.T_ImpIva2.Value:=ROUND(WinFacEmi1.T_BImp2.Value*WinFacEmi1.T_Iva2.Value/100,2)
    CASE LLAMADA="BIMP3"
        WinFacEmi1.T_ImpIva3.Value:=ROUND(WinFacEmi1.T_BImp3.Value*WinFacEmi1.T_Iva3.Value/100,2)
    CASE LLAMADA="REQ1"
        WinFacEmi1.T_ImpReq1.Value:=ROUND((WinFacEmi1.T_BImp1.Value+WinFacEmi1.T_BImp2.Value+WinFacEmi1.T_BImp3.Value)* ;
            WinFacEmi1.T_Req1.Value/100,2)
    CASE LLAMADA="RET1"
        WinFacEmi1.T_ImpRet1.Value:=ROUND((WinFacEmi1.T_BImp1.Value+WinFacEmi1.T_BImp2.Value+WinFacEmi1.T_BImp3.Value)* ;
            WinFacEmi1.T_Ret1.Value/100,2)
    ENDCASE
    FacEmi_Sumas2()

Return Nil

STATIC FUNCTION FacEmi_Sumas2()
    TFAC2:=WinFacEmi1.T_TFac.Value
    WinFacEmi1.T_TFac.Value:=WinFacEmi1.T_BImp1.Value+WinFacEmi1.T_ImpIva1.Value+ ;
        WinFacEmi1.T_BImp2.Value+WinFacEmi1.T_ImpIva2.Value+ ;
        WinFacEmi1.T_BImp3.Value+WinFacEmi1.T_ImpIva3.Value+ ;
        WinFacEmi1.T_ImpReq1.Value- ;
        WinFacEmi1.T_ImpRet1.Value
    WinFacEmi1.T_Pend.Value:=WinFacEmi1.T_Pend.Value-TFAC2+WinFacEmi1.T_TFac.Value

Return Nil


STATIC FUNCTION FacEmi_Asiento1()

    IF WinFacEmi1.T_Asiento.Value<>0
        br_suizoasi(RUTAEMPRESA,WinFacEmi1.T_Asiento.Value,NUMEMP)
        RETURN
    ENDIF

    IF MSGYESNO("¿Desea contabilizar la factura?")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinEmi2Con ;
        AT 0,0     ;
        WIDTH 600  ;
        HEIGHT 320 ;
        TITLE "Contabilizar" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("VERDECLARO") ;

        @15,10 LABEL L_FecAsi VALUE 'Fecha asiento' AUTOSIZE TRANSPARENT
        @10,110 DATEPICKER  D_FecAsi WIDTH 100 VALUE DATE()

        @040,010 TEXTBOX T_BImp1 WIDTH 100 VALUE WinFacEmi1.T_BImp1.Value+WinFacEmi1.T_BImp2.Value+WinFacEmi1.T_BImp3.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @040,120 BUTTONEX Bt_CodCtaIng CAPTION 'Cuenta ingreso' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaIng.Value),"WinEmi2Con","T_CodCtaIng") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIng.Value ,"WinEmi2Con","L_CodCtaIng") ) NOTABSTOP
        @040,240 TEXTBOX T_CodCtaIng WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaIng.Value:=PCODCTA3(WinEmi2Con.T_CodCtaIng.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIng.Value ,"WinEmi2Con","L_CodCtaIng") )
        @045,350 LABEL L_CodCtaIng VALUE 'Cuenta gasto' AUTOSIZE TRANSPARENT

        @070,010 TEXTBOX T_ImpIva1 WIDTH 100 VALUE WinFacEmi1.T_ImpIva1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @070,120 BUTTONEX Bt_CodCtaIva1 CAPTION 'Cuenta IVA 1' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaIva1.Value),"WinEmi2Con","T_CodCtaIva1") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva1.Value ,"WinEmi2Con","L_CodCtaIva1") ) NOTABSTOP
        @070,240 TEXTBOX T_CodCtaIva1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaIva1.Value:=PCODCTA3(WinEmi2Con.T_CodCtaIva1.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva1.Value ,"WinEmi2Con","L_CodCtaIva1") )
        @075,350 LABEL L_CodCtaIva1 VALUE 'Cuenta IVA 1' AUTOSIZE TRANSPARENT

        @100,010 TEXTBOX T_ImpIva2 WIDTH 100 VALUE WinFacEmi1.T_ImpIva2.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @100,120 BUTTONEX Bt_CodCtaIva2 CAPTION 'Cuenta IVA 2' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaIva2.Value),"WinEmi2Con","T_CodCtaIva2") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva2.Value ,"WinEmi2Con","L_CodCtaIva2") ) NOTABSTOP
        @100,240 TEXTBOX T_CodCtaIva2 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaIva2.Value:=PCODCTA3(WinEmi2Con.T_CodCtaIva2.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva2.Value ,"WinEmi2Con","L_CodCtaIva2") )
        @105,350 LABEL L_CodCtaIva2 VALUE 'Cuenta IVA 2' AUTOSIZE TRANSPARENT

        @130,010 TEXTBOX T_ImpIva3 WIDTH 100 VALUE WinFacEmi1.T_ImpIva3.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @130,120 BUTTONEX Bt_CodCtaIva3 CAPTION 'Cuenta IVA 3' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaIva3.Value),"WinEmi2Con","T_CodCtaIva3") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva3.Value ,"WinEmi2Con","L_CodCtaIva3") ) NOTABSTOP
        @130,240 TEXTBOX T_CodCtaIva3 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaIva3.Value:=PCODCTA3(WinEmi2Con.T_CodCtaIva3.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaIva3.Value ,"WinEmi2Con","L_CodCtaIva3") )
        @135,350 LABEL L_CodCtaIva3 VALUE 'Cuenta IVA 3' AUTOSIZE TRANSPARENT

        @160,010 TEXTBOX T_ImpReq1 WIDTH 100 VALUE WinFacEmi1.T_ImpReq1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @160,120 BUTTONEX Bt_CodCtaReq1 CAPTION 'Cuenta recargo' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaReq1.Value),"WinEmi2Con","T_CodCtaReq1") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaReq1.Value ,"WinEmi2Con","L_CodCtaReq1") ) NOTABSTOP
        @160,240 TEXTBOX T_CodCtaReq1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaReq1.Value:=PCODCTA3(WinEmi2Con.T_CodCtaReq1.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaReq1.Value ,"WinEmi2Con","L_CodCtaReq1") )
        @165,350 LABEL L_CodCtaReq1 VALUE 'Cuenta recargo' AUTOSIZE TRANSPARENT

        @190,010 TEXTBOX T_ImpRet1 WIDTH 100 VALUE WinFacEmi1.T_ImpRet1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @190,120 BUTTONEX Bt_CodCtaRet1 CAPTION 'Cuenta retencion' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaRet1.Value),"WinEmi2Con","T_CodCtaRet1") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaRet1.Value ,"WinEmi2Con","L_CodCtaRet1") ) NOTABSTOP
        @190,240 TEXTBOX T_CodCtaRet1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaRet1.Value:=PCODCTA3(WinEmi2Con.T_CodCtaRet1.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaRet1.Value ,"WinEmi2Con","L_CodCtaRet1") )
        @195,350 LABEL L_CodCtaRet1 VALUE 'Cuenta retencion' AUTOSIZE TRANSPARENT

        @220,010 TEXTBOX T_TFac WIDTH 100 VALUE WinFacEmi1.T_TFac.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @220,120 BUTTONEX Bt_CodCtaPag1 CAPTION 'Cuenta cobro' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinEmi2Con.T_CodCtaPag1.Value),"WinEmi2Con","T_CodCtaPag1") , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaPag1.Value ,"WinEmi2Con","L_CodCtaPag1") ) NOTABSTOP
        @220,240 TEXTBOX T_CodCtaPag1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinEmi2Con.T_CodCtaPag1.Value:=PCODCTA3(WinEmi2Con.T_CodCtaPag1.Value) , ;
            FacEmi_BusCue(WinEmi2Con.T_CodCtaPag1.Value ,"WinEmi2Con","L_CodCtaPag1") )
        @225,350 LABEL L_CodCtaPag1 VALUE 'Cuenta cobro' AUTOSIZE TRANSPARENT

        MiColor_Enabled(.F.,"WinEmi2Con","T_BImp1")
        MiColor_Enabled(.F.,"WinEmi2Con","T_ImpIva1")
        MiColor_Enabled(.F.,"WinEmi2Con","T_ImpIva2")
        MiColor_Enabled(.F.,"WinEmi2Con","T_ImpIva3")
        MiColor_Enabled(.F.,"WinEmi2Con","T_ImpReq1")
        MiColor_Enabled(.F.,"WinEmi2Con","T_ImpRet1")
        MiColor_Enabled(.F.,"WinEmi2Con","T_TFac")
        WinEmi2Con.D_FecAsi.Value:=GetProperty("WinFacEmi1","D_FecRec","Value")
        WinEmi2Con.T_CodCtaIng.Value :="70000000"
        WinEmi2Con.T_CodCtaIva1.Value:=STR(47700000+ROUND(GetProperty("WinFacEmi1","T_Iva1","Value"),0),8)
        WinEmi2Con.T_CodCtaIva2.Value:=STR(47700000+ROUND(GetProperty("WinFacEmi1","T_Iva2","Value"),0),8)
        WinEmi2Con.T_CodCtaIva3.Value:=STR(47700000+ROUND(GetProperty("WinFacEmi1","T_Iva3","Value"),0),8)
        WinEmi2Con.T_CodCtaReq1.Value:=STR(47701000+ROUND(GetProperty("WinFacEmi1","T_Req1","Value"),0),8)
        WinEmi2Con.T_CodCtaRet1.Value:=STR(47510000+ROUND(GetProperty("WinFacEmi1","T_Ret1","Value"),0),8)
        AbrirDBF("COBROS")
        SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
        IF .NOT. EOF()
            WinEmi2Con.T_CodCtaPag1.Value:=STR(BANCO)
        ELSE
            WinEmi2Con.T_CodCtaPag1.Value:="57000000"
        ENDIF
        FacEmi_BusCue(WinEmi2Con.T_CodCtaIng.Value ,"WinEmi2Con","L_CodCtaIng")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaIva1.Value ,"WinEmi2Con","L_CodCtaIva1")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaIva2.Value ,"WinEmi2Con","L_CodCtaIva2")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaIva3.Value ,"WinEmi2Con","L_CodCtaIva3")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaReq1.Value ,"WinEmi2Con","L_CodCtaReq1")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaRet1.Value ,"WinEmi2Con","L_CodCtaRet1")
        FacEmi_BusCue(WinEmi2Con.T_CodCtaPag1.Value ,"WinEmi2Con","L_CodCtaPag1")


        @260,410 BUTTON Bt_Conta ;
            CAPTION 'Contabilizar' ;
            WIDTH 90 HEIGHT 25 ;
            ACTION FacEmi_Asiento2() ;
            NOTABSTOP

        @260,510 BUTTONEX Bt_Salir ;
            CAPTION 'Salir' ;
            ICON icobus('salir') ;
            ACTION WinEmi2Con.Release ;
            WIDTH 80 HEIGHT 25 ;
            NOTABSTOP

    END WINDOW
    VentanaCentrar("WinEmi2Con","Ventana1")

    CENTER WINDOW WinEmi2Con
    ACTIVATE WINDOW WinEmi2Con

Return Nil

STATIC FUNCTION FacEmi_Asiento2()

    PonerEspera("Contabilizando la factura ..." )

    CONDES2:="Fac.emi."+LTRIM(STR(GetProperty("WinFacEmi1","T_NumFac","Value")))+"-"+GetProperty("WinFacEmi1","T_SerFac","Value")+";"+ ;
        PCODCTA4(GetProperty("WinFacEmi1","T_CodCta","Value"))
    CONDES3:="Cobro "+SUBSTR(CONDES2,5,LEN(CONDES2))

    AbrirDBF("CUENTAS",,,,,1)                       //ABRIR PARA SALDOS
    AbrirDBF("Apuntes",,,,,1)
    GO BOTT
    NASI2=IF(NASI<2,2,NASI+1)
    APU2:=1

    aASIENTO:={}
    AADD(aASIENTO,{VAL(PCODCTA3(GetProperty("WinFacEmi1","T_CodCta","Value"))),GetProperty("WinFacEmi1","T_Tfac","Value"),0})
    IF GetProperty("WinFacEmi1","T_ImpIva1","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIva1.Value),0,GetProperty("WinFacEmi1","T_ImpIva1","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_BImp2","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIng.Value),0,GetProperty("WinFacEmi1","T_BImp2","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_ImpIva2","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIva2.Value),0,GetProperty("WinFacEmi1","T_ImpIva2","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_BImp3","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIng.Value),0,GetProperty("WinFacEmi1","T_BImp3","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_ImpIva3","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIva3.Value),0,GetProperty("WinFacEmi1","T_ImpIva3","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_ImpReq1","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaReq1.Value),0,GetProperty("WinFacEmi1","T_ImpReq1","Value")})
    ENDIF
    IF GetProperty("WinFacEmi1","T_ImpRet1","Value")<>0
        AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaRet1.Value),GetProperty("WinFacEmi1","T_ImpRet1","Value"),0})
    ENDIF
    AADD(aASIENTO,{VAL(WinEmi2Con.T_CodCtaIng.Value),0,GetProperty("WinFacEmi1","T_BImp1","Value")})

    FOR N=1 TO LEN(aASIENTO)
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE APU WITH N
            REPLACE NEMP WITH NUMEMP
            REPLACE FECHA WITH WinEmi2Con.D_FecAsi.Value
            REPLACE CODCTA WITH aASIENTO[N,1]
            REPLACE NOMAPU WITH CONDES2
            REPLACE DEBE  WITH aASIENTO[N,2]
            REPLACE HABER WITH aASIENTO[N,3]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
        ENDIF
    NEXT

    AbrirDBF("FAC92",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    IF .NOT. EOF()
        IF RLOCK()
            REPLACE ASIENTO WITH NASI2
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
    ENDIF
    SetProperty("WinFacEmi1","T_Asiento","Value",NASI2)
    FacEmi_ActBotAsi()

    IF GetProperty("WinFacEmi1","T_Pend","Value")<>GetProperty("WinFacEmi1","T_Tfac","Value")
        AbrirDBF("COBROS",,,,,1)
        SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
        DO WHILE SERFAC=WinFacEmi1.T_SerFac.Value .AND. NFAC=WinFacEmi1.T_NumFac.Value .AND. .NOT. EOF()
            IF NASI=0
                NASI2++
                APU2:=1
                FASI2:=MAX(FCOB,FVTO)
                AbrirDBF("Apuntes",,,,,1)
                APPEND BLANK
                IF RLOCK()
                    REPLACE NASI WITH NASI2
                    REPLACE APU WITH APU2++
                    REPLACE NEMP WITH NUMEMP
                    REPLACE FECHA WITH FASI2
                    REPLACE CODCTA WITH VAL(PCODCTA3(GetProperty("WinFacEmi1","T_CodCta","Value")))
                    REPLACE NOMAPU WITH CONDES3
                    REPLACE DEBE WITH COBROS->IMPORTE
                    DBCOMMIT()
                    DBUNLOCK()
                    Suizo_saldocuenta(CODCTA,DEBE,HABER)
                ENDIF
                APPEND BLANK
                IF RLOCK()
                    REPLACE NASI WITH NASI2
                    REPLACE APU WITH APU2++
                    REPLACE NEMP WITH NUMEMP
                    REPLACE FECHA WITH FASI2
                    REPLACE CODCTA WITH VAL(WinEmi2Con.T_CodCtaPag1.Value)
                    REPLACE NOMAPU WITH CONDES3
                    REPLACE HABER WITH COBROS->IMPORTE
                    DBCOMMIT()
                    DBUNLOCK()
                    Suizo_saldocuenta(CODCTA,DEBE,HABER)
                ENDIF
                AbrirDBF("COBROS",,,,,1)
                IF RLOCK()
                    REPLACE NASI WITH NASI2
                    REPLACE FASI WITH FASI2
                    REPLACE NEMPASI WITH NUMEMP
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
            ENDIF
            AbrirDBF("COBROS",,,,,1)
            SKIP
        ENDDO
    ENDIF

    QuitarEspera()

    FacEmi_ActCobro1()
    MsgInfo('Los datos han sido guardados','Datos guardados')
    WinEmi2Con.Release

Return Nil


STATIC FUNCTION FacEmi_AltCob(LLAMADA)
    IF WinFacEmi1.T_NumFac.Value=0
        MSGINFO("No se ha seleccionado ninguna factura")
        RETURN
    ENDIF

    IF LLAMADA="GRID"
        IF WinFacEmi1.GR_Cobros.Value=0
            LLAMADA:="ALTA"
        ELSE
            LLAMADA:="MODIFICAR"
        ENDIF
    ENDIF

    IF LLAMADA<>"ALTA"
        IF WinFacEmi1.GR_Cobros.Value=0
            MSGINFO("No se ha seleccionado ningun cobro")
            RETURN
        ENDIF
    ENDIF

    DO CASE
    CASE LLAMADA="VERASIENTO"
        aCOBROS:=WinFacEmi1.GR_Cobros.Item(WinFacEmi1.GR_Cobros.Value)
        IF aCOBROS[6]<>0
            br_suizoasi(RUTAEMPRESA,aCOBROS[6],aCOBROS[7])
        ENDIF
        RETURN
    CASE LLAMADA="ELIMINAR"
        IF MSGYESNO("¿Desea eliminar el cobro?","Atencion")=.F.
            RETURN
        ENDIF
        aCOBROS:=WinFacEmi1.GR_Cobros.Item(WinFacEmi1.GR_Cobros.Value)
        AbrirDBF("COBROS")
        GO aCOBROS[1]
        IF NFAC<>WinFacEmi1.T_NumFac.Value
            MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
            MsgInfo('El cobro ha sido eliminado')
        ENDIF
        WinFacEmi1.T_Pend.Value:=WinFacEmi1.T_Pend.Value+aCOBROS[3]

    CASE LLAMADA="ALTA" .OR. LLAMADA="MODIFICAR"
        aBANCOS1:={}
        aBANCOS2:={}
        AbrirDBF("CUENTAS",,,,,1)
        SEEK 57000000
        DO WHILE CODCTA<57499999 .AND. .NOT. EOF()
            AADD(aBANCOS1,RTRIM(NOMCTA))
            AADD(aBANCOS2,CODCTA)
            SKIP
        ENDDO
        aLabels    :={'Fecha','Importe','Descripcion','Fec.Vto.','Banco'}
        aFormats   :={Nil,'9,999,999,999.99',30,Nil,aBANCOS1}
        IF LLAMADA="ALTA"
            DO CASE
            CASE WinFacEmi1.GR_Vtos.ItemCount>WinFacEmi1.GR_Cobros.ItemCount
                FCOBRO2:=WinFacEmi1.GR_Vtos.Cell(WinFacEmi1.GR_Cobros.ItemCount+1,5)
            CASE WinFacEmi1.GR_Cobros.ItemCount>=1
                FCOBRO2:=WinFacEmi1.GR_Cobros.Cell(WinFacEmi1.GR_Cobros.ItemCount,2)
            OTHERWISE
                FCOBRO2:=WinFacEmi1.D_FecRec.Value
            ENDCASE
            TIT2:="Alta de cobros"
            aInitValues:={FCOBRO2,WinFacEmi1.T_Pend.Value,'Cobrado',WinFacEmi1.D_FecRec.Value,1}
            aBOTON:={"Nuevo","Cancelar"}
        ELSE
            TIT2:="Modificar cobros"
            aCOBROS:=WinFacEmi1.GR_Cobros.Item(WinFacEmi1.GR_Cobros.Value)
            CODBAN2:=ASCAN(aBANCOS2,{|AVAL| AVAL=aCOBROS[5]})
            aInitValues:={aCOBROS[2],aCOBROS[3],aCOBROS[8],aCOBROS[4],CODBAN2}
            aBOTON:={"Modificar","Cancelar"}
        ENDIF

        AltaCob2:=InputWindow(TIT2, aLabels , aInitValues , aFormats,WinFacEmi1.Row+260,WinFacEmi1.Col+410 , .T. , aBOTON)

        IF AltaCob2[1] == Nil
            RETURN
        ENDIF

        IF AltaCob2[2]=0
            MSGSTOP("No se ha especificado el importe"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF

        AbrirDBF("COBROS")
        IF LLAMADA="ALTA"
            APPEND BLANK
        ELSE
            GO aCOBROS[1]
            IF NFAC<>WinFacEmi1.T_NumFac.Value
                MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
                RETURN
            ENDIF
        ENDIF

        AltaCob2[4]:=IF(AltaCob2[1]>AltaCob2[4],AltaCob2[1],AltaCob2[4])
        CODBAN2:=IF(AltaCob2[5]=0,0, aBANCOS2[AltaCob2[5]] )
        IF RLOCK()
            REPLACE SERFAC WITH WinFacEmi1.T_SerFac.Value
            REPLACE NFAC WITH WinFacEmi1.T_NumFac.Value
            REPLACE FFAC WITH WinFacEmi1.D_FecRec.Value
            REPLACE FCOB WITH AltaCob2[1]
            REPLACE IMPORTE WITH AltaCob2[2]
            REPLACE DESCRIP WITH AltaCob2[3]
            REPLACE FVTO WITH AltaCob2[4]
            REPLACE BANCO WITH CODBAN2
            REPLACE NEMP WITH NUMEMP
            DBCOMMIT()
            DBUNLOCK()
        ENDIF

   ***Comprobar fecha de asiento
        IF LLAMADA="MODIFICAR"
            IF aCOBROS[6]<>0
                IF aCOBROS[7]<>NUMEMP
                    RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,aCOBROS[7],"SUICONTA")
                    RUTA2:=RUTA1[1]
                ELSE
                    RUTA2:=RUTAEMPRESA
                ENDIF
                AbrirDBF("Apuntes",,,,RUTA2,1)
                SEEK aCOBROS[6]
                IF FECHA<>MAX(AltaCob2[1],AltaCob2[4])
                    IF MsgYesNo("La fecha ha sido modificada"+HB_OsNewLine()+ ;
                        "Fecha cobro: "+DIA(MAX(AltaCob2[1],AltaCob2[4]),10)+HB_OsNewLine()+ ;
                        "Fecha asiento: "+DIA(FECHA,10)+HB_OsNewLine()+ ;
                        "¿Desea modificar la fecha del asiento?")=.T.
                        DO WHILE NASI=aCOBROS[6]
                            IF RLOCK()
                                REPLACE FECHA WITH MAX(AltaCob2[1],AltaCob2[4])
                                DBCOMMIT()
                                DBUNLOCK()
                            ENDIF
                            SKIP
                        ENDDO
                    ENDIF
                ENDIF
            ENDIF
        ENDIF

        IF LLAMADA="ALTA"
            WinFacEmi1.T_Pend.Value:=WinFacEmi1.T_Pend.Value-AltaCob2[2]
        ELSE
            WinFacEmi1.T_Pend.Value:=WinFacEmi1.T_Pend.Value-AltaCob2[2]+aCOBROS[3]
        ENDIF

    CASE LLAMADA="CONTABILIZAR"
        aCOBROS:=WinFacEmi1.GR_Cobros.Item(WinFacEmi1.GR_Cobros.Value)

        IF aCOBROS[6]<>0
            MSGSTOP("El cobro ya esta contabilizado","error")
            RETURN
        ENDIF

        IF MSGYESNO("¿Desea contabilizar el cobro?")=.F.
            RETURN
        ENDIF

        FASI2:=MAX(aCOBROS[2],aCOBROS[4])

        IF YEAR(FASI2)<>EJERANY
            IF MSGYESNO("¿Desea contabilizar el cobro en el ejercicio "+LTRIM(STR(YEAR(FASI2)))+"?" )=.F.
                RETURN
            ENDIF
            RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,YEAR(FASI2),"SUICONTA")
            RUTA2:=RUTA1[1]
            NUMEMP2:=VAL(RUTA1[2])
        ELSE
            RUTA2:=RUTAEMPRESA
            NUMEMP2:=NUMEMP
        ENDIF

        IF .NOT. FILE(RUTA2+"\Apuntes.dbf")
            MSGSTOP("No se ha localizado el fichero de apuntes"+HB_OsNewLine()+ ;
                RUTA2+"\Apuntes.dbf","error")
            RETURN
        ENDIF

        AbrirDBF("CUENTAS",,,,RUTA2,1)              //ABRIR PARA SALDOS
        AbrirDBF("Apuntes",,,,RUTA2,1)
        GO BOTT
        NASI2=IF(NASI<2,2,NASI+1)
        APU2:=1
        IF UPPER(aCOBROS[8])="COBRADO"
            NOMCOB2:="Cobro fac."
        ELSE
		***PARA QUE EL MINIMO NO SEA CERO AÑADIR " " Y "."
            NUMAT2:=MIN(AT(" ",UPPER(aCOBROS[8])+" "),AT(".",UPPER(aCOBROS[8])+"."))
            IF NUMAT2>1
                NOMCOB2:=LEFT(aCOBROS[8],NUMAT2-1)+" fac."
            ELSE
                NOMCOB2:="Cobro fac."
            ENDIF
        ENDIF

        NOMCOB2:=NOMCOB2+LTRIM(STR(GetProperty("WinFacEmi1","T_NumFac","Value")))+"-"+GetProperty("WinFacEmi1","T_SerFac","Value")
        SELEC APUNTES
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE APU WITH APU2++
            REPLACE NEMP WITH NUMEMP2
            REPLACE FECHA WITH FASI2
            REPLACE CODCTA WITH aCOBROS[5]
            REPLACE NOMAPU WITH NOMCOB2+";"+PCODCTA4(GetProperty("WinFacEmi1","T_CodCta","Value"))
            REPLACE DEBE WITH aCOBROS[3]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
        ENDIF
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE APU WITH APU2++
            REPLACE NEMP WITH NUMEMP2
            REPLACE FECHA WITH FASI2
            REPLACE CODCTA WITH VAL(PCODCTA3(GetProperty("WinFacEmi1","T_CodCta","Value")))
            REPLACE NOMAPU WITH NOMCOB2+";"+PCODCTA4(aCOBROS[5])
            REPLACE HABER WITH aCOBROS[3]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
        ENDIF
        AbrirDBF("COBROS",,,,,1)
        GO aCOBROS[1]
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE FASI WITH FASI2
            REPLACE NEMPASI WITH NUMEMP2
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
        MSGINFO("El cobro se ha contabilizado correctamente"+HB_OsNewLine()+ ;
            "Empresa: "+LTRIM(STR(NUMEMP2))+HB_OsNewLine()+ ;
            "Asiento: "+LTRIM(STR(NASI2))+HB_OsNewLine()+ ;
            "Fecha: "+DIA(FASI2,10))

    ENDCASE

    AbrirDBF("FAC92",,,,,1)
    SEEK SERIE(WinFacEmi1.T_SerFac.Value,WinFacEmi1.T_NumFac.Value)
    IF .NOT. EOF()
        IF RLOCK()
            REPLACE PEND WITH WinFacEmi1.T_Pend.Value
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
    ENDIF


    FacEmi_ActCobro1()

Return Nil

STATIC FUNCTION FacEmi_BusCue(CODCTA,VENTANA,CONTROL)
    AbrirDBF("CUENTAS",,,,,1)
    SEEK VAL(CODCTA)
    IF .NOT. EOF()
        SetProperty(VENTANA,CONTROL,"Value",RTRIM(NOMCTA))
    ELSE
        SetProperty(VENTANA,CONTROL,"Value","")
    ENDIF
Return Nil

***REGIMEN IVA***
FUNCTION REGIVAEMI(REGIVA2)
    IF VALTYPE(REGIVA2)="N"
        DO CASE
        CASE REGIVA2=0
            REGIVA3:="Devengado en Regimen General       "
        CASE REGIVA2=1
            REGIVA3:="Devengado Operaciones Intracomunit."
        CASE REGIVA2=2
            REGIVA3:="Entregas Intracomunitarias Exentas "
        CASE REGIVA2=3
            REGIVA3:="Devengado en Exportaciones         "
        CASE REGIVA2=4
            REGIVA3:="Devengado inversion sujeto pasivo  "
        OTHERWISE
            REGIVA3:="Error                              "
        ENDCASE
    ELSE
        REGIVA3:={"Devengado en Regimen General       ",;
            "Devengado Operaciones Intracomunit.",;
            "Entregas Intracomunitarias Exentas ",;
            "Devengado en Exportaciones         ",;
            "Devengado inversion sujeto pasivo  "}
    ENDIF
RETURN(REGIVA3)
