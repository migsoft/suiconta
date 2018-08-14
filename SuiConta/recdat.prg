#include "minigui.ch"

Function RecDat()
    IF IsWindowDefined(WinRecDat1)=.T.
        DECLARE WINDOW WinRecDat1
        WinRecDat1.restore
        WinRecDat1.Show
        WinRecDat1.SetFocus
        RETURN
    ENDIF

    IF Ejer_Cerrado(EJERCERRADO,"VER")=.F.
        RETURN
    ENDIF
    DEFINE WINDOW WinRecDat1 ;
        AT 50,0     ;
        WIDTH 800  ;
        HEIGHT 600 ;
        TITLE "Facturas recibidas" ;
        CHILD NOMAXIMIZE ;
        NOSIZE BACKCOLOR MiColor("ARENA") ;
        ON RELEASE CloseTables()

        @015,010 LABEL L_CodRec VALUE 'Registro' AUTOSIZE TRANSPARENT
        @010,100 TEXTBOX T_CodRec WIDTH 100 VALUE 0 NUMERIC ;
            ON LOSTFOCUS RecDat_Actualizar("CODREC")
        @010,225 BUTTONEX Bt_BuscarRec ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION (br_rec1(WinRecDat1.T_CodRec.Value,"WinRecDat1","T_CodRec") , RecDat_Actualizar("CODREC") ) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @010,325 BUTTONEX Bt_CopiarDir CAPTION 'portapapeles' ICON icobus('portapapeles') ;
            ACTION CopiaPP("TEXTO","WinRecDat1") ;
            WIDTH 100 HEIGHT 25 NOTABSTOP

        @045,010 LABEL L_CodCta VALUE 'Cuenta' AUTOSIZE TRANSPARENT
        @040,100 TEXTBOX T_CodCta WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS (WinRecDat1.T_CodCta.Value:=PCODCTA3(WinRecDat1.T_CodCta.Value), RecDat_ActCta() )
        @040,225 BUTTONEX Bt_BuscarCue ;
            CAPTION 'Buscar' ICON icobus('buscar') ;
            ACTION RecDat_BotCue() ;
            WIDTH 90 HEIGHT 25 NOTABSTOP
        @040,325 BUTTON Bt_Mayor CAPTION 'Mayor' ;
            ACTION br_suizoext(RUTAPROGRAMA,VAL(WinRecDat1.T_CodCta.Value),DMA1,DMA2,STR(NUMEMP)) ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @075,010 LABEL L_NomCta VALUE 'Descripcion' AUTOSIZE TRANSPARENT
        @070,100 TEXTBOX T_NomCta WIDTH 300 MAXLENGTH 40 NOTABSTOP

        @105,010 LABEL L_CifRec VALUE 'Cif/Nif' AUTOSIZE TRANSPARENT
        @100,100 TEXTBOX T_CifRec WIDTH 125 MAXLENGTH 15 NOTABSTOP ;
            ON LOSTFOCUS WinRecDat1.T_CifRec.Value:=DNI_LET(WinRecDat1.T_CifRec.Value)

        @135,010 LABEL L_FecRec VALUE 'Fecha' AUTOSIZE TRANSPARENT
        @130,100 DATEPICKER D_FecRec WIDTH 100 VALUE DATE()

        @165,010 LABEL L_NumFac VALUE 'Factura' AUTOSIZE TRANSPARENT
        @160,100 TEXTBOX T_NumFac WIDTH 150 MAXLENGTH 15

        @195,010 LABEL L_Asiento VALUE 'Asiento' AUTOSIZE TRANSPARENT
        @190,100 TEXTBOX T_Asiento WIDTH 75 NUMERIC RIGHTALIGN ON CHANGE RecDat_ActBotAsi() NOTABSTOP
        @190,200 BUTTON Bt_Asiento CAPTION 'Ver asiento' ;
            ACTION RecDat_Asiento1() ;
            WIDTH 90 HEIGHT 25 NOTABSTOP

        @225,010 LABEL L_Regimen VALUE 'Regimen' AUTOSIZE TRANSPARENT
        @220,100 COMBOBOX C_Regimen WIDTH 250 ITEMS REGIVAREC("ARRAY") VALUE 1 NOTABSTOP

        @015,460 LABEL L_Iva1 VALUE 'Base imponible' WIDTH 110 HEIGHT 20 CENTER TRANSPARENT
        @015,580 LABEL L_Iva2 VALUE '% IVA'          WIDTH  50 HEIGHT 20 CENTER TRANSPARENT
        @015,640 LABEL L_Iva3 VALUE 'Cuota IVA'      WIDTH 120 HEIGHT 20 CENTER TRANSPARENT

        @040,460 TEXTBOX T_BImp1   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp1") RIGHTALIGN
        @040,580 TEXTBOX T_Iva1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp1") RIGHTALIGN
        @040,640 TEXTBOX T_ImpIva1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas2("BImp1") RIGHTALIGN

        @070,460 TEXTBOX T_BImp2   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp2") RIGHTALIGN
        @070,580 TEXTBOX T_Iva2    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp2") RIGHTALIGN
        @070,640 TEXTBOX T_ImpIva2 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas2("BImp2") RIGHTALIGN

        @100,460 TEXTBOX T_BImp3   WIDTH 110 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp3") RIGHTALIGN
        @100,580 TEXTBOX T_Iva3    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("BImp3") RIGHTALIGN
        @100,640 TEXTBOX T_ImpIva3 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas2("BImp3") RIGHTALIGN

        @135,460 LABEL L_Req1 VALUE 'Rec.Equiv.' AUTOSIZE TRANSPARENT
        @130,580 TEXTBOX T_Req1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("Req1") RIGHTALIGN
        @130,640 TEXTBOX T_ImpReq1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas2("Req1") RIGHTALIGN

        @165,460 LABEL L_Ret1 VALUE 'Retencion' AUTOSIZE TRANSPARENT
        @160,580 TEXTBOX T_Ret1    WIDTH  50 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '999.99'           FORMAT 'E' ON LOSTFOCUS RecDat_Sumas1("Ret1") RIGHTALIGN
        @160,640 TEXTBOX T_ImpRet1 WIDTH 120 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' ON LOSTFOCUS RecDat_Sumas2("Ret1") RIGHTALIGN

        @195,460 LABEL L_TFac VALUE 'Total factura' AUTOSIZE TRANSPARENT ;
            FONT 'Arial' SIZE 14
        @190,610 TEXTBOX T_TFac ;
            WIDTH 150 HEIGHT 30 ;
            VALUE 0 ;
            FONT 'Arial' SIZE 14 ;
            NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' ;
            RIGHTALIGN

        @225,460 LABEL L_Pend VALUE 'Pendiente de cobro' AUTOSIZE TRANSPARENT
        @225,610 TEXTBOX T_Pend WIDTH 150 HEIGHT 20 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' RIGHTALIGN


        @260,010 GRID GR_Vtos ;
            HEIGHT 200 ;
            WIDTH 380 ;
            HEADERS {'Registro','Vto.','Importe','Dias','Fec.Vto.','Banco'} ;
            WIDTHS {0,50,100,50,80,75 } ;
            ITEMS {} ;
            ON DBLCLICK RecDat_AltaVto("GRID") ;
            COLUMNCONTROLS {{'TEXTBOX','NUMERIC','99999'},{'TEXTBOX','NUMERIC'}, ;
            {'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
            {'TEXTBOX','NUMERIC','99999'},{'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'}} ;
            JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT, ;
            BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT}

        DEFINE CONTEXT MENU CONTROL GR_Vtos OF WinRecDat1
        ITEM "Añadir"                  ACTION RecDat_AltaVto("ALTA")
        ITEM "Modificar"               ACTION RecDat_AltaVto("MODIFICAR")
        ITEM "Eliminar"                ACTION RecDat_AltaVto("ELIMINAR")
    END MENU

    @470,010 BUTTON Bt_Alta_Vtos CAPTION 'Alta' ;
        ACTION RecDat_AltaVto("ALTA") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP
    @470,060 BUTTON Bt_Modif_Vtos CAPTION 'Modif.' ;
        ACTION RecDat_AltaVto("MODIFICAR") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP
    @470,110 BUTTON Bt_Elim_Vtos CAPTION 'Elim.' ;
        ACTION RecDat_AltaVto("ELIMINAR") ;
        WIDTH 40 HEIGHT 25 NOTABSTOP

    @475,180 LABEL L_Vtos VALUE 'Total Vencimientos' AUTOSIZE TRANSPARENT
    @470,290 TEXTBOX T_Vtos WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN

    @260,410 GRID GR_Pagos ;
        HEIGHT 200 ;
        WIDTH 380 ;
        HEADERS {'Registro','Fecha','Importe','Fec.Vto.','Banco','Asiento','Empresa','Descripcion'} ;
        WIDTHS {0,80,90,80,75,55,0,150 } ;
        ITEMS {} ;
        ON DBLCLICK RecDat_AltaPago("GRID") ;
        COLUMNCONTROLS {{'TEXTBOX','NUMERIC'},{'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC','9,999,999,999.99','E'}, ;
        {'TEXTBOX','DATE'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','NUMERIC'},{'TEXTBOX','CHARACTER'}} ;
        JUSTIFY {BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER,BROWSE_JTFY_RIGHT,BROWSE_JTFY_CENTER, ;
        BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_RIGHT,BROWSE_JTFY_LEFT}

    DEFINE CONTEXT MENU CONTROL GR_Pagos OF WinRecDat1
    ITEM "Añadir"                      ACTION RecDat_AltaPago("ALTA")
    ITEM "Modificar"                   ACTION RecDat_AltaPago("MODIFICAR")
    ITEM "Eliminar"                    ACTION RecDat_AltaPago("ELIMINAR")
    SEPARATOR
    ITEM "Contabilizar"                ACTION RecDat_AltaPago("CONTABILIZAR")
    ITEM "Ver asiento"                 ACTION RecDat_AltaPago("VERASIENTO")
    END MENU

    @470,410 BUTTON Bt_Alta_Pago CAPTION 'Alta' ;
    ACTION RecDat_AltaPago("ALTA") ;
    WIDTH 40 HEIGHT 25 NOTABSTOP
    @470,460 BUTTON Bt_Modif_Pago CAPTION 'Modif.' ;
    ACTION RecDat_AltaPago("MODIFICAR") ;
    WIDTH 40 HEIGHT 25 NOTABSTOP
    @470,510 BUTTON Bt_Elim_Pago CAPTION 'Elim.' ;
    ACTION RecDat_AltaPago("ELIMINAR") ;
    WIDTH 40 HEIGHT 25 NOTABSTOP

    @475,580 LABEL L_Pagos VALUE 'Total Pagos' AUTOSIZE TRANSPARENT
    @470,690 TEXTBOX T_Pagos WIDTH 100 VALUE 0 NUMERIC INPUTMASK '9,999,999,999.99' FORMAT 'E' RIGHTALIGN



    @525,010 BUTTONEX Bt_Nuevo CAPTION 'Nuevo' ICON icobus('nuevo') ;
    ACTION RecDat_Nuevo("NUEVO") ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    @525,110 BUTTONEX Bt_Modif CAPTION 'Modificar' ICON icobus('modificar') ;
    ACTION RecDat_Modif(.T.) ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    @525,210 BUTTONEX Bt_Guardar CAPTION 'Guardar' ICON icobus('guardar') ;
    ACTION RecDat_Guardar() ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    @525,310 BUTTONEX Bt_Cancelar CAPTION 'Cancelar' ICON icobus('cancelar') ;
    ACTION RecDat_Actualizar("CANCELAR") ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    @525,410 BUTTONEX Bt_Eliminar CAPTION 'Eliminar' ICON icobus('eliminar') ;
    ACTION RecDat_Eliminar() ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    @525,510 BUTTONEX Bt_Salir CAPTION 'Salir' ICON icobus('salir') ;
    ACTION WinRecDat1.Release ;
    WIDTH 95 HEIGHT 25 NOTABSTOP

    RecDat_Modif(.F.)

    END WINDOW

    VentanaCentrar("WinRecDat1","Ventana1","Alinear")
    CENTER WINDOW WinRecDat1
    ACTIVATE WINDOW WinRecDat1

Return Nil


STATIC FUNCTION RecDat_Nuevo(LLAMADA)
    AbrirDBF("FACREB",,,,,1)
    GO BOTT
    FEC2:=FREG

    IF LLAMADA="NUEVO"
        WinRecDat1.T_CodRec.Value := 0
    ENDIF
    WinRecDat1.T_CodCta.Value :=""
    WinRecDat1.T_NomCta.Value :=""
    WinRecDat1.T_CifRec.Value :=""
    IF FEC2>=DMA1 .AND. FEC2<=DMA2
        WinRecDat1.D_FecRec.Value:=FEC2
    ELSE
        WinRecDat1.D_FecRec.Value:=IF(DATE()<DMA2,DATE(),DMA2)
    ENDIF
    WinRecDat1.T_NumFac.Value :=""
    WinRecDat1.T_Asiento.Value:=0
    WinRecDat1.C_Regimen.Value:=1
    WinRecDat1.T_BImp1.Value  :=0
    WinRecDat1.T_Iva1.Value   :=0
    WinRecDat1.T_ImpIva1.Value:=0
    WinRecDat1.T_BImp2.Value  :=0
    WinRecDat1.T_Iva2.Value   :=0
    WinRecDat1.T_ImpIva2.Value:=0
    WinRecDat1.T_BImp3.Value  :=0
    WinRecDat1.T_Iva3.Value   :=0
    WinRecDat1.T_ImpIva3.Value:=0
    WinRecDat1.T_Req1.Value   :=0
    WinRecDat1.T_ImpReq1.Value:=0
    WinRecDat1.T_Ret1.Value   :=0
    WinRecDat1.T_ImpRet1.Value:=0
    WinRecDat1.T_TFac.Value   :=0
    WinRecDat1.T_Pend.Value   :=0
    WinRecDat1.GR_Vtos.DeleteAllItems
    WinRecDat1.GR_Vtos.Refresh
    WinRecDat1.T_Vtos.Value   :=0
    WinRecDat1.GR_Pagos.DeleteAllItems
    WinRecDat1.GR_Pagos.Refresh
    WinRecDat1.T_Pagos.Value  :=0
    IF LLAMADA="NUEVO"
        RecDat_Modif(.t.)
    ENDIF
    RecDat_ActBotAsi()
    WinRecDat1.T_CodCta.SetFocus
Return Nil


STATIC FUNCTION RecDat_Modif(ModifRec)
    SoloVer:=IF(ModifRec=.T.,.F.,.T.)

    WinRecDat1.T_CodRec.Enabled :=.F.
    WinRecDat1.T_CodCta.ReadOnly :=SoloVer
    WinRecDat1.T_NomCta.ReadOnly :=SoloVer
    WinRecDat1.T_CifRec.ReadOnly :=SoloVer
    WinRecDat1.D_FecRec.Enabled :=ModifRec
    WinRecDat1.T_NumFac.ReadOnly :=SoloVer
    WinRecDat1.T_Asiento.ReadOnly:=.T.
    WinRecDat1.C_Regimen.Enabled:=ModifRec
    WinRecDat1.T_BImp1.ReadOnly  :=SoloVer
    WinRecDat1.T_Iva1.ReadOnly   :=SoloVer
    WinRecDat1.T_ImpIva1.ReadOnly:=SoloVer
    WinRecDat1.T_BImp2.ReadOnly  :=SoloVer
    WinRecDat1.T_Iva2.ReadOnly   :=SoloVer
    WinRecDat1.T_ImpIva2.ReadOnly:=SoloVer
    WinRecDat1.T_BImp3.ReadOnly  :=SoloVer
    WinRecDat1.T_Iva3.ReadOnly   :=SoloVer
    WinRecDat1.T_ImpIva3.ReadOnly:=SoloVer
    WinRecDat1.T_Req1.ReadOnly   :=SoloVer
    WinRecDat1.T_ImpReq1.ReadOnly:=SoloVer
    WinRecDat1.T_Ret1.ReadOnly   :=SoloVer
    WinRecDat1.T_ImpRet1.ReadOnly:=SoloVer
    WinRecDat1.T_TFac.ReadOnly   :=.T.
    WinRecDat1.T_Pend.ReadOnly   :=.T.
    WinRecDat1.T_Vtos.ReadOnly   :=.T.
    WinRecDat1.T_Pagos.ReadOnly  :=.T.

    WinRecDat1.GR_Vtos.Enabled :=SoloVer
    WinRecDat1.GR_Pagos.Enabled:=SoloVer

    WinRecDat1.Bt_BuscarRec.Enabled:=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_BuscarCue.Enabled:=ModifRec
    WinRecDat1.Bt_Asiento.Enabled:=IF(ModifRec=.T.,.F.,.T.)

    WinRecDat1.Bt_Alta_Vtos.Enabled :=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_Modif_Vtos.Enabled:=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_Elim_Vtos.Enabled :=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_Alta_Pago.Enabled :=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_Modif_Pago.Enabled:=IF(ModifRec=.T.,.F.,.T.)
    WinRecDat1.Bt_Elim_Pago.Enabled :=IF(ModifRec=.T.,.F.,.T.)

    IF ModifRec = .T.
        WinRecDat1.Bt_Nuevo.Enabled    := .F.
        WinRecDat1.Bt_Modif.Enabled    := .F.
        WinRecDat1.Bt_Guardar.Enabled  := .T.
        WinRecDat1.Bt_Cancelar.Enabled := .T.
        WinRecDat1.Bt_Eliminar.Enabled := .F.
        WinRecDat1.Bt_Salir.Enabled    := .F.
    ELSE
        WinRecDat1.Bt_Nuevo.Enabled    := .T.
        WinRecDat1.Bt_Modif.Enabled    := .T.
        WinRecDat1.Bt_Guardar.Enabled  := .F.
        WinRecDat1.Bt_Cancelar.Enabled := .F.
        WinRecDat1.Bt_Eliminar.Enabled := .T.
        WinRecDat1.Bt_Salir.Enabled    := .T.
    ENDIF 
Return Nil

STATIC FUNCTION RecDat_Actualizar(LLAMADA)
    AbrirDBF("FACREB",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    IF .NOT. EOF()
        WinRecDat1.T_CodCta.Value :=PCODCTA2(CODIGO)
        WinRecDat1.T_NomCta.Value :=NOMCTA
        WinRecDat1.T_CifRec.Value :=CIF
        WinRecDat1.D_FecRec.Value :=FREG
        WinRecDat1.T_NumFac.Value :=REF
        WinRecDat1.T_Asiento.Value:=ASIENTO
        WinRecDat1.C_Regimen.Value:=REGIMEN+1
        WinRecDat1.T_BImp1.Value  :=BIMP
        WinRecDat1.T_Iva1.Value   :=IVA
        WinRecDat1.T_ImpIva1.Value:=CUOTA
        WinRecDat1.T_BImp2.Value  :=BIMPT2
        WinRecDat1.T_Iva2.Value   :=IVAT2
        WinRecDat1.T_ImpIva2.Value:=CUOTAT2
        WinRecDat1.T_BImp3.Value  :=BIMPT3
        WinRecDat1.T_Iva3.Value   :=IVAT3
        WinRecDat1.T_ImpIva3.Value:=CUOTAT3
        WinRecDat1.T_Req1.Value   :=REQ
        WinRecDat1.T_ImpReq1.Value:=IMPREQ
        WinRecDat1.T_Ret1.Value   :=RET
        WinRecDat1.T_ImpRet1.Value:=IMPRET
        WinRecDat1.T_TFac.Value   :=TFAC
        WinRecDat1.T_Pend.Value   :=PEND
        RecDat_ActVto1()
        RecDat_ActPago1()
    ELSE
        RecDat_Nuevo("ACTUALIZAR")
    ENDIF

    IF LLAMADA="CANCELAR"
        RecDat_Modif(.F.)
    ENDIF
    RecDat_ActBotAsi()

Return Nil

STATIC FUNCTION RecDat_ActBotAsi()
    IF WinRecDat1.T_Asiento.Value=0
        WinRecDat1.Bt_Asiento.Caption:="Contabilizar"
    ELSE
        WinRecDat1.Bt_Asiento.Caption:="Ver asiento"
    ENDIF
Return Nil

STATIC FUNCTION RecDat_ActVto1()
    WinRecDat1.T_Vtos.Value:=0
    WinRecDat1.GR_Vtos.DeleteAllItems
    AbrirDBF("FACVTO",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    DO WHILE NREG=WinRecDat1.T_CodRec.Value .AND. .NOT. EOF()
        aVTO:={}
        AADD(aVTO,RECNO())
        AADD(aVTO,WinRecDat1.GR_Vtos.ItemCount+1)
        AADD(aVTO,IMPORTE)
        AADD(aVTO,FVTO-WinRecDat1.D_FecRec.Value)
        AADD(aVTO,FVTO)
        AADD(aVTO,BANCO)
        WinRecDat1.GR_Vtos.AddItem(aVTO)
        WinRecDat1.T_Vtos.Value:=WinRecDat1.T_Vtos.Value+IMPORTE
        SKIP
    ENDDO
    WinRecDat1.GR_Vtos.Refresh
Return Nil

STATIC FUNCTION RecDat_ActPago1()
    WinRecDat1.T_Pagos.Value:=0
    WinRecDat1.GR_Pagos.DeleteAllItems
    AbrirDBF("PAGOS",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    DO WHILE NREG=WinRecDat1.T_CodRec.Value .AND. .NOT. EOF()
        aVTO:={}
        AADD(aVTO,RECNO())
        AADD(aVTO,FPAG)
        AADD(aVTO,IMPORTE)
        AADD(aVTO,FVTO)
        AADD(aVTO,BANCO)
        AADD(aVTO,NASI)
        AADD(aVTO,NEMPASI)
        AADD(aVTO,RTRIM(DESCRIP))
        WinRecDat1.GR_Pagos.AddItem(aVTO)
        WinRecDat1.T_Pagos.Value:=WinRecDat1.T_Pagos.Value+IMPORTE
        SKIP
    ENDDO
    WinRecDat1.GR_Pagos.Refresh
Return Nil


STATIC FUNCTION RecDat_Guardar()
    IF LEN(RTRIM(WinRecDat1.T_NumFac.Value))=0
        MSGSTOP("No ha especificado el numero de factura")
        RETURN
    ENDIF
    IF LEN(RTRIM(WinRecDat1.T_CodCta.Value))=0
        MSGSTOP("No ha especificado el codigo de cuenta")
        RETURN
    ENDIF
    IF WinRecDat1.T_TFac.Value=0
        MSGSTOP("El importe de la factura es cero")
        RETURN
    ENDIF

***No duplicar la factura***
    AbrirDBF("FACREB",,,,,1)
    GO TOP
    DO WHILE .NOT. EOF()
        DO EVENTS
        IF CODIGO=VAL(PCODCTA3(WinRecDat1.T_CodCta.Value)) .AND. REF=WinRecDat1.T_NumFac.Value .AND. NREG<>WinRecDat1.T_CodRec.Value
            MsgStop("La factura esta introducida"+HB_OsNewLine()+"Registro "+LTRIM(STR(NREG))+" fecha "+DIA(FREG),"error")
            WinRecDat1.T_NumFac.SetFocus
            RETURN
        ENDIF
        SKIP
    ENDDO
***FIN No duplicar la factura***


    AbrirDBF("FACREB",,,,,1)
    IF WinRecDat1.T_CodRec.Value=0
        GO BOTT
        IF .NOT. EOF()
            WinRecDat1.T_CodRec.Value:=NREG+1
        ELSE
            WinRecDat1.T_CodRec.Value:=1
        ENDIF
        APPEND BLANK
    ELSE
        SEEK WinRecDat1.T_CodRec.Value
    ENDIF

    IF RLOCK()
        REPLACE NREG WITH WinRecDat1.T_CodRec.Value
        REPLACE CODIGO WITH VAL(PCODCTA3(WinRecDat1.T_CodCta.Value))
        REPLACE NOMCTA WITH WinRecDat1.T_NomCta.Value
        REPLACE CIF WITH WinRecDat1.T_CifRec.Value
        REPLACE FREG WITH WinRecDat1.D_FecRec.Value
        REPLACE REF WITH WinRecDat1.T_NumFac.Value
        REPLACE ASIENTO WITH WinRecDat1.T_Asiento.Value
        REPLACE REGIMEN WITH WinRecDat1.C_Regimen.Value-1
        REPLACE BIMP WITH WinRecDat1.T_BImp1.Value
        REPLACE IVA WITH WinRecDat1.T_Iva1.Value
        REPLACE CUOTA WITH WinRecDat1.T_ImpIva1.Value
        REPLACE BIMPT2 WITH WinRecDat1.T_BImp2.Value
        REPLACE IVAT2 WITH WinRecDat1.T_Iva2.Value
        REPLACE CUOTAT2 WITH WinRecDat1.T_ImpIva2.Value
        REPLACE BIMPT3 WITH WinRecDat1.T_BImp3.Value
        REPLACE IVAT3 WITH WinRecDat1.T_Iva3.Value
        REPLACE CUOTAT3 WITH WinRecDat1.T_ImpIva3.Value
        REPLACE REQ WITH WinRecDat1.T_Req1.Value
        REPLACE IMPREQ WITH WinRecDat1.T_ImpReq1.Value
        REPLACE RET WITH WinRecDat1.T_Ret1.Value
        REPLACE IMPRET WITH WinRecDat1.T_ImpRet1.Value
        REPLACE TFAC WITH WinRecDat1.T_TFac.Value
        REPLACE PEND WITH WinRecDat1.T_Pend.Value
        REPLACE NEMP WITH NUMEMP
        DBCOMMIT()
        DBUNLOCK()
    ENDIF

   ***Comprobar fecha de asiento
    IF WinRecDat1.T_Asiento.Value<>0
        AbrirDBF("Apuntes",,,,,1)
        SEEK WinRecDat1.T_Asiento.Value
        IF FECHA<>WinRecDat1.D_FecRec.Value
            IF MsgYesNo("La fecha ha sido modificada"+HB_OsNewLine()+ ;
                "Fecha factura: "+DIA(WinRecDat1.D_FecRec.Value,10)+HB_OsNewLine()+ ;
                "Fecha asiento: "+DIA(FECHA,10)+HB_OsNewLine()+ ;
                "¿Desea modificar la fecha del asiento?")=.T.
                DO WHILE NASI=WinRecDat1.T_Asiento.Value
                    IF RLOCK()
                        REPLACE FECHA WITH WinRecDat1.D_FecRec.Value
                        DBCOMMIT()
                        DBUNLOCK()
                    ENDIF
                    SKIP
                ENDDO
            ENDIF
        ENDIF
    ENDIF
   ***FIN Comprobar fecha de asiento

    MsgInfo('Los datos han sido guardados','Datos guardados')
    RecDat_Modif(.F.)

Return Nil


STATIC FUNCTION RecDat_Eliminar()
    IF MSGYESNO("¿Desea eliminar la factura?","Atencion")=.F.
        RETURN
    ENDIF

    IF WinRecDat1.T_Asiento.Value<>0
        IF MSGYESNO("¡Atencion! la factura esta contabilizada"+HB_OsNewLine()+ ;
            "¿Desea eliminar la factura?","Atencion")=.F.
            RETURN
        ENDIF
    ENDIF

    AbrirDBF("FACREB",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    IF EOF()
        MSGSTOP("No existe la factura")
        RETURN
    ENDIF

    IF RLOCK()
        DELETE
        DBCOMMIT()
        DBUNLOCK()

        AbrirDBF("FACVTO",,,,,1)
        SEEK WinRecDat1.T_CodRec.Value
        DO WHILE NREG=WinRecDat1.T_CodRec.Value .AND. .NOT. EOF()
            IF RLOCK()
                DELETE
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
            SKIP
        ENDDO
        AbrirDBF("PAGOS",,,,,1)
        SEEK WinRecDat1.T_CodRec.Value
        DO WHILE NREG=WinRecDat1.T_CodRec.Value .AND. .NOT. EOF()
            IF RLOCK()
                DELETE
                DBCOMMIT()
                DBUNLOCK()
            ENDIF
            SKIP
        ENDDO

        MsgInfo('La factura ha sido eliminada')

        RecDat_Actualizar("BORRADO")

    ELSE
        MsgStop('No se ha eliminado el registro'+HB_OsNewLine()+ ;
            'El registro esta siendo utilizado por otro usuario'+HB_OsNewLine()+ ;
            'Por favor, intentelo mas tarde','Error')
    ENDIF

Return Nil


STATIC FUNCTION RecDat_ActCta()
    AbrirDBF("CUENTAS",,,,,1)
    SEEK VAL(WinRecDat1.T_CodCta.Value)
    IF .NOT. EOF()
        WinRecDat1.T_NomCta.Value:=RTRIM(NOMCTA)
        WinRecDat1.T_CifRec.Value:=CIF
        WinRecDat1.C_Regimen.Value:=REGIMEN+1
    ELSE
        WinRecDat1.T_NomCta.Value:=""
        WinRecDat1.T_CifRec.Value:=""
        WinRecDat1.C_Regimen.Value:=1
    ENDIF
Return Nil


STATIC FUNCTION RecDat_Sumas1(LLAMADA)
    LLAMADA:=UPPER(LLAMADA)
    DO CASE
    CASE LLAMADA="BIMP1"
        WinRecDat1.T_ImpIva1.Value:=ROUND(WinRecDat1.T_BImp1.Value*WinRecDat1.T_Iva1.Value/100,2)
    CASE LLAMADA="BIMP2"
        WinRecDat1.T_ImpIva2.Value:=ROUND(WinRecDat1.T_BImp2.Value*WinRecDat1.T_Iva2.Value/100,2)
    CASE LLAMADA="BIMP3"
        WinRecDat1.T_ImpIva3.Value:=ROUND(WinRecDat1.T_BImp3.Value*WinRecDat1.T_Iva3.Value/100,2)
    CASE LLAMADA="REQ1"
        WinRecDat1.T_ImpReq1.Value:=ROUND((WinRecDat1.T_BImp1.Value+WinRecDat1.T_BImp2.Value+WinRecDat1.T_BImp3.Value)* ;
            WinRecDat1.T_Req1.Value/100,2)
    CASE LLAMADA="RET1"
        WinRecDat1.T_ImpRet1.Value:=ROUND((WinRecDat1.T_BImp1.Value+WinRecDat1.T_BImp2.Value+WinRecDat1.T_BImp3.Value)* ;
            WinRecDat1.T_Ret1.Value/100,2)
    ENDCASE
    RecDat_Sumas2()
Return Nil

STATIC FUNCTION RecDat_Sumas2()
    TFAC2:=WinRecDat1.T_TFac.Value
    WinRecDat1.T_TFac.Value:=WinRecDat1.T_BImp1.Value+WinRecDat1.T_ImpIva1.Value+ ;
        WinRecDat1.T_BImp2.Value+WinRecDat1.T_ImpIva2.Value+ ;
        WinRecDat1.T_BImp3.Value+WinRecDat1.T_ImpIva3.Value+ ;
        WinRecDat1.T_ImpReq1.Value- ;
        WinRecDat1.T_ImpRet1.Value
    WinRecDat1.T_Pend.Value:=WinRecDat1.T_Pend.Value-TFAC2+WinRecDat1.T_TFac.Value

Return Nil


STATIC FUNCTION RecDat_Asiento1()

    IF WinRecDat1.T_Asiento.Value<>0
        br_suizoasi(RUTAEMPRESA,WinRecDat1.T_Asiento.Value,NUMEMP)
        RETURN
    ENDIF

    IF MSGYESNO("¿Desea contabilizar la factura?")=.F.
        RETURN
    ENDIF

    DEFINE WINDOW WinRec2Con ;
        AT 0,0     ;
        WIDTH 600  ;
        HEIGHT 320 ;
        TITLE "Contabilizar" ;
        MODAL      ;
        NOSIZE BACKCOLOR MiColor("VERDECLARO") ;

        @15,10 LABEL L_FecAsi VALUE 'Fecha asiento' AUTOSIZE TRANSPARENT
        @10,110 DATEPICKER  D_FecAsi WIDTH 100 VALUE DATE()

        @040,010 TEXTBOX T_BImp1 WIDTH 100 VALUE WinRecDat1.T_BImp1.Value+WinRecDat1.T_BImp2.Value+WinRecDat1.T_BImp3.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @040,120 BUTTONEX Bt_CodCtaGas CAPTION 'Cuenta gasto' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaGas.Value),"WinRec2Con","T_CodCtaGas","L_CodCtaGas") ) NOTABSTOP
        @040,240 TEXTBOX T_CodCtaGas WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaGas.Value:=PCODCTA3(WinRec2Con.T_CodCtaGas.Value) , ;
            WinRec2Con.L_CodCtaGas.Value :=Codigos_NomCta(WinRec2Con.T_CodCtaGas.Value) )
        @045,350 LABEL L_CodCtaGas VALUE 'Cuenta gasto' AUTOSIZE TRANSPARENT

        @070,010 TEXTBOX T_ImpIva1 WIDTH 100 VALUE WinRecDat1.T_ImpIva1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @070,120 BUTTONEX Bt_CodCtaIva1 CAPTION 'Cuenta IVA 1' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaIva1.Value),"WinRec2Con","T_CodCtaIva1","L_CodCtaIva1") ) NOTABSTOP
        @070,240 TEXTBOX T_CodCtaIva1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaIva1.Value:=PCODCTA3(WinRec2Con.T_CodCtaIva1.Value) , ;
            WinRec2Con.L_CodCtaIva1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva1.Value) )
        @075,350 LABEL L_CodCtaIva1 VALUE 'Cuenta IVA 1' AUTOSIZE TRANSPARENT

        @100,010 TEXTBOX T_ImpIva2 WIDTH 100 VALUE WinRecDat1.T_ImpIva2.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @100,120 BUTTONEX Bt_CodCtaIva2 CAPTION 'Cuenta IVA 2' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaIva2.Value),"WinRec2Con","T_CodCtaIva2","L_CodCtaIva2") ) NOTABSTOP
        @100,240 TEXTBOX T_CodCtaIva2 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaIva2.Value:=PCODCTA3(WinRec2Con.T_CodCtaIva2.Value) , ;
            WinRec2Con.L_CodCtaIva2.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva2.Value) )
        @105,350 LABEL L_CodCtaIva2 VALUE 'Cuenta IVA 2' AUTOSIZE TRANSPARENT

        @130,010 TEXTBOX T_ImpIva3 WIDTH 100 VALUE WinRecDat1.T_ImpIva3.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @130,120 BUTTONEX Bt_CodCtaIva3 CAPTION 'Cuenta IVA 3' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaIva3.Value),"WinRec2Con","T_CodCtaIva3","L_CodCtaIva3") ) NOTABSTOP
        @130,240 TEXTBOX T_CodCtaIva3 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaIva3.Value:=PCODCTA3(WinRec2Con.T_CodCtaIva3.Value) , ;
            WinRec2Con.L_CodCtaIva3.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva3.Value) )
        @135,350 LABEL L_CodCtaIva3 VALUE 'Cuenta IVA 3' AUTOSIZE TRANSPARENT

        @160,010 TEXTBOX T_ImpReq1 WIDTH 100 VALUE WinRecDat1.T_ImpReq1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @160,120 BUTTONEX Bt_CodCtaReq1 CAPTION 'Cuenta recargo' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaReq1.Value),"WinRec2Con","T_CodCtaReq1","L_CodCtaReq1") ) NOTABSTOP
        @160,240 TEXTBOX T_CodCtaReq1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaReq1.Value:=PCODCTA3(WinRec2Con.T_CodCtaReq1.Value) , ;
            WinRec2Con.L_CodCtaReq1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaReq1.Value) )
        @165,350 LABEL L_CodCtaReq1 VALUE 'Cuenta recargo' AUTOSIZE TRANSPARENT

        @190,010 TEXTBOX T_ImpRet1 WIDTH 100 VALUE WinRecDat1.T_ImpRet1.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @190,120 BUTTONEX Bt_CodCtaRet1 CAPTION 'Cuenta retencion' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaRet1.Value),"WinRec2Con","T_CodCtaRet1","L_CodCtaRet1") ) NOTABSTOP
        @190,240 TEXTBOX T_CodCtaRet1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaRet1.Value:=PCODCTA3(WinRec2Con.T_CodCtaRet1.Value) , ;
            WinRec2Con.L_CodCtaRet1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaRet1.Value) )
        @195,350 LABEL L_CodCtaRet1 VALUE 'Cuenta retencion' AUTOSIZE TRANSPARENT

        @220,010 TEXTBOX T_TFac WIDTH 100 VALUE WinRecDat1.T_TFac.Value NUMERIC INPUTMASK '9,999,999,999.99 €' FORMAT 'E' NOTABSTOP
        @220,120 BUTTONEX Bt_CodCtaPag1 CAPTION 'Cuenta pago' WIDTH 110 HEIGHT 25 ICON icobus('buscar') ;
            ACTION ( br_cue1(VAL(WinRec2Con.T_CodCtaPag1.Value),"WinRec2Con","T_CodCtaPag1","L_CodCtaPag1") ) NOTABSTOP
        @220,240 TEXTBOX T_CodCtaPag1 WIDTH 100 MAXLENGTH 8 ;
            ON LOSTFOCUS ( WinRec2Con.T_CodCtaPag1.Value:=PCODCTA3(WinRec2Con.T_CodCtaPag1.Value) , ;
            WinRec2Con.L_CodCtaPag1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaPag1.Value) )
        @225,350 LABEL L_CodCtaPag1 VALUE 'Cuenta pago' AUTOSIZE TRANSPARENT

        MiColor_Enabled(.F.,"WinRec2Con","T_BImp1")
        MiColor_Enabled(.F.,"WinRec2Con","T_ImpIva1")
        MiColor_Enabled(.F.,"WinRec2Con","T_ImpIva2")
        MiColor_Enabled(.F.,"WinRec2Con","T_ImpIva3")
        MiColor_Enabled(.F.,"WinRec2Con","T_ImpReq1")
        MiColor_Enabled(.F.,"WinRec2Con","T_ImpRet1")
        MiColor_Enabled(.F.,"WinRec2Con","T_TFac")
        WinRec2Con.D_FecAsi.Value:=GetProperty("WinRecDat1","D_FecRec","Value")
        WinRec2Con.T_CodCtaGas.Value :="60000000"
        WinRec2Con.T_CodCtaIva1.Value:=STR(47200000+((GetProperty("WinRecDat1","C_Regimen","Value")-1)*1000)+ROUND(GetProperty("WinRecDat1","T_Iva1","Value"),0),8)
        WinRec2Con.T_CodCtaIva2.Value:=STR(47200000+ROUND(GetProperty("WinRecDat1","T_Iva2","Value"),0),8)
        WinRec2Con.T_CodCtaIva3.Value:=STR(47200000+ROUND(GetProperty("WinRecDat1","T_Iva3","Value"),0),8)
        WinRec2Con.T_CodCtaReq1.Value:=STR(47200800+((GetProperty("WinRecDat1","C_Regimen","Value")-1)*1000)+ROUND(GetProperty("WinRecDat1","T_Req1","Value"),0),8)
        WinRec2Con.T_CodCtaRet1.Value:=STR(47510000+ROUND(GetProperty("WinRecDat1","T_Ret1","Value"),0),8)
        AbrirDBF("PAGOS")
        SEEK WinRecDat1.T_CodRec.Value
        IF .NOT. EOF()
            WinRec2Con.T_CodCtaPag1.Value:=STR(BANCO)
        ELSE
            WinRec2Con.T_CodCtaPag1.Value:="57000000"
        ENDIF
        WinRec2Con.L_CodCtaGas.Value :=Codigos_NomCta(WinRec2Con.T_CodCtaGas.Value)
        WinRec2Con.L_CodCtaIva1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva1.Value)
        WinRec2Con.L_CodCtaIva2.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva2.Value)
        WinRec2Con.L_CodCtaIva3.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaIva3.Value)
        WinRec2Con.L_CodCtaReq1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaReq1.Value)
        WinRec2Con.L_CodCtaRet1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaRet1.Value)
        WinRec2Con.L_CodCtaPag1.Value:=Codigos_NomCta(WinRec2Con.T_CodCtaPag1.Value)


        @260,410 BUTTON Bt_Conta ;
            CAPTION 'Contabilizar' ;
            WIDTH 90 HEIGHT 25 ;
            ACTION RecDat_Asiento2() ;
            NOTABSTOP

        @260,510 BUTTONEX Bt_Salir ;
            CAPTION 'Salir' ;
            ICON icobus('salir') ;
            ACTION WinRec2Con.Release ;
            WIDTH 80 HEIGHT 25 ;
            NOTABSTOP

    END WINDOW
    VentanaCentrar("WinRec2Con","Ventana1")
    CENTER WINDOW WinRec2Con
    ACTIVATE WINDOW WinRec2Con

Return Nil

STATIC FUNCTION RecDat_Asiento2()

    PonerEspera("Contabilizando la factura ..." )

    CONDES2:="Fac.rec."+GetProperty("WinRecDat1","T_NumFac","Value")+";"+ ;
        PCODCTA4(GetProperty("WinRecDat1","T_CodCta","Value"))
    CONDES3:="Pago fac."+SUBSTR(CONDES2,9,LEN(CONDES2))

    AbrirDBF("CUENTAS",,,,,1)                       //ABRIR PARA SALDOS
    AbrirDBF("Apuntes",,,,,1)
    GO BOTT
    NASI2=IF(NASI<2,2,NASI+1)
    APU2:=1

    aASIENTO:={}
    AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaGas.Value),GetProperty("WinRecDat1","T_BImp1","Value"),0})
    IF GetProperty("WinRecDat1","T_ImpIva1","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaIva1.Value),GetProperty("WinRecDat1","T_ImpIva1","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_BImp2","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaGas.Value),GetProperty("WinRecDat1","T_BImp2","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_ImpIva2","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaIva2.Value),GetProperty("WinRecDat1","T_ImpIva2","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_BImp3","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaGas.Value),GetProperty("WinRecDat1","T_BImp3","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_ImpIva3","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaIva3.Value),GetProperty("WinRecDat1","T_ImpIva3","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_ImpReq1","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaReq1.Value),GetProperty("WinRecDat1","T_ImpReq1","Value"),0})
    ENDIF
    IF GetProperty("WinRecDat1","T_ImpRet1","Value")<>0
        AADD(aASIENTO,{VAL(WinRec2Con.T_CodCtaRet1.Value),0,GetProperty("WinRecDat1","T_ImpRet1","Value")})
    ENDIF
    IF GetProperty("WinRecDat1","C_Regimen","Value")-1 = 2
        CODCTA2:=STR(47700000+((GetProperty("WinRecDat1","C_Regimen","Value")-1)*1000)+ROUND(GetProperty("WinRecDat1","T_Iva1","Value"),0),8)
        AADD(aASIENTO,{VAL(CODCTA2),0,GetProperty("WinRecDat1","T_ImpIva1","Value")})
        TFAC2:=GetProperty("WinRecDat1","T_Tfac","Value")-GetProperty("WinRecDat1","T_ImpIva1","Value")
    ELSE
        TFAC2:=GetProperty("WinRecDat1","T_Tfac","Value")
    ENDIF
    AADD(aASIENTO,{VAL(PCODCTA3(GetProperty("WinRecDat1","T_CodCta","Value"))),0,TFAC2})

    FOR N=1 TO LEN(aASIENTO)
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE APU WITH N
            REPLACE NEMP WITH NUMEMP
            REPLACE FECHA WITH WinRec2Con.D_FecAsi.Value
            REPLACE CODCTA WITH aASIENTO[N,1]
            REPLACE NOMAPU WITH CONDES2
            REPLACE DEBE  WITH aASIENTO[N,2]
            REPLACE HABER WITH aASIENTO[N,3]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
        ENDIF
    NEXT

    AbrirDBF("FACREB",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    IF .NOT. EOF()
        IF RLOCK()
            REPLACE ASIENTO WITH NASI2
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
    ENDIF
    SetProperty("WinRecDat1","T_Asiento","Value",NASI2)
    RecDat_ActBotAsi()

    IF GetProperty("WinRecDat1","T_Pend","Value")<>GetProperty("WinRecDat1","T_Tfac","Value")
        AbrirDBF("PAGOS",,,,,1)
        SEEK WinRecDat1.T_CodRec.Value
        DO WHILE NREG=WinRecDat1.T_CodRec.Value
            IF NASI=0
                NASI2++
                APU2:=1
                FASI2:=MAX(FPAG,FVTO)
                AbrirDBF("Apuntes",,,,,1)
                APPEND BLANK
                IF RLOCK()
                    REPLACE NASI WITH NASI2
                    REPLACE APU WITH APU2++
                    REPLACE NEMP WITH NUMEMP
                    REPLACE FECHA WITH FASI2
                    REPLACE CODCTA WITH VAL(PCODCTA3(GetProperty("WinRecDat1","T_CodCta","Value")))
                    REPLACE NOMAPU WITH CONDES3
                    REPLACE DEBE WITH PAGOS->IMPORTE
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
                    REPLACE CODCTA WITH VAL(WinRec2Con.T_CodCtaPag1.Value)
                    REPLACE NOMAPU WITH CONDES3
                    REPLACE HABER WITH PAGOS->IMPORTE
                    DBCOMMIT()
                    DBUNLOCK()
                    Suizo_saldocuenta(CODCTA,DEBE,HABER)
                ENDIF
                AbrirDBF("PAGOS",,,,,1)
                IF RLOCK()
                    REPLACE NASI WITH NASI2
                    REPLACE FASI WITH FASI2
                    REPLACE NEMPASI WITH NUMEMP
                    DBCOMMIT()
                    DBUNLOCK()
                ENDIF
            ENDIF
            AbrirDBF("PAGOS",,,,,1)
            SKIP
        ENDDO
    ENDIF

    QuitarEspera()

    DO EVENTS
    MsgInfo("La factura se ha contabilizado correctamente"+HB_OsNewLine()+ ;
        "Empresa: "+LTRIM(STR(NUMEMP))+HB_OsNewLine()+ ;
        "Asiento: "+LTRIM(STR(GetProperty("WinRecDat1","T_Asiento","Value")))+HB_OsNewLine()+ ;
        "Fecha: "+DIA(WinRec2Con.D_FecAsi.Value,10))

    RecDat_ActPago1()
    WinRec2Con.Release

Return Nil


STATIC FUNCTION RecDat_AltaVto(LLAMADA)
    IF WinRecDat1.T_CodRec.Value=0
        MSGINFO("No se ha seleccionado ninguna factura")
        RETURN
    ENDIF

    IF LLAMADA="GRID"
        IF WinRecDat1.GR_Vtos.Value=0
            LLAMADA:="ALTA"
        ELSE
            LLAMADA:="MODIFICAR"
        ENDIF
    ENDIF

    IF LLAMADA="ELIMINAR"
        IF WinRecDat1.GR_Vtos.Value=0
            MSGINFO("No se ha seleccionado ningun vencimiento")
            RETURN
        ENDIF
        IF MSGYESNO("¿Desea eliminar el vencimiento?","Atencion")=.F.
            RETURN
        ENDIF
        aVTOS:=WinRecDat1.GR_Vtos.Item(WinRecDat1.GR_Vtos.Value)
        AbrirDBF("FACVTO")
        GO aVTOS[1]
        IF NREG<>WinRecDat1.T_CodRec.Value
            MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
            MsgInfo('El vencimiento ha sido eliminado')
        ENDIF
    ELSE

        aBANCOS1:={}
        aBANCOS2:={}
        AbrirDBF("CUENTAS",,,,,1)
        SEEK 57000000
        DO WHILE CODCTA<57499999 .AND. .NOT. EOF()
            AADD(aBANCOS1,RTRIM(NOMCTA))
            AADD(aBANCOS2,CODCTA)
            SKIP
        ENDDO
        aLabels    :={'Fec.Vto.','Importe','Banco'}
        aFormats   :={Nil,'9,999,999,999.99',aBANCOS1}
        IF LLAMADA="ALTA"
            TIT2:="Alta de vencimiento"
            aInitValues:={WinRecDat1.D_FecRec.Value,WinRecDat1.T_TFac.Value-WinRecDat1.T_Vtos.Value,1}
            aBOTON:={"Nuevo","Cancelar"}
        ELSE
            IF WinRecDat1.GR_Vtos.Value=0
                MSGINFO("No se ha seleccionado ningun vencimiento")
                RETURN
            ENDIF
            TIT2:="Modificar vencimiento"
            aVTOS:=WinRecDat1.GR_Vtos.Item(WinRecDat1.GR_Vtos.Value)
            CODBAN2:=ASCAN(aBANCOS2,{|AVAL| AVAL=aVTOS[6]})
            aInitValues:={aVTOS[5],aVTOS[3],CODBAN2}
            aBOTON:={"Modificar","Cancelar"}
        ENDIF

        AltaPag2:=InputWindow(TIT2, aLabels , aInitValues , aFormats,WinRecDat1.Row+260,WinRecDat1.Col+010 , .T. , aBOTON)

        IF AltaPag2[1] == Nil
            RETURN
        ENDIF

        IF AltaPag2[2]=0
            MSGSTOP("No se ha especificado el importe"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF

        AbrirDBF("FACVTO")
        IF LLAMADA="ALTA"
            APPEND BLANK
        ELSE
            GO aVTOS[1]
            IF NREG<>WinRecDat1.T_CodRec.Value
                MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
                RETURN
            ENDIF
        ENDIF

        CODBAN2:=IF(AltaPag2[3]=0,0, aBANCOS2[AltaPag2[3]] )
        IF RLOCK()
            REPLACE NREG WITH WinRecDat1.T_CodRec.Value
            REPLACE FVTO WITH AltaPag2[1]
            REPLACE CODIGO WITH VAL(PCODCTA3(WinRecDat1.T_CodCta.Value))
            REPLACE NOMCTA WITH WinRecDat1.T_NomCta.Value
            REPLACE REF WITH WinRecDat1.T_NumFac.Value
            REPLACE IMPORTE WITH AltaPag2[2]
            REPLACE BANCO WITH CODBAN2
            REPLACE NEMP WITH NUMEMP
            REPLACE REGANY WITH EJERANY
            DBCOMMIT()
            DBUNLOCK()
        ENDIF

    ENDIF

    RecDat_ActVto1()

Return Nil

STATIC FUNCTION RecDat_AltaPago(LLAMADA)
    IF WinRecDat1.T_CodRec.Value=0
        MSGINFO("No se ha seleccionado ninguna factura")
        RETURN
    ENDIF

    IF LLAMADA="GRID"
        IF WinRecDat1.GR_Pagos.Value=0
            LLAMADA:="ALTA"
        ELSE
            LLAMADA:="MODIFICAR"
        ENDIF
    ENDIF

    IF LLAMADA<>"ALTA"
        IF WinRecDat1.GR_Pagos.Value=0
            MSGINFO("No se ha seleccionado ningun pago")
            RETURN
        ENDIF
    ENDIF

    DO CASE
    CASE LLAMADA="VERASIENTO"
        aPAGOS:=WinRecDat1.GR_Pagos.Item(WinRecDat1.GR_Pagos.Value)
        IF aPAGOS[6]<>0
            br_suizoasi(RUTAEMPRESA,aPAGOS[6],aPAGOS[7])
        ENDIF
        RETURN
    CASE LLAMADA="ELIMINAR"
        IF MSGYESNO("¿Desea eliminar el pago?","Atencion")=.F.
            RETURN
        ENDIF
        aPAGOS:=WinRecDat1.GR_Pagos.Item(WinRecDat1.GR_Pagos.Value)
        AbrirDBF("PAGOS")
        GO aPAGOS[1]
        IF NREG<>WinRecDat1.T_CodRec.Value
            MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF
        IF RLOCK()
            DELETE
            DBCOMMIT()
            DBUNLOCK()
            MsgInfo('El pago ha sido eliminado')
        ENDIF
        WinRecDat1.T_Pend.Value:=WinRecDat1.T_Pend.Value+aPAGOS[3]

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
            CASE WinRecDat1.GR_Vtos.ItemCount>WinRecDat1.GR_Pagos.ItemCount
                FPAGO2:=WinRecDat1.GR_Vtos.Cell(WinRecDat1.GR_Pagos.ItemCount+1,5)
            CASE WinRecDat1.GR_Pagos.ItemCount>=1
                FPAGO2:=WinRecDat1.GR_Pagos.Cell(WinRecDat1.GR_Pagos.ItemCount,2)
            OTHERWISE
                FPAGO2:=WinRecDat1.D_FecRec.Value
            ENDCASE
            TIT2:="Alta de pagos"
            aInitValues:={FPAGO2,WinRecDat1.T_Pend.Value,'Pagado',WinRecDat1.D_FecRec.Value,1}
            aBOTON:={"Nuevo","Cancelar"}
        ELSE
            TIT2:="Modificar pagos"
            aPAGOS:=WinRecDat1.GR_Pagos.Item(WinRecDat1.GR_Pagos.Value)
            CODBAN2:=ASCAN(aBANCOS2,{|AVAL| AVAL=aPAGOS[5]})
            aInitValues:={aPAGOS[2],aPAGOS[3],aPAGOS[8],aPAGOS[4],CODBAN2}
            aBOTON:={"Modificar","Cancelar"}
        ENDIF

        AltaPag2:=InputWindow(TIT2, aLabels , aInitValues , aFormats,WinRecDat1.Row+260,WinRecDat1.Col+410 , .T. , aBOTON)

        IF AltaPag2[1] == Nil
            RETURN
        ENDIF

        IF AltaPag2[2]=0
            MSGSTOP("No se ha especificado el importe"+HB_OsNewLine()+"Proceso cancelado")
            RETURN
        ENDIF

        AbrirDBF("PAGOS")
        IF LLAMADA="ALTA"
            APPEND BLANK
        ELSE
            GO aPAGOS[1]
            IF NREG<>WinRecDat1.T_CodRec.Value
                MSGSTOP("La base de datos esta bloqueada"+HB_OsNewLine()+"Proceso cancelado")
                RETURN
            ENDIF
        ENDIF

        AltaPag2[4]:=IF(AltaPag2[1]>AltaPag2[4],AltaPag2[1],AltaPag2[4])
        CODBAN2:=IF(AltaPag2[5]=0,0, aBANCOS2[AltaPag2[5]] )
        IF RLOCK()
            REPLACE NREG WITH WinRecDat1.T_CodRec.Value
            REPLACE FREG WITH WinRecDat1.D_FecRec.Value
            REPLACE FPAG WITH AltaPag2[1]
            REPLACE IMPORTE WITH AltaPag2[2]
            REPLACE DESCRIP WITH AltaPag2[3]
            REPLACE FVTO WITH AltaPag2[4]
            REPLACE BANCO WITH CODBAN2
            REPLACE NEMP WITH NUMEMP
            DBCOMMIT()
            DBUNLOCK()
        ENDIF

   ***Comprobar fecha de asiento
        IF LLAMADA="MODIFICAR"
            IF aPAGOS[6]<>0
                IF aPAGOS[7]<>NUMEMP
                    RUTA1:=BUSRUTAEMP(RUTAPROGRAMA,NUMEMP,aPAGOS[7],"SUICONTA")
                    RUTA2:=RUTA1[1]
                ELSE
                    RUTA2:=RUTAEMPRESA
                ENDIF
                AbrirDBF("Apuntes",,,,RUTA2,1)
                SEEK aPAGOS[6]
                IF FECHA<>MAX(AltaPag2[1],AltaPag2[4])
                    IF MsgYesNo("La fecha ha sido modificada"+HB_OsNewLine()+ ;
                        "Fecha pago: "+DIA(MAX(AltaPag2[1],AltaPag2[4]),10)+HB_OsNewLine()+ ;
                        "Fecha asiento: "+DIA(FECHA,10)+HB_OsNewLine()+ ;
                        "¿Desea modificar la fecha del asiento?")=.T.
                        DO WHILE NASI=aPAGOS[6]
                            IF RLOCK()
                                REPLACE FECHA WITH MAX(AltaPag2[1],AltaPag2[4])
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
            WinRecDat1.T_Pend.Value:=WinRecDat1.T_Pend.Value-AltaPag2[2]
        ELSE
            WinRecDat1.T_Pend.Value:=WinRecDat1.T_Pend.Value-AltaPag2[2]+aPAGOS[3]
        ENDIF

    CASE LLAMADA="CONTABILIZAR"
        aPAGOS:=WinRecDat1.GR_Pagos.Item(WinRecDat1.GR_Pagos.Value)

        IF aPAGOS[6]<>0
            MSGSTOP("El pago ya esta contabilizado","error")
            RETURN
        ENDIF

        IF MSGYESNO("¿Desea contabilizar el pago?")=.F.
            RETURN
        ENDIF

        FASI2:=MAX(aPAGOS[2],aPAGOS[4])

        IF YEAR(FASI2)<>EJERANY
            IF MSGYESNO("¿Desea contabilizar el pago en el ejercicio "+LTRIM(STR(YEAR(FASI2)))+"?" )=.F.
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
        FASI2:=MAX(aPAGOS[2],aPAGOS[4])
        IF UPPER(aPAGOS[8])="PAGADO"
            NOMPAG2:="Pago fac."
        ELSE
		***PARA QUE EL MINIMO NO SEA CERO AÑADIR " " Y "."
            NUMAT2:=MIN(AT(" ",UPPER(aPAGOS[8])+" "),AT(".",UPPER(aPAGOS[8])+"."))
            IF NUMAT2>1
                NOMPAG2:=LEFT(aPAGOS[8],NUMAT2-1)+" fac."
            ELSE
                NOMPAG2:="Pago fac."
            ENDIF
        ENDIF

        SELEC APUNTES
        APPEND BLANK
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE APU WITH APU2++
            REPLACE NEMP WITH NUMEMP2
            REPLACE FECHA WITH FASI2
            REPLACE CODCTA WITH VAL(PCODCTA3(GetProperty("WinRecDat1","T_CodCta","Value")))
            REPLACE NOMAPU WITH NOMPAG2+GetProperty("WinRecDat1","T_NumFac","Value")+";"+PCODCTA4(aPAGOS[5])
            REPLACE DEBE WITH aPAGOS[3]
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
            REPLACE CODCTA WITH aPAGOS[5]
            REPLACE NOMAPU WITH NOMPAG2+GetProperty("WinRecDat1","T_NumFac","Value")+";"+PCODCTA4(GetProperty("WinRecDat1","T_CodCta","Value"))
            REPLACE HABER WITH aPAGOS[3]
            DBCOMMIT()
            DBUNLOCK()
            Suizo_saldocuenta(CODCTA,DEBE,HABER)
        ENDIF
        AbrirDBF("PAGOS",,,,,1)
        GO aPAGOS[1]
        IF RLOCK()
            REPLACE NASI WITH NASI2
            REPLACE FASI WITH FASI2
            REPLACE NEMPASI WITH NUMEMP2
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
        MsgInfo("El pago se ha contabilizado correctamente"+HB_OsNewLine()+ ;
            "Empresa: "+LTRIM(STR(NUMEMP2))+HB_OsNewLine()+ ;
            "Asiento: "+LTRIM(STR(NASI2))+HB_OsNewLine()+ ;
            "Fecha: "+DIA(FASI2,10))

    ENDCASE

    AbrirDBF("FACREB",,,,,1)
    SEEK WinRecDat1.T_CodRec.Value
    IF .NOT. EOF()
        IF RLOCK()
            REPLACE PEND WITH WinRecDat1.T_Pend.Value
            DBCOMMIT()
            DBUNLOCK()
        ENDIF
    ENDIF

    RecDat_ActPago1()

RETURN Nil



STATIC FUNCTION RecDat_BotCue()
    IF VAL(WinRecDat1.T_CodCta.Value)=0
        AbrirDBF("FACREB",,,,,1)
        GO BOTT
        CODCTA2:=CODIGO
    ELSE
        CODCTA2:=VAL(WinRecDat1.T_CodCta.Value)
    ENDIF
    br_cue1(CODCTA2,"WinRecDat1","T_CodCta")
    RecDat_ActCta()
    WinRecDat1.D_FecRec.SetFocus
RETURN Nil


***REGIMEN IVA***
FUNCTION REGIVAREC(REGIVA2)
    IF VALTYPE(REGIVA2)="N"
        DO CASE
        CASE REGIVA2=0
            REGIVA3:="Deducible en Operaciones Interiores"
        CASE REGIVA2=1
            REGIVA3:="Deducible en Importaciones         "
        CASE REGIVA2=2
            REGIVA3:="Deducible Adquisicion Intracomunit."
        CASE REGIVA2=3
            REGIVA3:="Compensa.Reg.Esp.Agricola,Ganaderia"
        CASE REGIVA2=4
            REGIVA3:="Regulacion de Inversiones          "
        CASE REGIVA2=5
            REGIVA3:="Deducible inversion sujeto pasivo  "
        OTHERWISE
            REGIVA3:="Error                              "
        ENDCASE
    ELSE
        REGIVA3:={"Deducible en Operaciones Interiores",;
            "Deducible en Importaciones         ",;
            "Deducible Adquisicion Intracomunit.",;
            "Compensa.Reg.Esp.Agricola,Ganaderia",;
            "Regulacion de Inversiones          ",;
            "Deducible inversion sujeto pasivo  "}
    ENDIF
    
RETURN(REGIVA3)
