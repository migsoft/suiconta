FUNCTION CREAR(FICNOM1)
Local aArq := {}
IF FILE(FICNOM1)
   RETURN
ENDIF
DO CASE
CASE AT('ALBARAN.DBF',FICNOM1)<>0
FICBASE:={{'SERALB      ','C',         1,         0},;
          {'NALB        ','N',         6,         0},;
          {'LINALB      ','N',         3,         0},;
          {'FALB        ','D',        10,         0},;
          {'CODCLI      ','N',         5,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'PRECIO      ','N',        16,         5},;
          {'DESCUENTO   ','N',         5,         2},;
          {'IVA         ','N',         5,         2},;
          {'REQ         ','N',         5,         2},;
          {'TIPIVA      ','N',         1,         0},;
          {'ACTA        ','N',        14,         3},;
          {'USUARIO     ','N',         4,         0},;
          {'VEND        ','N',         4,         0},;
          {'COMIVEN     ','N',         5,         2},;
          {'NEMP        ','N',         3,         0},;
          {'SERFAC      ','C',         1,         0},;
          {'NFAC        ','N',         6,         0},;
          {'SERPED      ','C',         1,         0},;
          {'NPED        ','N',         6,         0},;
          {'LINPED      ','N',         3,         0},;
          {'FPED        ','D',        10,         0},;
          {'TRANSP      ','C',        15,         0},;
          {'PORTES      ','N',         1,         0},;
          {'BULTOS      ','N',         2,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'DPAGO       ','C',         4,         0},;
          {'CODIGOS     ','C',        10,         0},;
          {'PCOSTE      ','N',        16,         5},;
          {'PERSON      ','C',        30,         0}}
IF FILE("\SUIZO\CONTA2\JOSE.BAT")
   AADD(FICBASE, {'NOTASDIM','C',20,0} )
ENDIF
DBCREATE(FICNOM1,FICBASE)

CASE AT('TICKET.DBF',FICNOM1)<>0
FICBASE:={{'SERTIC      ','C',         1,         0},;
          {'NTIC        ','N',         6,         0},;
          {'LINTIC      ','N',         3,         0},;
          {'FTIC        ','D',        10,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'PVENTA      ','N',        16,         5},;
          {'PCOSTE      ','N',        16,         5},;
          {'IVA         ','N',         5,         2},;
          {'REQ         ','N',         5,         2},;
          {'IMPENT      ','N',        16,         5},;
          {'FPAGO       ','N',         4,         0},;
          {'USUARIO     ','N',         4,         0},;
          {'VEND        ','N',         4,         0},;
          {'COMIVEN     ','N',         5,         2},;
          {'NEMP        ','N',         3,         0},;
          {'NASI        ','N',         6,         0},;
          {'NEMPASI     ','N',         6,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('PEDIDOS.DBF',FICNOM1)<>0
FICBASE:={{'SERPED      ','C',         1,         0},;
          {'NPED        ','N',         6,         0},;
          {'LINPED      ','N',         3,         0},;
          {'FPED        ','D',        10,         0},;
          {'CODCLI      ','N',         5,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'UNIALB      ','N',        13,         2},; //ENTREGADO
          {'PRECIO      ','N',        16,         5},;
          {'DESCUENTO   ','N',         5,         2},;
          {'IVA         ','N',         5,         2},;
          {'REQ         ','N',         5,         2},;
          {'TIPIVA      ','N',         1,         0},;
          {'ACTA        ','N',        14,         3},;
          {'USUARIO     ','N',         4,         0},;
          {'VEND        ','N',         4,         0},;
          {'COMIVEN     ','N',         5,         2},;
          {'NEMP        ','N',         3,         0},;
          {'SEROFE      ','C',         1,         0},;
          {'NOFE        ','N',         6,         0},;
          {'LINOFE      ','N',         3,         0},;
          {'SERALB      ','C',         1,         0},;
          {'NALB        ','N',         6,         0},;
          {'FALB        ','D',        10,         0},;
          {'TRANSP      ','C',        15,         0},;
          {'PORTES      ','N',         1,         0},;
          {'CODIGOS     ','C',        10,         0},;
          {'BULTOS      ','N',         2,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'DPAGO       ','C',         4,         0},;
          {'PCOSTE      ','N',        16,         5},;
          {'FENT        ','D',        10,         0},;
          {'PERSON      ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('APUNTES.DBF',FICNOM1)<>0
FICBASE:={{'NASI        ','N',         6,         0},;
          {'APU         ','N',         3,         0},;
          {'FECHA       ','D',        10,         0},;
          {'CODCTA      ','N',         8,         0},;
          {'NOMAPU      ','C',        40,         0},;
          {'DEBE        ','N',        14,         3},;
          {'HABER       ','N',        14,         3},;
          {'SALDO       ','N',        14,         3},;
          {'NEMP        ','N',         3,         0},;
          {'ASIENTO     ','C',        10,         0},;
          {'BANDERA     ','C',        15,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CIERRE.DBF',FICNOM1)<>0
FICBASE:={{'NASI        ','N',         6,         0},;
          {'APU         ','N',         3,         0},;
          {'FECHA       ','D',        10,         0},;
          {'CODCTA      ','N',         8,         0},;
          {'NOMAPU      ','C',        40,         0},;
          {'DEBE        ','N',        14,         3},;
          {'HABER       ','N',        14,         3},;
          {'SALDO       ','N',        14,         3},;
          {'NEMP        ','N',         3,         0},;
          {'ASIENTO     ','C',        10,         0},;
          {'BANDERA     ','C',        15,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('ARTICULO.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','C',        10,         0},;
          {'ARTICULO    ','C',        30,         0},;
          {'PRECIO      ','N',        16,         5},;
          {'PCOSTE      ','N',        16,         5},;
          {'MARGEN      ','N',        10,         2},;
          {'ACTMARGEN   ','L',         1,         0},;
          {'IVA         ','N',         5,         2},;
          {'UNIDADES    ','N',        13,         2},;
          {'REPRE       ','N',         5,         0},;
          {'FAMI        ','N',         3,         0},;
          {'RESERVA     ','N',        13,         2},;
          {'PEDPROV     ','N',        13,         2},;
          {'UNIMINIMO   ','N',        13,         2},;
          {'FVENTA      ','D',        10,         0},;
          {'FCOMPRA     ','D',        10,         0},;
          {'SUBCTA      ','N',         8,         0},;
          {'REFPROV     ','C',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CODBAR.DBF',FICNOM1)<>0
FICBASE:={{'CODART      ','C',        10,         0},;
          {'CODBAR      ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CAJA.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         6,         0},;
          {'FECHA       ','D',        10,         0},;
          {'CODCTA      ','N',         8,         0},;
          {'CONCEPTO    ','C',        30,         0},;
          {'DEBE        ','N',        14,         3},;
          {'HABER       ','N',        14,         3},;
          {'SALDO       ','N',        14,         3},;
          {'SERIE       ','N',         2,         0},;
          {'EMP         ','N',         3,         0},;
          {'ASIENTO     ','N',         6,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CLIENTES.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         5,         0},;
          {'CLIENTE     ','C',        30,         0},;
          {'COMERCIAL   ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'PAIS        ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'CIF         ','C',        15,         0},;
          {'TEL1        ','C',        15,         0},;
          {'FAX         ','C',        15,         0},;
          {'MOVIL       ','C',        15,         0},;
          {'PAGO        ','N',         4,         0},;
          {'DPAGO       ','C',         4,         0},;
          {'DIAVTO      ','C',        20,         0},;
          {'RET         ','N',         5,         2},;
          {'BANNOM      ','C',        30,         0},;
          {'BANDIR      ','C',        30,         0},;
          {'BANPOB      ','C',        30,         0},;
          {'BANCTA      ','C',        30,         0},;
          {'RIESGO      ','N',        14,         3},;
          {'VEND        ','N',         3,         0},;
          {'ZONA        ','N',         3,         0},;
          {'RUTA        ','N',         3,         0},;
          {'ACTIVI      ','N',         3,         0},;
          {'REPRE       ','N',         3,         0},;
          {'OBS         ','C',        30,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'ENVCLI      ','C',        30,         0},;
          {'ENVDIR      ','C',        30,         0},;
          {'ENVPOB      ','C',        30,         0},;
          {'ENVPROV     ','C',        30,         0},;
          {'ENVCP       ','C',         5,         0},;
          {'TEL2        ','C',        15,         0},;
          {'FAX2        ','C',        15,         0},;
          {'TRANSP      ','C',        15,         0},;
          {'SUBCTA      ','N',         8,         0},;
          {'FALTA       ','D',        10,         0},;
          {'FMOD        ','D',        10,         0},;
          {'CODIGOS     ','C',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CLINF.DBF',FICNOM1)<>0
FICBASE:={{'TEL1        ','C',        15,         0},;
          {'CLIENTE     ','C',        30,         0},;
          {'COMERCIAL   ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'FAX         ','C',        15,         0},;
          {'MOVIL       ','C',        15,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'OBS         ','C',        30,         0},;
          {'FALTA       ','D',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('COBROS.DBF',FICNOM1)<>0
FICBASE:={{'SERFAC      ','C',         1,         0},;
          {'NFAC        ','N',         6,         0},;
          {'FFAC        ','D',        10,         0},;
          {'FCOB        ','D',        10,         0},;
          {'IMPORTE     ','N',        14,         3},;
          {'DESCRIP     ','C',        30,         0},;
          {'FVTO        ','D',        10,         0},;
          {'BANCO       ','N',         8,         0},;
          {'NEMP        ','N',         5,         0},;
          {'NASI        ','N',         6,         0},;
          {'NEMPASI     ','N',         5,         0},;
          {'FASI        ','D',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CODIGOS.DBF',FICNOM1)<>0 .OR. AT('TOPCOD.DBF',FICNOM1)<>0
FICBASE:={{'GRUPO       ','N',         2,         0},;
          {'COD         ','N',         4,         0},;
          {'NOMCOD      ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'CIF         ','C',        15,         0},;
          {'TEL1        ','C',        15,         0},;
          {'FAX         ','C',        15,         0},;
          {'MOVIL       ','C',        15,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'MES1        ','N',        16,         5},;
          {'MES2        ','N',        16,         5},;
          {'MES3        ','N',        16,         5},;
          {'MES4        ','N',        16,         5},;
          {'MES5        ','N',        16,         5},;
          {'MES6        ','N',        16,         5},;
          {'MES7        ','N',        16,         5},;
          {'MES8        ','N',        16,         5},;
          {'MES9        ','N',        16,         5},;
          {'MES10       ','N',        16,         5},;
          {'MES11       ','N',        16,         5},;
          {'MES12       ','N',        16,         5},;
          {'MEST        ','N',        16,         5}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('CUENTAS.DBF',FICNOM1)<>0
FICBASE:={{'CODCTA      ','N',         8,         0},;
          {'NOMCTA      ','C',        40,         0},;
          {'DEBE        ','N',        14,         3},;
          {'HABER       ','N',        14,         3},;
          {'SALDO       ','N',        14,         3},;
          {'DIRCTA      ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'POBCTA      ','C',        30,         0},;
          {'PROCTA      ','C',        30,         0},;
          {'PAIS        ','C',        30,         0},;
          {'REGIMEN     ','N',         4,         0},;
          {'CIF         ','C',        15,         0},;
          {'TEL1        ','C',        15,         0},;
          {'FAX1        ','C',        15,         0},;
          {'BANCTA      ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('EMPRESA.DBF',FICNOM1)<>0 .OR. AT('EMP.DBF',FICNOM1)<>0 .OR. AT('EMPCON.DBF',FICNOM1)<>0
FICBASE:={{'NEMP        ','N',         3,         0},;
          {'EMP         ','C',        30,         0},;
          {'EJERCICIO   ','N',         4,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'PAIS        ','C',        30,         0},;
          {'CIF         ','C',        15,         0},;
          {'TEL1        ','C',        15,         0},;
          {'FAX1        ','C',        15,         0},;
          {'MOVIL       ','C',        15,         0},;
          {'BANCTA      ','C',        30,         0},;
          {'OBS         ','C',        30,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'NEMPANT     ','N',         3,         0},;
          {'NEMPSIG     ','N',         3,         0},;
          {'NEMPBUS     ','N',         3,         0},;
          {'CIERREJER   ','N',         1,         0},;
          {'RUTA        ','C',        15,         0},;
          {'RUTACONTA   ','C',        50,         0},;
          {'RUTALOGO    ','C',        50,         0},;
          {'IMPRESORA   ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FACREB.DBF',FICNOM1)<>0
FICBASE:={{'NREG        ','N',         5,         0},;
          {'FREG        ','D',        10,         0},;
          {'CODIGO      ','N',         8,         0},;
          {'NOMCTA      ','C',        40,         0},;
          {'CIF         ','C',        15,         0},;
          {'REF         ','C',        15,         0},;
          {'BIMP        ','N',        14,         3},;
          {'IVA         ','N',         5,         2},;
          {'CUOTA       ','N',        14,         3},;
          {'BIMPT2      ','N',        14,         3},;
          {'IVAT2       ','N',         5,         2},;
          {'CUOTAT2     ','N',        14,         3},;
          {'BIMPT3      ','N',        14,         3},;
          {'IVAT3       ','N',         5,         2},;
          {'CUOTAT3     ','N',        14,         3},;
          {'REQ         ','N',         5,         2},;
          {'IMPREQ      ','N',        14,         3},;
          {'RET         ','N',         5,         2},;
          {'IMPRET      ','N',        14,         3},;
          {'TFAC        ','N',        14,         3},;
          {'PEND        ','N',        14,         3},;
          {'ASIENTO     ','N',         6,         0},;
          {'REGIMEN     ','N',         4,         0},;
          {'NEMP        ','N',         5,         0},;
          {'NUMREG      ','N',         5,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FACTURAS.DBF',FICNOM1)<>0 .OR. AT('FAC92.DBF',FICNOM1)<>0
FICBASE:={{'SERFAC      ','C',         1,         0},;
          {'NFAC        ','N',         6,         0},;
          {'FFAC        ','D',        10,         0},;
          {'COD         ','N',         5,         0},;
          {'CLIENTE     ','C',        30,         0},;
          {'IMPBRUTO    ','N',        14,         3},;
          {'DTO1        ','N',         5,         2},;
          {'IMPDTO1     ','N',        14,         3},;
          {'BIMP        ','N',        14,         3},;
          {'IVA         ','N',         5,         2},;
          {'IMPIVA      ','N',        14,         3},;
          {'BIMPT2      ','N',        14,         3},;
          {'IVAT2       ','N',         5,         2},;
          {'IMPIVAT2    ','N',        14,         3},;
          {'BIMPT3      ','N',        14,         3},;
          {'IVAT3       ','N',         5,         2},;
          {'IMPIVAT3    ','N',        14,         3},;
          {'REQ         ','N',         5,         2},;
          {'IMPREQ      ','N',        14,         3},;
          {'RET         ','N',         5,         2},;
          {'IMPRET      ','N',        14,         3},;
          {'TFAC        ','N',        14,         3},;
          {'PEND        ','N',        14,         3},;
          {'ASIENTO     ','N',         6,         0},;
          {'REGIMEN     ','N',         4,         0},;
          {'REF         ','C',        30,         0},;
          {'CODCTA      ','N',         8,         0},;
          {'EMP         ','N',         5,         0},;
          {'NEMP        ','N',         5,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'DPAGO       ','C',         4,         0},;
          {'ACTA        ','N',        14,         3},;
          {'USUARIO     ','N',         4,         0},;
          {'VEND        ','N',         4,         0},;
          {'COMIVEN     ','N',         5,         2}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FACVTO.DBF',FICNOM1)<>0
FICBASE:={{'NREG        ','N',         5,         0},;
          {'FVTO        ','D',        10,         0},;
          {'CODIGO      ','N',         8,         0},;
          {'NOMCTA      ','C',        40,         0},;
          {'REF         ','C',        15,         0},;
          {'IMPORTE     ','N',        14,         3},;
          {'BANCO       ','N',         8,         0},;
          {'NEMP        ','N',         5,         0},;
          {'REGANY      ','N',         4,         0},;
          {'NUMREG      ','N',         5,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('INFORM.DBF',FICNOM1)<>0
FICBASE:={{'VEND        ','N',         3,         0},;
          {'FECHA       ','D',        10,         0},;
          {'CODCLI      ','N',         5,         0},;
          {'TCODCLI     ','C',         6,         0},;
          {'TEL1        ','C',        15,         0},;
          {'NOMCLI      ','C',        30,         0},;
          {'POBCLI      ','C',        30,         0},;
          {'REPRE       ','N',         3,         0},;
          {'CODVEN      ','N',         3,         0},;
          {'CONTACTO    ','C',        20,         0},;
          {'GESTION     ','C',        55,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('OFERTAS.DBF',FICNOM1)<>0
FICBASE:={{'SEROFE      ','C',         1,         0},;
          {'NOFE        ','N',         6,         0},;
          {'FOFE        ','D',        10,         0},;
          {'CONFIR      ','L',         1,         0},;
          {'MATERIAL    ','C',        30,         0},;
          {'IMPORTE     ','N',        14,         3},;
          {'DESCUENTO   ','N',         5,         2},;
          {'IVA         ','N',         5,         2},;
          {'REQ         ','N',         5,         2},;
          {'TIPIVA      ','N',         1,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'VEND        ','N',         3,         0},;
          {'USUARIO     ','N',         4,         0},;
          {'COMIVEN     ','N',         5,         2},;
          {'NEMP        ','N',         3,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'ACTA        ','N',        14,         3},;
          {'PORTES      ','N',         1,         0},;
          {'FENT        ','D',        10,         0},;
          {'FVALIDA     ','D',        10,         0},;
          {'CODCLI      ','N',         5,         0},;
          {'CLIENTE     ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'TEL1        ','C',        12,         0},;
          {'FAX         ','C',        12,         0},;
          {'MOVIL       ','C',        12,         0},;
          {'EMAIL       ','C',        40,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('OFERTAL.DBF',FICNOM1)<>0
FICBASE:={{'SEROFE      ','C',         1,         0},;
          {'NOFE        ','N',         6,         0},;
          {'LINOFE      ','N',         3,         0},;
          {'CODCLI      ','N',         5,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'UNIPED      ','N',        13,         2},; //PEDIDO
          {'PRECIO      ','N',        16,         5},;
          {'PCOSTE      ','N',        16,         5},;
          {'NEMP        ','N',         3,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('PAGOS.DBF',FICNOM1)<>0
FICBASE:={{'NREG        ','N',         5,         0},;
          {'FREG        ','D',        10,         0},;
          {'FPAG        ','D',        10,         0},;
          {'IMPORTE     ','N',        14,         3},;
          {'DESCRIP     ','C',        30,         0},;
          {'FVTO        ','D',        10,         0},;
          {'BANCO       ','N',         8,         0},;
          {'NEMP        ','N',         5,         0},;
          {'NASI        ','N',         6,         0},;
          {'NEMPASI     ','N',         5,         0},;
          {'FASI        ','D',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('PROVEE.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         5,         0},;
          {'PROVEE      ','C',        30,         0},;
          {'COMERCIAL   ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'PAIS        ','C',        30,         0},;
          {'CODPOS      ','C',         5,         0},;
          {'CIF         ','C',        15,         0},;
          {'TEL1        ','C',        15,         0},;
          {'TEL2        ','C',        15,         0},;
          {'FAX1        ','C',        15,         0},;
          {'OBS         ','C',        30,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'REPRE       ','N',         3,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'BANCTA      ','C',        30,         0},;
          {'TIVA        ','N',         1,         0},;
          {'DTO1        ','N',         5,         2},;
          {'TRANSP      ','C',        15,         0},;
          {'SUBCTA      ','N',         8,         0},;
          {'PORTES      ','N',         1,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('PROVENV.DBF',FICNOM1)<>0
FICBASE:={{'CODPROV     ','N',         5,         0},;
          {'CODENV      ','N',         2,         0},;
          {'PREFERIDA   ','N',         1,         0},;
          {'ENVNOM      ','C',        30,         0},;
          {'ENVDIR      ','C',        30,         0},;
          {'ENVPOB      ','C',        30,         0},;
          {'ENVPROV     ','C',        30,         0},;
          {'ENVCP       ','C',         5,         0},;
          {'ENVTEL1     ','C',        15,         0},;
          {'ENVFAX1     ','C',        15,         0},;
          {'CONTACTO    ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('RARTI.DBF',FICNOM1)<>0
FICBASE:={{'NREC        ','C',        13,         0},;
          {'FREC        ','D',        10,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'PRECIO      ','N',        16,         5},;
          {'NENT        ','N',         5,         0},; //PARA PASAR A ENTRADA
          {'FENT        ','D',        10,         0},;
          {'PROV        ','N',         5,         0},;
          {'NALB        ','C',        10,         0},;
          {'FALB        ','D',        10,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('ENTRADA.DBF',FICNOM1)<>0
FICBASE:={{'SERENT      ','C',         1,         0},;
          {'NENT        ','N',         6,         0},;
          {'LINENT      ','N',         3,         0},;
          {'FENT        ','D',        10,         0},;
          {'PROV        ','N',         5,         0},;
          {'NALB        ','C',        10,         0},;
          {'FALB        ','D',        10,         0},;
          {'NEMPSUI     ','N',         5,         0},;
          {'NREGPROV    ','N',         5,         0},;
          {'NFRAPROV    ','C',        15,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'PRECIO      ','N',        16,         5},;
          {'PREIDTO     ','N',        16,         5},;
          {'REFPROV     ','C',        10,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'TIVA        ','N',         1,         0},;
          {'DTO1        ','N',         5,         2},;
          {'TRANSP      ','C',        15,         0},;
          {'PORTES      ','N',         1,         0},;
          {'BULTOS      ','N',         2,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'SERDEM      ','C',         1,         0},;
          {'NDEM        ','N',         6,         0},;
          {'LINDEM      ','N',         3,         0},;
          {'FDEM        ','D',        10,         0},;
          {'NEMP        ','N',         3,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('DEMANDA.DBF',FICNOM1)<>0
FICBASE:={{'SERDEM      ','C',         1,         0},;
          {'NDEM        ','N',         6,         0},;
          {'LINDEM      ','N',         3,         0},;
          {'FDEM        ','D',        10,         0},;
          {'PROV        ','N',         5,         0},;
          {'FPLAZO      ','D',        10,         0},;
          {'CODART      ','C',        10,         0},;
          {'NOMART      ','C',        30,         0},;
          {'UNIDAD      ','N',        13,         2},;
          {'UNIENT      ','N',        13,         2},; //RECIBIDO
          {'PRECIO      ','N',        16,         5},;
          {'PREIDTO     ','N',        16,         5},;
          {'REFPROV     ','C',        10,         0},;
          {'FPAGO       ','N',         4,         0},;
          {'TIVA        ','N',         1,         0},;
          {'DTO1        ','N',         5,         2},;
          {'TRANSP      ','C',        15,         0},;
          {'PORTES      ','N',         1,         0},;
          {'BULTOS      ','N',         2,         0},;
          {'CONTACTO    ','C',        30,         0},;
          {'SERALB      ','C',         1,         0},;
          {'NALB        ','N',         6,         0},;
          {'NEMP        ','N',         5,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('REMESA.DBF',FICNOM1)<>0
FICBASE:={{'SERIE       ','N',         2,         0},;
          {'NREM        ','N',         5,         0},;
          {'LREM        ','N',         3,         0},;
          {'FREM        ','D',        10,         0},;
          {'NEMPREG     ','N',         3,         0},;
          {'NREG        ','N',         5,         0},;
          {'NFRA        ','C',        15,         0},;
          {'CODCTA      ','N',         8,         0},;
          {'NOMCTA      ','C',        30,         0},;
          {'IMPORTE     ','N',        14,         3},;
          {'BANCTA      ','C',        30,         0},;
          {'TALON       ','C',        30,         0},;
          {'FVTO        ','D',        10,         0},;
          {'CODBAN      ','N',         8,         0},;
          {'NASI        ','N',         6,         0},;
          {'TRANOMI     ','N',         1,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('REMLIN.DBF',FICNOM1)<>0
FICBASE:={{'SERIE       ','N',         2,         0},;
          {'NREM        ','N',         5,         0},;
          {'LREM        ','N',         3,         0},;
          {'LINEA       ','C',        40,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FORMATOS.DBF',FICNOM1)<>0
FICBASE:={{'CODFOR      ','N',         3,         0},;
          {'NOMFOR      ','C',        30,         0},;
          {'TIPFOR      ','C',         6,         0},;
          {'FORDEF      ','L',         1,         0},;
          {'DirMarSup   ','N',         3,         0},;
          {'DirMarIzq   ','N',         3,         0},;
          {'LogMarSup   ','N',         3,         0},;
          {'LogMarIzq   ','N',         3,         0},;
          {'CabMarSup   ','N',         3,         0},;
          {'CabMarIzq   ','N',         3,         0},;
          {'CueMarSup   ','N',         3,         0},;
          {'CueMarIzq   ','N',         3,         0},;
          {'PieMarSup   ','N',         3,         0},;
          {'PieMarIzq   ','N',         3,         0},;
          {'LinCuerpo   ','N',         3,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FTOREM.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         3,         0},;
          {'FORMATO     ','C',        30,         0},;
          {'FICDAT      ','C',        12,         0},;
          {'LINPAG      ','N',         3,         0},;
          {'LINCAB      ','N',         3,         0},;
          {'LINCPO      ','N',         3,         0},;
          {'PAGSAL      ','N',         3,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('MASCARA.DBF',FICNOM1)<>0
FICBASE:={{'BUSCAR      ','C',        10,         0},;
          {'COD         ','N',         3,         0},;
          {'LIN         ','N',         3,         0},;
          {'POS         ','N',         3,         0},;
          {'ESP         ','N',         3,         0},;
          {'CAM         ','N',         3,         0},;
          {'NOMCAM      ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('BANCOS.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         5,         0},;
          {'NOMBRE      ','C',        30,         0},;
          {'DIRECCION   ','C',        30,         0},;
          {'POBLACION   ','C',        30,         0},;
          {'PROVINCIA   ','C',        30,         0},;
          {'CODCTA      ','C',        30,         0},;
          {'TEL1        ','C',        12,         0},;
          {'FAX1        ','C',        12,         0},;
          {'EMPNOM      ','C',        30,         0},;
          {'EMPDIR      ','C',        30,         0},;
          {'EMPPOB      ','C',        30,         0},;
          {'EMPPROV     ','C',        30,         0},;
          {'EMPTEL1     ','C',        12,         0},;
          {'EMPFAX1     ','C',        12,         0},;
          {'EMPCIF      ','C',        12,         0},;
          {'MPROG       ','C',         3,         0},;
          {'NEMP        ','N',         5,         0},;
          {'EMPEJER     ','N',         4,         0},;
          {'FMODEMP     ','D',        10,         0},;
          {'RUTAPROG    ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('SERIES.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         2,         0},;
          {'NOMCOD      ','C',        30,         0},;
          {'MONEDA      ','C',        10,         0},;
          {'DECIMAL     ','N',         2,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('LIGAS.DBF',FICNOM1)<>0
FICBASE:={{'CODLIG      ','N',         5,         0},;
          {'NOMLIG      ','C',        35,         0},;
          {'ANY         ','N',         4,         0},;
          {'NCLUBS      ','N',         3,         0},;
          {'DV          ','N',         1,         0},;
          {'DES1        ','N',         1,         0},;
          {'DES2        ','N',         1,         0},;
          {'DES3        ','N',         1,         0},;
          {'DES4        ','N',         1,         0},;
          {'DES5        ','N',         1,         0},;
          {'PUNMAX      ','N',         5,         1},;
          {'PUNG        ','N',         5,         1},;
          {'PUNT        ','N',         5,         1},;
          {'PUNP        ','N',         5,         1},;
          {'ASCENSO     ','N',         3,         0},;
          {'DESCENSO    ','N',         3,         0},;
          {'PROMO       ','N',         3,         0},;
          {'FICHERO     ','C',         8,         0},;
          {'FICHTM      ','C',         8,         0},;
          {'RESALTAR    ','N',         5,         0},;
          {'TITLIG      ','C',        70,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('TORNEOS.DBF',FICNOM1)<>0
FICBASE:={{'CODTOR      ','N',         5,         0},;
          {'NOMTOR      ','C',        35,         0},;
          {'ANY         ','N',         4,         0},;
          {'RONDAS      ','N',         3,         0},;
          {'DES1        ','N',         1,         0},;
          {'DES2        ','N',         1,         0},;
          {'DES3        ','N',         1,         0},;
          {'PUNG        ','N',         3,         1},;
          {'PUNT        ','N',         3,         1},;
          {'PUNP        ','N',         3,         1},;
          {'PUND        ','N',         3,         1},;
          {'NINC        ','N',         1,         0},;
          {'FINAL       ','N',         1,         0},;
          {'FICHERO     ','C',         8,         0},;
          {'TITTOR      ','C',        70,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('AGENDA.DBF',FICNOM1)<>0
FICBASE:={{'COD         ','N',         5,         0},;
          {'NOMBRE      ','C',        35,         0},;
          {'NOMTEL      ','C',        35,         0},;
          {'TEL1        ','C',        15,         0},;
          {'MOVIL       ','C',        15,         0},;
          {'FAX1        ','C',        15,         0},;
          {'NOMEMAIL    ','C',        35,         0},;
          {'EMAIL       ','C',        40,         0},;
          {'DIRNOM      ','C',        35,         0},;
          {'DIRDIR      ','C',        35,         0},;
          {'DIRPOB      ','C',        35,         0},;
          {'DIRPROV     ','C',        35,         0},;
          {'FECNAC      ','D',        10,         0},;
          {'NOTA1       ','C',        40,         0},;
          {'NOTA2       ','C',        40,         0},;
          {'PROGRAMA    ','C',        30,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('PGC.DBF',FICNOM1)<>0
FICBASE:={{'CODPGC      ','C',         4,         0},;
          {'NOMPGC1     ','C',        40,         0},;
          {'NOMPGC2     ','C',       100,         0}}
DBCREATE(FICNOM1,FICBASE)

CASE AT('FIESTAS.DBF',FICNOM1)<>0 //esta duplicado en LIB y LIB2
   Aadd( aArq , { 'FECFIE'    , 'D' , 10  , 0 } ) //fecha
   Aadd( aArq , { 'FIESTA'    , 'N' ,  1  , 0 } ) //abierto, cerrado, fiesta
   Aadd( aArq , { 'ANUAL'     , 'N' ,  1  , 0 } ) //todos los años, solo este
   Aadd( aArq , { 'NOMFIESTA' , 'C' , 30  , 0 } ) //nombre de la fiesta
   DBCreate( FICNOM1 , aArq  )

ENDCASE
