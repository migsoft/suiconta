#include 'minigui.ch'
#include 'miniprint.ch'
#include 'winprint.ch'

#include 'hbclass.ch'
#include 'common.ch'
#include "fileio.ch"

*-------------------------
FUNCTION TPrint( clibx )
*-------------------------
LOCAL o_Print_
if clibx=NIL
   if type("_HMG_printlibrary")="C"
      if _HMG_printlibrary="HBPRINTER"
         o_print_:=thbprinter()
      elseif _HMG_printlibrary="MINIPRINT"
         o_print_:=tminiprint()
      elseif _HMG_printlibrary="DOSPRINT"
         o_print_:=tdosprint()
      elseif _HMG_printlibrary="CALCPRINT"
         o_print_:=tcalcprint()
      elseif _HMG_printlibrary="EXCELPRINT"
         o_print_:=texcelprint()
      elseif _HMG_printlibrary="RTFPRINT"
         o_print_:=trtfprint()
      elseif _HMG_printlibrary="CSVPRINT"
         o_print_:=tcsvprint()
      elseif _HMG_printlibrary="HTMLPRINT"
         o_print_:=thtmlprint()
      elseif _HMG_printlibrary="PDFPRINT"
         o_print_:=tpdfprint()
      else
         o_print_:=thbprinter()
      endif      
   else
      o_print_:=thbprinter()
      _HMG_printlibrary:="HBPRINTER"
   endif
else
   if valtype(clibx)="C"
      clibx:=UPPER(clibx)
      if clibx="HBPRINTER"
         o_print_:=thbprinter()
      elseif clibx="MINIPRINT"
         o_print_:=tminiprint()
      elseif clibx="DOSPRINT"
         o_print_:=tdosprint()
      elseif clibx="CALCPRINT"
         o_print_:=tcalcprint()
      elseif clibx="EXCELPRINT"
         o_print_:=texcelprint()
      elseif clibx="RTFPRINT"
         o_print_:=trtfprint()
      elseif clibx="CSVPRINT"
         o_print_:=tcsvprint()
      elseif clibx="HTMLPRINT"
         o_print_:=thtmlprint()
      elseif clibx="PDFPRINT"
         o_print_:=tpdfprint()
      else
         o_print_:=thbprinter()
      endif
    else
      o_print_:=thbprinter()
      _HMG_printlibrary:="HBPRINTER"
    endif
endif
RETURN o_Print_
   
   


//////////////////////////////////////
CREATE CLASS TPRINTBASE

DATA cprintlibrary      INIT "HBPRINTER"  READONLY
DATA nmhor              INIT (10)/4.75    READONLY
DATA nmver              INIT (10)/2.35    READONLY
DATA nhfij              INIT (12/3.70)    READONLY
DATA nvfij              INIT (12/1.65)    READONLY
DATA cunits             INIT "ROWCOL"     READONLY
DATA nlmargin           INIT 0
DATA ntmargin           INIT 0
DATA cprinter           INIT ""           READONLY
   
DATA aprinters          INIT {}           READONLY
DATA aports             INIT {}           READONLY
   
DATA lprerror           INIT .F.          READONLY
DATA exit               INIT .F.          READONLY
DATA acolor             INIT {0,0,0}      READONLY
DATA cfontname          INIT "Courier New" READONLY
DATA nfontsize          INIT 12            /////// nunca ponerle read only
DATA nwpen              INIT 0.1          READONLY //// ancho del pen
DATA tempfile           INIT gettempdir()+"T"+alltrim(str(int(hb_random(999999)),8))+".prn" READONLY
DATA impreview          INIT .F.          READONLY
DATA lwinhide           INIT .T.          READONLY
DATA cversion           INIT "MiTPRINT v2.5a" READONLY
DATA cargo              INIT  .F.
////DATA cString            INIT  ""

DATA nlinpag            INIT 0            READONLY
DATA alincelda          INIT {}           READONLY
DATA cunitslin          INIT 1            READONLY
DATA lprop              INIT .F.          READONLY


*-------------------------
METHOD init()
*-------------------------

*-------------------------
method initx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setprop()
*-------------------------

*-------------------------
METHOD setcpl()
*-------------------------

*-------------------------
METHOD begindoc()
*-------------------------

*-------------------------
METHOD begindocx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD enddoc()
*-------------------------

*-------------------------
METHOD enddocx() BLOCK { || nil }
*-------------------------

*-------------------------
method printdos()
*-------------------------

*-------------------------
method printdosx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD beginpage()
*-------------------------

*-------------------------
METHOD beginpagex() BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD condendos()  BLOCK { || nil }
*-------------------------

*-------------------------
METHOD condendosx()  BLOCK { || nil }
*-------------------------
   

*-------------------------
METHOD NORMALDOS()  BLOCK { || nil }
*-------------------------

*-------------------------
METHOD NORMALDOSx()  BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD endpage()
*-------------------------

*-------------------------
METHOD endpagex() BLOCK { || nil }
*-------------------------
*-------------------------
METHOD release()
*-------------------------

*-------------------------
METHOD releasex() BLOCK { || nil }
*-------------------------
*-------------------------
METHOD printdata()
*-------------------------

*-------------------------
METHOD printdatax() BLOCK { || nil }
*-------------------------
*-------------------------
METHOD printimage()
*-------------------------

*-------------------------
METHOD printimagex() BLOCK { || nil }
*-------------------------
*-------------------------
METHOD printline()
*-------------------------

*-------------------------
 METHOD printlinex() BLOCK { || nil }
*-------------------------
   

*-------------------------
METHOD printrectangle()
*-------------------------

*-------------------------
METHOD printrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinter()
*-------------------------

*-------------------------
METHOD selprinterx() BLOCK { || nil }
*-------------------------
  
*-------------------------
METHOD getdefprinter()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD setcolor()
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD setpreviewsize()
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD setunits()   ////// mm o rowcol
*-------------------------
   
*-------------------------
METHOD setunitsx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectangle()
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------
   
*-------------------------
METHOD version()  INLINE ::cversion
*-------------------------
   
*-------------------------
METHOD setlmargin()
*-------------------------

*-------------------------
METHOD settmargin()
*-------------------------

ENDCLASS

*-------------------------

*-------------------------
METHOD setpreviewsize(ntam) CLASS TPRINTBASE
*-------------------------
if ntam=NIL .or. ntam>5
   ntam:=3
endif
::setpreviewsizex(ntam)
return self


*-------------------------
METHOD setprop(lmode) CLASS TPRINTBASE
*-------------------------
DEFAULT lmode to .F.
if lmode
   ::lprop:=.T.
else
   ::lprop:=.F.
endif
return nil

*-------------------------
METHOD setcpl(ncpl) CLASS TPRINTBASE
*-------------------------
do case
   case ncpl=60
        ::nfontsize:=14
   case ncpl=80
        ::nfontsize:=12
   case ncpl=96
        ::nfontsize:=10
   case ncpl=120
        ::nfontsize:=8
   case ncpl=140
        ::nfontsize:=7
   case ncpl=160
        ::nfontsize:=6
   otherwise
        ::nfontsize:=12
endcase
return nil

*-------------------------
METHOD release() CLASS TPRINTBASE
*-------------------------
if ::exit
   return nil
endif
////setinteractiveclose(::cargo)
::releasex()
return nil

*-------------------------
METHOD init() CLASS TPRINTBASE
*-------------------------
if _IsWindowDefined ( "_HMG_PRINTER_SHOWPREVIEW" )
   msgstop("Print preview pending, close first","Stop")
   ::exit:=.T.
   return
endif
// public _HMG_printer_docname
::initx()
return self

*-------------------------
METHOD selprinter( lselect , lpreview, llandscape , npapersize ,cprinterx, nres,nWidth,nHeight ) CLASS TPRINTBASE
*-------------------------
local lsucess := .T.
if ::exit
   ::lprerror:=.T.
   return nil
endif
/*
if lhide#NIL
   ::lwinhide:=lhide
endif
*/
SETPRC(0,0)
DEFAULT llandscape to .F.
DEFAULT nWidth to 0
DEFAULT nHeight to 0

::selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,  nres ,nWidth,nHeight)
RETURN self


*-------------------------
METHOD BEGINDOC(cdoc) CLASS TPRINTBASE
*-------------------------
DEFAULT cDoc to "MiniGui printing"
** Anulado al activar el preview cierra la aplicacion
**	cuando cierra la ventana no elimina el TIMER
/*
IF ::lwinhide=.T.
   DEFINE WINDOW TprintEspera  ;
   AT 0,0 ;
   WIDTH 400 HEIGHT 120 ;
   TITLE cdoc MODAL NOSIZE NOSYSMENU NOCAPTION
   
   @ 5,5 frame myframe width 390 height 110
   
   @ 15,10 IMAGE IMAGE_101 picture 'hbprint_print' WIDTH 25 HEIGHT 30 STRETCH
   @ 22,40 LABEL LABEL_101 VALUE '' WIDTH 350 FONT "Courier New" SIZE 10
   @ 55,10 LABEL label_1 value cdoc WIDTH 380 HEIGHT 32 FONT "Courier New"

   DEFINE TIMER TIMER_101  ;
   INTERVAL 1000  ;
   ACTION action_timer()

   end window
   center window TprintEspera
   activate window TprintEspera NOWAIT
ENDIF
*/
DO EVENTS
::begindocx(cdoc)
return self


METHOD setlmargin(nlmar) CLASS TPRINTBASE
::nlmargin := nlmar
RETURN self

METHOD settmargin(ntmar) CLASS TPRINTBASE
::ntmargin := ntmar
RETURN self

/*
*-------------------------
static function action_timer()
*-------------------------
IF iswindowdefined(TprintEspera)
   TprintEspera.label_1.fontbold  := IIF(TprintEspera.label_1.fontbold,.F.,.T.)
   TprintEspera.IMAGE_101.picture := IIF(TprintEspera.label_1.fontbold,'hbprint_print','hbprint_next')
   TprintEspera.LABEL_101.Value:=PADR(".",IF(LEN(TprintEspera.LABEL_101.Value)=40,1,LEN(TprintEspera.LABEL_101.Value)+1),".")
ENDIF 
RETURN nil
*/

*-------------------------
METHOD ENDDOC() CLASS TPRINTBASE
*-------------------------
::enddocx()
/*
IF iswindowdefined(TprintEspera)
   TprintEspera.release
ENDIF
*/
return self


*-------------------------
METHOD SETCOLOR(atColor) CLASS TPRINTBASE
*-------------------------
::acolor:=atColor
::setcolorx()
return self

*-------------------------
METHOD beginPAGE(nWidth,nHeight) CLASS TPRINTBASE
*-------------------------
DEFAULT nWidth to 0
DEFAULT nHeight to 0
::beginpagex(nWidth,nHeight)
return self

*-------------------------
METHOD ENDPAGE() CLASS TPRINTBASE
*-------------------------
::endpagex()
return self

*-------------------------
METHOD getdefprinter() CLASS TPRINTBASE
*-------------------------
RETURN ::getdefprinterx()

*-------------------------
METHOD setunits(cunitsx) CLASS TPRINTBASE
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
RETURN self

*-------------------------
METHOD printdata(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,nangle,litalic) CLASS TPRINTBASE
*-------------------------
local ctext,cspace,uAux,cTipo:=Valtype(data)
do While cTipo == "B"       
   uAux:=EVal(data)
   cTipo:=valType(uAux)
   data:=uAux
enddo
do case
case cTipo == 'C'
   ctext:=data
case cTipo == 'N'
   ctext:=alltrim(str(data))
case cTipo == 'D'
   ctext:=dtoc(data)
case Ctipo == 'L'
   ctext:= iif(data,'T','F')
case cTipo == 'M'
   ctext:=data
case cTipo == 'O'
   ctext:="< Object >"
case cTipo == 'A'
   ctext:="< Array >"
otherwise
   ctext:=""
endcase

DEFAULT calign to "L"
DEFAULT nlen to 15

do case
case calign = "C"
   cspace=  space( (int(nlen)-len(ctext))/2 )
case calign = "R"
   cspace = space(int(nlen)-len(ctext))
otherwise
   cspace:=""
endcase

DEFAULT nlin to 1
DEFAULT ncol to 1
DEFAULT ctext to ""
DEFAULT lbold to .F.
DEFAULT litalic to .F.
DEFAULT cfont to ::cfontname
DEFAULT nsize to ::nfontsize
DEFAULT acolor to ::acolor
DEFAULT nangle to 0

if ::cunits="MM"
   ::nmver:=1
   ::nvfij:=0
   ::nmhor:=1
   ::nhfij:=0
else
   ::nmhor  := nsize/4.75
   if ::lprop
      ::nmver  := (::nfontsize)/2.35
   else
      ::nmver  :=  10/2.35
   endif
   ::nvfij  := (12/1.65)
   ::nhfij  := (12/3.70)
endif

if ::cunits="MM"
   ctext:= ctext
else
   ctext:= cspace + ctext
endif
::printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic)
return self


*-------------------------
METHOD printimage(nlin,ncol,nlinf,ncolf,cimage) CLASS TPRINTBASE
*-------------------------
DEFAULT nlin to 1
DEFAULT ncol to 1
DEFAULT cimage to ""
DEFAULT nlinf to 4
DEFAULT ncolf to 4

if ::cunits="MM"
   ::nmver:=1
   ::nvfij:=0
   ::nmhor:=1
   ::nhfij:=0
else
   ::nmhor  := (::nfontsize)/4.75
   if ::lprop
      ::nmver  := (::nfontsize)/2.35
   else
      ::nmver  :=  10/2.35
   endif

   ::nvfij  := (12/1.65)
   ::nhfij  := (12/3.70)
endif
::printimagex(nlin,ncol,nlinf,ncolf,cimage)
return self


*-------------------------
METHOD printline(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TPRINTBASE
*-------------------------
DEFAULT nlin to 1
DEFAULT ncol to 1
DEFAULT nlinf to 4
DEFAULT ncolf to 4
DEFAULT atcolor to ::acolor
DEFAULT ntwpen to ::nwpen

if ::cunits="MM"
   ::nmver:=1
   ::nvfij:=0
   ::nmhor:=1
   ::nhfij:=0
else
   ::nmhor  := (::nfontsize)/4.75
   if ::lprop
      ::nmver  := (::nfontsize)/2.35
   else
      ::nmver  :=  10/2.35
   endif

   ::nvfij  := (12/1.65)
   ::nhfij  := (12/3.70)
endif
::printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen )
return self

*-------------------------
METHOD printrectangle(nlin,ncol,nlinf,ncolf,atcolor,ntwpen,arcolor ) CLASS TPRINTBASE
*-------------------------

DEFAULT nlin to 1
DEFAULT ncol to 1
DEFAULT nlinf to 4
DEFAULT ncolf to 4
DEFAULT atcolor to ::acolor
DEFAULT ntwpen to ::nwpen
DEFAULT arcolor to {255,255,255}

if ::cunits="MM"
   ::nmver:=1
   ::nvfij:=0
   ::nmhor:=1
   ::nhfij:=0
else
   ::nmhor  := (::nfontsize)/4.75

   if ::lprop
      ::nmver  := (::nfontsize)/2.35
   else
      ::nmver  :=  10/2.35
   endif


   ::nvfij  := (12/1.65)
   ::nhfij  := (12/3.70)
endif
::printrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen,arcolor )

return self


*-------------------------
METHOD printroundrectangle(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TPRINTBASE
*-------------------------
DEFAULT nlin to 1
DEFAULT ncol to 1
DEFAULT nlinf to 4
DEFAULT ncolf to 4
DEFAULT atcolor to ::acolor
DEFAULT ntwpen to ::nwpen

if ::cunits="MM"
   ::nmver:=1
   ::nvfij:=0
   ::nmhor:=1
   ::nhfij:=0
else
   ::nmhor  := (::nfontsize)/4.75

   if ::lprop
      ::nmver  := (::nfontsize)/2.35
   else
      ::nmver  :=  10/2.35
   endif

   ::nvfij  := (12/1.65)
   ::nhfij  := (12/3.70)
endif

::printroundrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen )

return self

*-------------------------
method printdos() CLASS TPRINTBASE
*-------------------------
local cbat, nHdl
cbat:='b'+alltrim(str(random(999999),6))+'.bat'
nHdl := FCREATE( cBat )
FWRITE( nHdl, "copy " + ::tempfile + " prn" + CHR( 13 ) + CHR( 10 ) )
FWRITE( nHdl, "rem comando auxiliar de impresion" + CHR( 13 ) + CHR( 10 ) )
FCLOSE( nHdl )
waitrun( cBat, 0 )
erase &cbat
return nil





//////////////////////////////////////MiniPrint
CREATE CLASS TMINIPRINT FROM TPRINTBASE


*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex()
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimagex()
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx()
*-------------------------

*-------------------------
METHOD printroundrectanglex()
*-------------------------
ENDCLASS

*-------------------------
METHOD initx() CLASS TMINIPRINT
*-------------------------
public _HMG_PRINTER_APRINTERPROPERTIES
public _HMG_PRINTER_HDC
public _HMG_PRINTER_COPIES
public _HMG_PRINTER_COLLATE
public _HMG_PRINTER_PREVIEW
public _HMG_PRINTER_TIMESTAMP
public _HMG_PRINTER_NAME
public _HMG_PRINTER_PAGECOUNT
public _HMG_PRINTER_HDC_BAK

::aprinters:=aprinters()
::cprintlibrary:="MINIPRINT"
return self


*-------------------------
METHOD begindocx(cdoc) CLASS TMINIPRINT
*-------------------------
if cDoc#Nil
   START PRINTDOC NAME cDoc
else
   START PRINTDOC
endif
return self

*-------------------------
METHOD enddocx() CLASS TMINIPRINT
*-------------------------
END PRINTDOC
return self

*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TMINIPRINT
*-------------------------
START PRINTPAGE
return self

*-------------------------
METHOD endpagex() CLASS TMINIPRINT
*-------------------------
END PRINTPAGE
return self


*-------------------------
METHOD releasex() CLASS TMINIPRINT
*-------------------------
release _HMG_PRINTER_APRINTERPROPERTIES
release _HMG_PRINTER_HDC
release _HMG_PRINTER_COPIES
release _HMG_PRINTER_COLLATE
release _HMG_PRINTER_PREVIEW
release _HMG_PRINTER_TIMESTAMP
release _HMG_PRINTER_NAME
release _HMG_PRINTER_PAGECOUNT
release _HMG_PRINTER_HDC_BAK
return nil

*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TMINIPRINT
*-------------------------
Empty( Data )
DEFAULT aColor to ::acolor
Empty( nLen )
if litalic
if ::cunits="MM"
   if .not. lbold
      if calign="R"
   @ nlin, ncol PRINT (ctext) font cfont size nsize ITALIC COLOR acolor RIGHT
      elseif calign="C"
   @ nlin, ncol PRINT (ctext) font cfont size nsize ITALIC COLOR acolor CENTER
      else
   @ nlin, ncol PRINT (ctext) font cfont size nsize ITALIC COLOR acolor
      endif
   else
      if calign="R"
   @ nlin, ncol PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor RIGHT
      elseif calign="C"
   @ nlin, ncol PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor CENTER
      else
   @ nlin, ncol PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor
      endif
   endif
else
   if .not. lbold
      if calign="R"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize ITALIC COLOR acolor RIGHT
      elseif calign="C"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize ITALIC COLOR acolor CENTER
      else
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 PRINT (ctext) font cfont size nsize ITALIC COLOR acolor
      endif
   else
      if calign="R"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor RIGHT
      elseif calign="C"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor CENTER
      else
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 PRINT (ctext) font cfont size nsize BOLD ITALIC COLOR acolor
      endif
   endif
endif
else
if ::cunits="MM"
   if .not. lbold
      if calign="R"
   @ nlin, ncol PRINT (ctext) font cfont size nsize COLOR acolor RIGHT
      elseif calign="C"
   @ nlin, ncol PRINT (ctext) font cfont size nsize COLOR acolor CENTER
      else
   @ nlin, ncol PRINT (ctext) font cfont size nsize COLOR acolor
      endif
   else
      if calign="R"
   @ nlin, ncol PRINT (ctext) font cfont size nsize  BOLD COLOR acolor RIGHT
      elseif calign="C"
   @ nlin, ncol PRINT (ctext) font cfont size nsize  BOLD COLOR acolor CENTER
      else
   @ nlin, ncol PRINT (ctext) font cfont size nsize  BOLD COLOR acolor
      endif
   endif
else
   if .not. lbold
      if calign="R"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize COLOR acolor RIGHT
      elseif calign="C"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize COLOR acolor CENTER
      else
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 PRINT (ctext) font cfont size nsize COLOR acolor
      endif
   else
      if calign="R"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize  BOLD COLOR acolor RIGHT
      elseif calign="C"
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 +(len(ctext))*::nmhor  PRINT (ctext) font cfont size nsize  BOLD COLOR acolor CENTER
      else
   @ nlin*::nmver+::nvfij, ncol*::nmhor+ ::nhfij*2 PRINT (ctext) font cfont size nsize  BOLD COLOR acolor
      endif
   endif
endif
endif
return self

*-------------------------
METHOD printimagex(nlin,ncol,nlinf,ncolf,cimage) CLASS TMINIPRINT
*-------------------------
if ::cunits="MM"
@  nlin,ncol PRINT IMAGE cimage WIDTH ncolf-ncol HEIGHT nlinf - nlin
else
@  nlin*::nmver+::nvfij , ncol*::nmhor+::nhfij*2 PRINT IMAGE cimage WIDTH ((ncolf - ncol-1)*::nmhor + ::nhfij) HEIGHT ((nlinf+0.5 - nlin)*::nmver+::nvfij)
endif
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TMINIPRINT
*-------------------------
default atColor to ::acolor
if ::cunits="MM"
@ nlin,ncol PRINT LINE TO nlinf,ncolf  COLOR atcolor PEnWidth ntwpen  //// CPEN
else
@ (nlin+.2)*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 PRINT LINE TO  (nlinf+.2)*::nmver+::nvfij,ncolf*::nmhor+::nhfij*2  COLOR atcolor PEnWidth ntwpen  //// CPEN
endif
return self

*-------------------------
METHOD printrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen,arcolor ) CLASS TMINIPRINT
*-------------------------
if ::cunits="MM"
@ nlin,ncol PRINT RECTANGLE TO  nlinf,ncolf COLOR atcolor PEnWidth ntwpen  //// CPEN
   IF LEN(arcolor)=3
   @ nlin+((nlinf-nlin)/4),ncol+((nlinf-nlin)/4) PRINT RECTANGLE TO  nlinf-((nlinf-nlin)/4),ncolf-((nlinf-nlin)/4) COLOR arcolor PEnWidth ((nlinf-nlin)/2)  //// CPEN
   ENDIF
else
@ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 PRINT RECTANGLE TO  (nlinf+0.5)*::nmver+::nvfij,ncolf*::nmhor+::nhfij*2 COLOR atcolor  PEnWidth ntwpen  //// CPEN
   IF LEN(arcolor)=3
   nlin2:=nlin*::nmver+::nvfij
   nlin2f:=(nlinf+0.5)*::nmver+::nvfij
   ncol2:=ncol*::nmhor+::nhfij*2
   ncol2f:=ncolf*::nmhor+::nhfij*2
   @ nlin2+((nlin2f-nlin2)/4),ncol2+((nlin2f-nlin2)/4) PRINT RECTANGLE TO  nlin2f-((nlin2f-nlin2)/4),ncol2f-((nlin2f-nlin2)/4) COLOR arcolor PEnWidth ((nlin2f-nlin2)/2)  //// CPEN
   ENDIF
endif
return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TMINIPRINT
*-------------------------
local worientation,lsucess

if nres=NIL
   nres:=PRINTER_RES_MEDIUM
endif
IF llandscape
   Worientation:= PRINTER_ORIENT_LANDSCAPE
ELSE
   Worientation:= PRINTER_ORIENT_PORTRAIT
ENDIF

IF lselect
   cprinterx := NIL
ENDIF

/*
PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth
*/

IF lselect .and. lpreview .and. cprinterx = NIL
   ::cPrinter := GetPrinter()
   If Empty (::cPrinter)
      ::lprerror:=.T.
      Return Nil
   EndIf
   if npapersize#NIL
      SELECT PRINTER ::cprinter to lsucess ;
      ORIENTATION worientation ;
      PAPERSIZE npapersize       ;
      QUALITY nres ;
      PREVIEW
   else
		IF nWidth<>0 .AND. nHeight<>0
	      SELECT PRINTER ::cprinter to lsucess ;
	      ORIENTATION worientation ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
	      QUALITY nres ;
	      PREVIEW
		ELSE
	      SELECT PRINTER ::cprinter to lsucess ;
	      ORIENTATION worientation ;
	      QUALITY nres ;
	      PREVIEW
		ENDIF
   endif
endif

if (.not. lselect) .and. lpreview .and. cprinterx = NIL

   if npapersize#NIL
      SELECT PRINTER DEFAULT TO lsucess ;
      ORIENTATION worientation  ;
      PAPERSIZE npapersize       ;
      QUALITY nres ;
      PREVIEW
   else
		IF nWidth<>0 .AND. nHeight<>0
			SELECT PRINTER DEFAULT TO lsucess ;
			ORIENTATION worientation  ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
			QUALITY nres ;
			PREVIEW
		ELSE
			SELECT PRINTER DEFAULT TO lsucess ;
			ORIENTATION worientation  ;
			QUALITY nres ;
			PREVIEW
		ENDIF
   endif
endif

if (.not. lselect) .and. (.not. lpreview) .and. cprinterx = NIL

   if npapersize#NIL
      SELECT PRINTER DEFAULT TO lsucess  ;
      ORIENTATION worientation  ;
      QUALITY nres ;
      PAPERSIZE npapersize
   else
		IF nWidth<>0 .AND. nHeight<>0
	      SELECT PRINTER DEFAULT TO lsucess  ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
	      QUALITY nres ;
	      ORIENTATION worientation
		ELSE
	      SELECT PRINTER DEFAULT TO lsucess  ;
	      QUALITY nres ;
	      ORIENTATION worientation
		ENDIF
   endif
endif

IF lselect .and. .not. lpreview .and. cprinterx = NIL
   ::cPrinter := GetPrinter()
   If Empty (::cPrinter)
      ::lprerror:=.T.
      Return Nil
   EndIf

   if npapersize#NIL
      SELECT PRINTER ::cprinter to lsucess ;
      ORIENTATION worientation ;
      QUALITY nres ;
      PAPERSIZE npapersize
   else
		IF nWidth<>0 .AND. nHeight<>0
	      SELECT PRINTER ::cprinter to lsucess ;
	      ORIENTATION worientation ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
	      QUALITY nres
		ELSE
	      SELECT PRINTER ::cprinter to lsucess ;
	      ORIENTATION worientation ;
	      QUALITY nres
		ENDIF
   endif
endif

if cprinterx # NIL .AND. lpreview
   If Empty (cprinterx)
      ::lprerror:=.T.
      Return Nil
   EndIf

   if npapersize#NIL
      SELECT PRINTER cprinterx to lsucess ;
      ORIENTATION worientation ;
      QUALITY nres ;
      PAPERSIZE npapersize ;
      PREVIEW
   else
		IF nWidth<>0 .AND. nHeight<>0
	      SELECT PRINTER cprinterx to lsucess ;
	      ORIENTATION worientation ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
	      QUALITY nres ;
	      PREVIEW
		ELSE
	      SELECT PRINTER cprinterx to lsucess ;
	      ORIENTATION worientation ;
	      QUALITY nres ;
	      PREVIEW
		ENDIF
   endif
endif

if cprinterx # NIL .AND. .not. lpreview
   If Empty (cprinterx)
      ::lprerror:=.T.
      Return Nil
   EndIf

   if npapersize#NIL
      SELECT PRINTER cprinterx to lsucess ;
      ORIENTATION worientation ;
      QUALITY nres ;
      PAPERSIZE npapersize
   else
		IF nWidth<>0 .AND. nHeight<>0
	      SELECT PRINTER cprinterx to lsucess ;
	      ORIENTATION worientation ;
			PAPERSIZE PRINTER_PAPER_USER PAPERLENGTH nHeight PAPERWIDTH nWidth ;
	      QUALITY nres
		ELSE
	      SELECT PRINTER cprinterx to lsucess ;
	      ORIENTATION worientation ;
	      QUALITY nres
		ENDIF
   endif
endif

IF .NOT. lsucess
   ::lprerror:=.T.
   return nil
ENDIF
return self

*-------------------------
METHOD getdefprinterx() CLASS TMINIPRINT
*-------------------------
return getDefaultPrinter()


*-------------------------
METHOD printroundrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TMINIPRINT
*-------------------------
default atColor to ::acolor
@  nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 PRINT RECTANGLE TO  (nlinf+0.5)*::nmver+::nvfij,ncolf*::nmhor+::nhfij*2 COLOR atcolor  PEnWidth ntwpen  ROUNDED //// CPEN
return self





//////////////////////////////////////HbPrint
CREATE CLASS THBPRINTER FROM TPRINTBASE


*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex()
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimagex()
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx()
*-------------------------

*-------------------------
METHOD setcolorx()
*-------------------------

*-------------------------
METHOD setpreviewsizex()
*-------------------------

*-------------------------
METHOD printroundrectanglex()
*-------------------------
ENDCLASS


*-------------------------
METHOD INITx() CLASS THBPRINTER
*-------------------------
public hbprn

INIT PRINTSYS
GET PRINTERS TO ::aprinters
GET PORTS TO ::aports
SET UNITS MM
::cprintlibrary:="HBPRINTER"
return self


*-------------------------
METHOD BEGINDOCx (cdoc) CLASS THBPRINTER
*-------------------------
::setpreviewsize(2)
START DOC NAME cDoc
 define font "F0" name "courier new" size 10
 define font "F1" name "courier new" size 10 BOLD


 define font "F2" name "courier new" size 10 ITALIC
 define font "F3" name "courier new" size 10 BOLD ITALIC
return self


*-------------------------
METHOD ENDDOCx() CLASS THBPRINTER
*-------------------------
END DOC
return self


*-------------------------
METHOD BEGINPAGEx(nWidth,nHeight) CLASS THBPRINTER
*-------------------------
START PAGE
return self


*-------------------------
METHOD ENDPAGEx() CLASS THBPRINTER
*-------------------------
END PAGE
return self


*-------------------------
METHOD RELEASEx() CLASS THBPRINTER
*-------------------------
RELEASE PRINTSYS
RELEASE HBPRN
return self


*-------------------------
METHOD PRINTDATAx(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS THBPRINTER
*-------------------------
Empty( Data )

default aColor to ::acolor
Empty( nLen )

select font "F0"
change font "F0" name cfont size nsize angle nangle

IF lbold
   select font "F1"
   change font "F1" name cfont size nsize angle nangle BOLD
ENDIF 

IF litalic
   IF lbold
      select font "F3"
      change font "F3" name cfont size nsize angle nangle BOLD ITALIC
   ELSE
      select font "F2"
      change font "F2" name cfont size nsize angle nangle ITALIC
   ENDIF 
ENDIF 

SET TEXTCOLOR acolor
if ::cunits="MM"
   if .not. lbold
      if calign="R"
   @ nlin,ncol SAY (ctext) ALIGN TA_RIGHT TO PRINT
      elseif calign="C"
   @ nlin,ncol SAY (ctext) ALIGN TA_CENTER TO PRINT
      else
   @ nlin,ncol SAY (ctext) TO PRINT
      endif
   else
      if calign="R"
   @ nlin,ncol  SAY (ctext) ALIGN TA_RIGHT TO PRINT
      elseif calign="C"
   @ nlin,ncol  SAY (ctext) ALIGN TA_CENTER TO PRINT
      else
   @ nlin,ncol  SAY (ctext) TO PRINT
      endif
   endif
else
   if .not. lbold
      if calign="R"
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 +(len(ctext))*::nmhor  SAY (ctext) ALIGN TA_RIGHT TO PRINT
      elseif calign="C"
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 +(len(ctext))*::nmhor  SAY (ctext) ALIGN TA_CENTER TO PRINT
      else
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 SAY (ctext) TO PRINT
      endif
   else
      if calign="R"
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 +(len(ctext))*::nmhor  SAY (ctext) ALIGN TA_RIGHT TO PRINT
      elseif calign="C"
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 +(len(ctext))*::nmhor  SAY (ctext) ALIGN TA_CENTER TO PRINT
      else
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2  SAY (ctext) TO PRINT
      endif
   endif
endif
return self


*-------------------------
METHOD printimagex(nlin,ncol,nlinf,ncolf,cimage) CLASS thbprinter
*-------------------------
if ::cunits="MM"
   @ nlin,ncol PICTURE cimage SIZE nlinf-nlin,ncolf-ncol
else
   @ nlin*::nmver+::nvfij ,ncol*::nmhor+::nhfij*2 PICTURE cimage SIZE (nlinf+0.5-nlin-4)*::nmver+::nvfij , (ncolf-ncol-3)*::nmhor+::nhfij*2
endif
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS thbprinter
*-------------------------
default atColor to ::acolor
CHANGE PEN "C0" WIDTH ntwpen*10  COLOR atcolor
SELECT PEN "C0"
if ::cunits="MM"
   @ nlin,ncol,nlinf,ncolf LINE PEN "C0"  //// CPEN
else
   @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2 , (nlinf)*::nmver+::nvfij,ncolf*::nmhor+::nhfij*2  LINE PEN "C0"  //// CPEN
endif
return self


*-------------------------
METHOD printrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen,arcolor ) CLASS THBPRINTER
*-------------------------
default atColor to ::acolor
CHANGE PEN "C0" WIDTH ntwpen*10 COLOR atcolor
SELECT PEN "C0"
IF LEN(arcolor)=3
   CHANGE brush "BR01" style PS_SOLID color arcolor
ENDIF
if ::cunits="MM"
   IF LEN(arcolor)=3
      @ nlin,ncol, nlinf, ncolf  RECTANGLE PEN "C0" BRUSH "BR01"
   ELSE
      @ nlin,ncol, nlinf, ncolf  RECTANGLE PEN "C0" //// CPEN  RECTANGLE  ///// [PEN <cpen>] [BRUSH <cbrush>]
   ENDIF
else
   IF LEN(arcolor)=3
      @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2, (nlinf+0.5)*::nmver+::nvfij, ncolf*::nmhor+::nhfij*2  RECTANGLE  PEN "C0" BRUSH "BR01"
   ELSE
      @ nlin*::nmver+::nvfij,ncol*::nmhor+::nhfij*2, (nlinf+0.5)*::nmver+::nvfij, ncolf*::nmhor+::nhfij*2  RECTANGLE  PEN "C0" //// CPEN  RECTANGLE  ///// [PEN <cpen>] [BRUSH <cbrush>]
   ENDIF
endif
return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight ) CLASS THBPRINTER
*-------------------------

DO CASE
CASE lselect .and. lpreview
   SELECT BY DIALOG PREVIEW
CASE lselect .and. (.not. lpreview)
   SELECT BY DIALOG
CASE (.not. lselect) .and. lpreview
   IF cprinterx=NIL
      SELECT DEFAULT PREVIEW
   ELSE
      SELECT PRINTER cprinterx PREVIEW
   ENDIF
CASE (.not. lselect) .and. (.not. lpreview)
   IF cprinterx=NIL
      SELECT DEFAULT
   ELSE
      SELECT PRINTER cprinterx
   ENDIF
OTHERWISE
   SELECT DEFAULT
ENDCASE


IF HBPRNERROR != 0
   ::lprerror:=.T.
   return nil
ENDIF
define font "f0" name ::cfontname size ::nfontsize
define font "f1" name ::cfontname size ::nfontsize BOLD
define pen "C0" WIDTH ::nwpen COLOR ::acolor
select pen "C0"
define brush "BR01" style PS_SOLID color {255,255,255}

if llandscape
   set page orientation DMORIENT_LANDSCAPE font "f0"
else
   set page orientation DMORIENT_PORTRAIT  font "f0"
endif

IF nWidth<>0 .AND. nHeight<>0
	SET USER PAPERSIZE WIDTH nWidth*10 HEIGHT nHeight*10
else
	if npapersize#NIL
	   set page papersize npapersize
	endif
ENDIF

if nres#NIL
   SET QUALITY nres   ////:=PRINTER_RES_MEDIUM
endif

return self


*-------------------------
METHOD getdefprinterx() CLASS THBPRINTER
*-------------------------
local cdefprinter
GET DEFAULT PRINTER TO cdefprinter
return cdefprinter


*-------------------------
METHOD setcolorx() CLASS THBPRINTER
*-------------------------
CHANGE PEN "C0" WIDTH ::nwpen COLOR ::acolor
SELECT PEN "C0"
return self


*-------------------------
METHOD setpreviewsizex(ntam) CLASS THBPRINTER
*-------------------------
SET PREVIEW SCALE ntam
return self


*-------------------------
METHOD printroundrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS THBPRINTER
*-------------------------
default atColor to ::acolor
CHANGE PEN "C0" WIDTH ntwpen*10 COLOR atcolor
SELECT PEN "C0"
hbprn:RoundRect( nlin*::nmver+::nvfij  ,ncol*::nmhor+::nhfij*2 ,(nlinf+0.5)*::nmver+::nvfij ,ncolf*::nmhor+::nhfij*2 ,10, 10,"C0")
return self









//////////////////////////////////////Dos
CREATE CLASS TDOSPRINT FROM TPRINTBASE


DATA cString INIT ""
DATA cbusca INIT ""
DATA nOccur INIT 0

*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimage() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx()
*-------------------------

*-------------------------
method normaldosx()
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

method searchstring()

method nextsearch()

ENDCLASS



*-------------------------
METHOD initx() CLASS TDOSPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="DOSPRINT"
return self


*-------------------------
METHOD begindocx() CLASS TDOSPRINT
*-------------------------
SET PRINTER TO &(::tempfile)
SET DEVICE TO PRINT
return self

*-------------------------
METHOD enddocx() CLASS TDOSPRINT
*-------------------------
local _nhandle,wr,nx,ny

nx:=getdesktopwidth()
ny:=getdesktopheight()

SET DEVICE TO SCREEN
SET PRINTER TO
_nhandle:=FOPEN(::tempfile,0+64)
if ::impreview
   wr:=memoread((::tempfile))
   DEFINE WINDOW PRINT_PREVIEW  ;
   AT 0,0 ;
   WIDTH nx HEIGHT ny-70 ;
   TITLE 'Preview -----> ' + ::tempfile ;
   MODAL

   @ 0,0 RICHEDITBOX EDIT_P ;
   OF PRINT_PREVIEW ;
   WIDTH nx-50 ;
   HEIGHT ny-40-70 ;
   VALUE WR ;
   READONLY ;
   FONT 'Courier New' ;
   SIZE 10 ;
   BACKCOLOR WHITE

   @ 010,nx-40 button but_4 caption "X" width 30 action ( print_preview.release() ) tooltip "close"
   @ 090,nx-40 button but_1 caption "+ +" width 30 action zoom("+") tooltip "zoom +"
   @ 170,nx-40 button but_2 caption "- -" width 30 action zoom("-") tooltip "zoom -"
   @ 250,nx-40 button but_3 caption "P" width 30 action (::printdos()) tooltip "Print DOS mode"
   @ 330,nx-40 button but_5 caption "S" width 30 action  (::searchstring(print_preview.edit_p.value)) tooltip "Search"
   @ 410,nx-40 button but_6 caption "N" width 30 action  ::nextsearch() tooltip "Next Search"

   END WINDOW

   CENTER WINDOW PRINT_PREVIEW
   ACTIVATE WINDOW PRINT_PREVIEW

else

  ::PRINTDOS()

endif

IF FILE(::tempfile)
   fclose(_nhandle)
   ERASE &(::tempfile)
ENDIF

RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TDOSPRINT
*-------------------------
@ 0,0 SAY ""
return self


*-------------------------
METHOD endpagex() CLASS TDOSPRINT
*-------------------------
EJECT
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS TDOSPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self

*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TDOSPRINT
*-------------------------
Empty( Data )
Empty( cFont )
Empty( nSize )
Empty( aColor )
Empty( cAlign )
Empty( nLen )
Empty( nangle)
Empty( lItalic)

if .not. lbold
   @ nlin,ncol say (ctext)
else
   @ nlin,ncol say (ctext)
   @ nlin,ncol say (ctext)
endif
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf /* ,atcolor,ntwpen */ ) CLASS TDOSPRINT
*-------------------------
if nlin=nlinf
   @ nlin,ncol say replicate("-",ncolf-ncol+1)
endif
return self


*-------------------------
METHOD selprinterx( lselect , lpreview /* , llandscape , npapersize ,cprinterx,nres,nWidth,nHeight */ ) CLASS TDOSPRINT
*-------------------------
Empty( lSelect )
do while file(::tempfile)
   ::tempfile:=gettempdir()+"T"+alltrim(str(int(hb_random(999999)),8))+".prn"
enddo
if lpreview
   ::impreview:=.T.
endif
return self


*-------------------------
METHOD condendosx() CLASS TDOSPRINT
*-------------------------
@ prow(), pcol() say chr(15)
return self


*-------------------------
METHOD normaldosx() CLASS TDOSPRINT
*-------------------------
@ prow(), pcol() say chr(18)
return self


*-------------------------
static function zoom(cOp)
*-------------------------
if cOp="+" .and. print_preview.edit_p.fontsize <= 24
   print_preview.edit_p.fontsize:=  print_preview.edit_p.fontsize + 2
endif

if cOp="-" .and. print_preview.edit_p.fontsize > 7
   print_preview.edit_p.fontsize:=  print_preview.edit_p.fontsize - 2
endif
return nil


*-------------------------
Method searchstring(cTarget) CLASS TDOSPRINT
*-------------------------
print_preview.but_6.enabled:=.F.
print_preview.edit_p.caretpos:=1
::nOccur:=0
::cBusca:= cTarget
::cString := ""
::cString  := inputbox('Text: ','Search String')
if empty(::cString)
   return(NIL)
endif
print_preview.but_6.enabled:=.t.
::nextsearch()
return(nil)

*-----------------------------------------------------*
Method nextsearch( )
*-----------------------------------------------------*
local cString,ncaretpos
cString := UPPER(::cstring)
////ncount:=STRCOUNT( chr(13),cString, print_preview.edit_p.caretpos )
nCaretpos := ATplus(ALLTRIM(cString),UPPER(::cBusca),::nOccur)
::nOccur:=nCaretpos+1

print_preview.edit_p.setfocus
if nCaretpos>0
   print_preview.edit_p.caretpos:=nCaretPos
   print_preview.edit_p.refresh
else
   print_preview.but_6.enabled:=.F.
   msginfo("End search","Information")
endif
return nil

static function strcount(cbusca,cencuentra,n)
local nc:=0,i
cbusca:=alltrim(cbusca)
for i:=1 to n
    if upper(substr(cencuentra,i,len(cbusca)))=upper(cbusca)
       nc++
    endif
next i
return nc


*-------------------------
static function atplus(cbusca,ctodo,ninicio)
*-------------------------
local i,nposluna,lencbusca,lenctodo,uppercbusca
nposluna:=0
lenctodo:=len(ctodo)
lencbusca:=len(cbusca)
uppercbusca:=upper(cbusca)
for i:= ninicio to lenctodo
    if upper(substr(ctodo,i,lencbusca))=uppercbusca
       nposluna:=i
       exit
    endif
next i
return nposluna






//////////////////////////////////////Calc
CREATE CLASS TCALCPRINT FROM TPRINTBASE

    DATA oServiceManager INIT nil
    DATA oDesktop INIT nil
    DATA oDocument INIT nil
    DATA oSchedule INIT nil
    DATA oSheet INIT nil
    DATA oCell INIT nil
    DATA oColums INIT nil
    DATA oColumn INIT nil

*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex()
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimage() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printlinex() BLOCK {|| NIL }
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx() BLOCK {|| NIL }
*-------------------------

*-------------------------
method normaldosx() BLOCK {|| NIL }
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

ENDCLASS


*-------------------------
METHOD initx() CLASS TCALCPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="CALCPRINT"
return self

*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TCALCPRINT
*-------------------------
empty(lselect)
empty(lpreview)
empty(llandscape)
empty(npapersize)
empty(cprinterx)

::oServiceManager := TOleAuto():New("com.sun.star.ServiceManager")
::oDesktop := ::oServiceManager:createInstance("com.sun.star.frame.Desktop")
IF ::oDesktop = NIL
   MsgStop('OpenOffice Calc no encontrado','error')
   RETURN Nil
ENDIF
return self

*-------------------------
METHOD begindocx() CLASS TCALCPRINT
*-------------------------
::oDocument := ::oDesktop:loadComponentFromURL("private:factory/scalc","_blank", 0, {})
::oSchedule := ::oDocument:GetSheets()
::oSheet := ::oSchedule:GetByIndex(0)
*::oSheet:CharFontName := ::cfontname
*::oSheet:CharHeight   := ::nfontsize
return self


*-------------------------
METHOD enddocx() CLASS TCALCPRINT
*-------------------------
::oSheet:getColumns():setPropertyValue("OptimalWidth", .T.)
RETURN self


*-------------------------
METHOD releasex() CLASS TCALCPRINT
*-------------------------
   ::oServiceManager := nil
   ::oDesktop := nil
   ::oDocument := nil
   ::oSchedule := nil
   ::oSheet := nil
   ::oCell := nil
   ::oColums := nil
   ::oColumn := nil
RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TCALCPRINT
*-------------------------
return self


*-------------------------
METHOD endpagex() CLASS TCALCPRINT
*-------------------------
local MiLinea,nTEXECLPRINT1,MiCol2
ASORT(::alincelda,,, { |x, y| STR(x[1])+STR(x[2]) < STR(y[1])+STR(y[2]) })
IF ::nlinpag=0 .AND. LEN(::alincelda)>0
   ::cfontname:=::alincelda[1,6]
   ::nfontsize:=::alincelda[1,4]
   ::oSheet:CharFontName := ::cfontname
   ::oSheet:CharHeight   := ::nfontsize
ENDIF
MiLinea:=0
MiCol2:=1
FOR nTEXECLPRINT1=1 TO LEN(::alincelda)
   DO WHILE MiLinea<::alincelda[nTEXECLPRINT1,1]
      MiLinea++
      MiCol2:=1
   ENDDO
   ::oCell:=::oSheet:getCellByPosition(MiCol2-1,::nlinpag+MiLinea) // columna,linea
   IF VALTYPE(::alincelda[nTEXECLPRINT1,9])="N"
      ::oCell:SetValue(::alincelda[nTEXECLPRINT1,3])
   ELSE
      ::oCell:SetString(::alincelda[nTEXECLPRINT1,3])
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,6]<>::cfontname
      ::oCell:CharFontName:=::alincelda[nTEXECLPRINT1,6]
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,4]<>::nfontsize
      ::oCell:CharHeight  :=::alincelda[nTEXECLPRINT1,4]
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,7]=.T.
      ::oCell:CharWeight=150
   ENDIF
   do case
   case ::alincelda[nTEXECLPRINT1,5]="R"
      ::oCell:HoriJustify:=3
   case ::alincelda[nTEXECLPRINT1,5]="C"
      ::oCell:HoriJustify:=2
   endcase
   aMiColor:=::alincelda[nTEXECLPRINT1,8]
   IF aMiColor[3]<>0 .OR. aMiColor[2]<>0 .OR. aMiColor[1]<>0
      ::oCell:CharColor:=RGB(aMiColor[3],aMiColor[2],aMiColor[1])
   ENDIF
   MiCol2++
NEXT
::nlinpag:= ::nlinpag + MiLinea+1
::alincelda:={}
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS TCALCPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self


*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TCALCPRINT
*-------------------------
local alinceldax
empty(ncol)
empty(data)
empty(acolor)
empty(nlen)
if ::cunitslin>1
   nlin:=round(nlin/::cunitslin,0)
endif
IF ::cunits="MM"
   ctext:=ALLTRIM(ctext)
ENDIF
AADD(::alincelda,{nlin,ncol,ctext,nsize,calign,cfont,lbold,acolor,data})
return self






//////////////////////////////////////Excel
CREATE CLASS TEXCELPRINT FROM TPRINTBASE

    DATA oExcel INIT nil
    DATA oHoja  INIT nil

*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex()
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimagex()
*-------------------------

*-------------------------
METHOD printlinex() BLOCK {|| NIL }
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx() BLOCK {|| NIL }
*-------------------------

*-------------------------
method normaldosx() BLOCK {|| NIL }
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

ENDCLASS


*-------------------------
METHOD initx() CLASS TEXCELPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="EXCELPRINT"
return self

*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TEXCELPRINT
*-------------------------
empty(lselect)
empty(lpreview)
empty(llandscape)
empty(npapersize)
empty(cprinterx)

::oExcel := TOleAuto():New( "Excel.Application" )
IF Ole2TxtError() != 'S_OK'
   MsgStop('Excel not found','error')
   ::lprerror:=.T.
   ::exit:=.T.
   RETURN Nil
ENDIF
return self

*-------------------------
METHOD begindocx() CLASS TEXCELPRINT
*-------------------------
::oExcel:WorkBooks:Add()
::oHoja:=::oExcel:ActiveSheet()
::oHoja:Name := "List"
::oHoja:Cells:Font:Name := ::cfontname
::oHoja:Cells:Font:Size := ::nfontsize
return self


*-------------------------
METHOD enddocx() CLASS TEXCELPRINT
*-------------------------
local nCol
FOR nCol:=1 TO ::oHoja:UsedRange:Columns:Count()
    ::oHoja:Columns( nCol ):AutoFit()
NEXT
::oHoja:Cells( 1, 1 ):Select()
::oExcel:Visible := .T.
::oHoja:= NIL
::oExcel:= NIL
RETURN self


*-------------------------
METHOD releasex() CLASS TEXCELPRINT
*-------------------------
   ::oHoja := nil
   ::oExcel := nil
RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TEXCELPRINT
*-------------------------
return self


*-------------------------
METHOD endpagex() CLASS TEXCELPRINT
*-------------------------
local MiLinea,nTEXECLPRINT1,MiCol2
ASORT(::alincelda,,, { |x, y| STR(x[1])+STR(x[2]) < STR(y[1])+STR(y[2]) })
IF ::nlinpag=0 .AND. LEN(::alincelda)>0
   ::cfontname:=::alincelda[1,6]
   ::nfontsize:=::alincelda[1,4]
   ::oHoja:Cells:Font:Name := ::cfontname
   ::oHoja:Cells:Font:Size := ::nfontsize
ENDIF
MiLinea:=0
MiCol2:=1
FOR nTEXECLPRINT1=1 TO LEN(::alincelda)
   DO WHILE MiLinea<::alincelda[nTEXECLPRINT1,1]
      MiLinea++
      MiCol2:=1
   ENDDO
   ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):Value := ::alincelda[nTEXECLPRINT1,3]
   IF ::alincelda[nTEXECLPRINT1,6]<>::cfontname
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):Font:Name := ::alincelda[nTEXECLPRINT1,6]
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,4]<>::nfontsize
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):Font:Size := ::alincelda[nTEXECLPRINT1,4]
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,7]=.T.
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):Font:Bold := ::alincelda[nTEXECLPRINT1,7]
   ENDIF
   IF ::alincelda[nTEXECLPRINT1,9]=.T.
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):Font:Italic := ::alincelda[nTEXECLPRINT1,9]
   ENDIF
   do case
   case ::alincelda[nTEXECLPRINT1,5]="R"
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):HorizontalAlignment:= -4152  //Derecha
   case ::alincelda[nTEXECLPRINT1,5]="C"
      ::oHoja:Cells(::nlinpag+MiLinea,MiCol2):HorizontalAlignment:= -4108  //Centrar
   endcase
   MiCol2++
NEXT
::nlinpag:= ::nlinpag + MiLinea+1
::alincelda:={}
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS TEXCELPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self


*-------------------------
METHOD printimagex(nlin,ncol,nlinf,ncolf,cimage) CLASS TEXCELPRINT
*-------------------------
local cfolder :=  getcurrentfolder()+"\"
local ccol :=alltrim(str(ncol))
local crango := "A"+ccol+":"+"A"+ccol
empty(nlin)
empty(nlinf)
empty(ncolf)

::oHoja:range( crango ):Select()
IF AT("\",cimage)=0
	cimage:=cfolder+cimage
ENDIF
::oHoja:Pictures:Insert(cimage)

RETURN self

*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TEXCELPRINT
*-------------------------
local alinceldax
empty(ncol)
empty(data)
empty(acolor)
empty(nlen)
if ::cunitslin>1
   nlin:=round(nlin/::cunitslin,0)
endif
IF ::cunits="MM"
   ctext:=ALLTRIM(ctext)
ENDIF
AADD(::alincelda,{nlin,ncol,ctext,nsize,calign,cfont,lbold,acolor,litalic})
return self







//////////////////////////////////////RTF
CREATE CLASS TRTFPRINT FROM TPRINTBASE

*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimage() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx()
*-------------------------

*-------------------------
method normaldosx()
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

ENDCLASS


*-------------------------
METHOD initx() CLASS TRTFPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="RTFPRINT"
return self

*-------------------------
METHOD begindocx(cdoc) CLASS TRTFPRINT
*-------------------------
local   MARGENSUP:=LTRIM(STR(ROUND(15*56.7,0)))
local   MARGENINF:=LTRIM(STR(ROUND(15*56.7,0)))
local   MARGENIZQ:=LTRIM(STR(ROUND(10*56.7,0)))
local   MARGENDER:=LTRIM(STR(ROUND(10*56.7,0)))

////Empty( cdoc )

Default cDoc to "List"

AADD(oPrintRTF1,"{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang3082\deflangfe3082{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}")
AADD(oPrintRTF1,"{\f106\froman\fcharset238\fprq2 Times New Roman CE;}{\f107\froman\fcharset204\fprq2 Times New Roman Cyr;}{\f109\froman\fcharset161\fprq2 Times New Roman Greek;}{\f110\froman\fcharset162\fprq2 Times New Roman Tur;}")
AADD(oPrintRTF1,"{\f111\froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\f112\froman\fcharset178\fprq2 Times New Roman (Arabic);}{\f113\froman\fcharset186\fprq2 Times New Roman Baltic;}{\f122\fmodern\fcharset238\fprq1 Courier New CE;}")
AADD(oPrintRTF1,"{\f123\fmodern\fcharset204\fprq1 Courier New Cyr;}{\f125\fmodern\fcharset161\fprq1 Courier New Greek;}{\f126\fmodern\fcharset162\fprq1 Courier New Tur;}{\f127\fmodern\fcharset177\fprq1 Courier New (Hebrew);}")
AADD(oPrintRTF1,"{\f128\fmodern\fcharset178\fprq1 Courier New (Arabic);}{\f129\fmodern\fcharset186\fprq1 Courier New Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;")
AADD(oPrintRTF1,"\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{")
AADD(oPrintRTF1,"\ql \li0\ri0\widctlpar\faauto\adjustright\rin0\lin0\itap0 \fs20\lang3082\langfe3082\cgrid\langnp3082\langfenp3082 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}}")
AADD(oPrintRTF1,"{\info{\title "+cDoc+"}{\creatim\yr"+STRZERO(YEAR(DATE()),4)+"\mo"+STRZERO(MONTH(DATE()),2)+"\dy"+STRZERO(DAY(DATE()),2)+"\hr"+LEFT(TIME(),2)+"\min"+SUBSTR(TIME(),4,2)+"}{\version10}{\edmins16}{\nofpages1}{\nofwords167}{\nofchars954}{\nofcharsws0}{\vern8249}}")
**AADD(oPrintRTF1,"\paperw11907\paperh16840\margl284\margr284\margt1134\margb1134 \widowctrl\ftnbj\aenddoc\hyphhotz425\noxlattoyen\expshrtn\noultrlspc\dntblnsbdb\nospaceforul\hyphcaps0\horzdoc\dghspace120\dgvspace120\dghorigin1701\dgvorigin1984\dghshow0\dgvshow3")
AADD(oPrintRTF1,IF(oPrintRTF3=.T.,"\landscape\paperw16840\paperh11907","\paperw11907\paperh16840")+ ;
   "\margl"+MARGENIZQ+"\margr"+MARGENDER+"\margt"+MARGENSUP+"\margb"+MARGENINF+ ;
   " \widowctrl\ftnbj\aenddoc\hyphhotz425\noxlattoyen\expshrtn\noultrlspc\dntblnsbdb\nospaceforul\hyphcaps0\horzdoc\dghspace120\dgvspace120\dghorigin1701\dgvorigin1984\dghshow0\dgvshow3")
AADD(oPrintRTF1,"\jcompress\viewkind1\viewscale80\nolnhtadjtbl \fet0\sectd \psz9\linex0\headery851\footery851\colsx709\sectdefaultcl {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3")
AADD(oPrintRTF1,"\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}")
AADD(oPrintRTF1,"{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain ")
AADD(oPrintRTF1,"\qj \li0\ri0\nowidctlpar\faauto\adjustright\rin0\lin0\itap0 \fs20\lang3082\langfe3082\cgrid\langnp3082\langfenp3082")
IF ::cunits="MM"
   AADD(oPrintRTF1,"{\b\f0\lang1034\langfe3082\cgrid0\langnp1034")
ELSE
   AADD(oPrintRTF1,"{\b\f2\lang1034\langfe3082\cgrid0\langnp1034")
ENDIF

return self


*-------------------------
METHOD enddocx() CLASS TRTFPRINT
*-------------------------
local nTRTFPRINT
private rutaficrtf1
IF RIGHT(oPrintRTF1[LEN(oPrintRTF1)],6)=" \page"
   oPrintRTF1[LEN(oPrintRTF1)]:=LEFT(oPrintRTF1[LEN(oPrintRTF1)] , LEN(oPrintRTF1[LEN(oPrintRTF1)])-6 )
ENDIF
AADD(oPrintRTF1,"\par }}")
RUTAFICRTF1:=GetCurrentFolder()
SET PRINTER TO &RUTAFICRTF1\Print.rtf
SET DEVICE TO PRINTER
SETPRC(0,0)
FOR nTRTFPRINT=1 TO LEN(oPrintRTF1)
   @ PROW(),PCOL() SAY oPrintRTF1[nTRTFPRINT]
   @ PROW()+1,0 SAY ""
NEXT
SET DEVICE TO SCREEN
SET PRINTER TO
RELEASE oPrintRTF1,oPrintRTF2,oPrintRTF3
IF ShellExecute(0, "open", "soffice.exe", "Print.rtf" , RUTAFICRTF1 , 1)<=32
   IF ShellExecute(0, "open", "WinWord.exe", "Print.rtf" , RUTAFICRTF1 , 1)<=32
      IF ShellExecute(0, "open", "Print.rtf" , RUTAFICRTF1 , , 1)<=32
         MSGINFO("No se ha localizado el programa asociado a la extension RTF"+HB_OsNewLine()+HB_OsNewLine()+ ;
                 "El fichero se ha guardado en:"+HB_OsNewLine()+RUTAFICRTF1+"\Print.rtf")
      ENDIF
   ENDIF
ENDIF
RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TRTFPRINT
*-------------------------
return self

*-------------------------
METHOD endpagex() CLASS TRTFPRINT
*-------------------------
local milinea,nTRTFPRINT1,nTRTFPRINT2
ASORT(::alincelda,,, { |x, y| STR(x[1])+STR(x[2]) < STR(y[1])+STR(y[2]) })
MiLinea:=0
FOR nTRTFPRINT1=1 TO LEN(::alincelda)
   IF MiLinea<::alincelda[nTRTFPRINT1,1]
      DO WHILE MiLinea<::alincelda[nTRTFPRINT1,1]
         AADD(oPrintRTF1,"\par ")
         MiLinea++
      ENDDO
      IF ::cunits="MM"
         oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}\pard \qj \li0\ri0\nowidctlpar"
         FOR nTRTFPRINT2=1 TO LEN(::alincelda)
            IF ::alincelda[nTRTFPRINT2,1]=::alincelda[nTRTFPRINT1,1]
               do case
               case ::alincelda[nTRTFPRINT2,5]="R"
                  oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"\tqr"+"\tx"+LTRIM(STR(ROUND((::alincelda[nTRTFPRINT2,2]-10)*56.7,0)))
               case ::alincelda[nTRTFPRINT2,5]="C"
                  oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"\tqc"+"\tx"+LTRIM(STR(ROUND((::alincelda[nTRTFPRINT2,2]-10)*56.7,0)))
               otherwise
                  oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"\tql"+"\tx"+LTRIM(STR(ROUND((::alincelda[nTRTFPRINT2,2]-10)*56.7,0)))
               endcase
            ENDIF
         NEXT
         oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"\faauto\adjustright\rin0\lin0\itap0 {"
      ENDIF
   ENDIF
   IF oPrintRTF2<>::alincelda[nTRTFPRINT1,4]
      oPrintRTF2:=::alincelda[nTRTFPRINT1,4]
      DO CASE
      CASE oPrintRTF2<=8
         IF ::cunits="MM"
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f0\fs16\lang1034\langfe3082\cgrid0\langnp1034"
         ELSE
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f2\fs16\lang1034\langfe3082\cgrid0\langnp1034"
         ENDIF
      CASE oPrintRTF2>=9 .AND. oPrintRTF2<=10
         IF ::cunits="MM"
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f0\lang1034\langfe3082\cgrid0\langnp1034"
         ELSE
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f2\lang1034\langfe3082\cgrid0\langnp1034"
         ENDIF
      CASE oPrintRTF2>=11 .AND. oPrintRTF2<=12
         IF ::cunits="MM"
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f0\fs24\lang1034\langfe3082\cgrid\langnp1034"
         ELSE
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f2\fs24\lang1034\langfe3082\cgrid\langnp1034"
         ENDIF
      CASE oPrintRTF2>=13
         IF ::cunits="MM"
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f0\fs28\lang1034\langfe3082\cgrid\langnp1034"
         ELSE
            oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"}{\b\f2\fs28\lang1034\langfe3082\cgrid\langnp1034"
         ENDIF
      ENDCASE
   ENDIF

   IF ::cunits="MM"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"\tab "
   ENDIF

   oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+::alincelda[nTRTFPRINT1,3]
NEXT
oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+" \page"
::alincelda:={}
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS TRTFPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self


*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TRTFPRINT
*-------------------------
Empty( data )
Empty( cfont )
Empty( lbold )
Empty( acolor )
Empty( nlen )
if ::cunitslin>1
   nlin:=round(nlin/::cunitslin,0)
endif
IF ::cunits="MM"
   ctext:=ALLTRIM(ctext)
ENDIF
AADD(::alincelda,{nlin,ncol,ctext,nsize,calign,cfont,lbold,acolor})
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TRTFPRINT
*-------------------------
Empty( nlin )
Empty( ncol )
Empty( nlinf )
Empty( ncolf )
Empty( atcolor )
Empty( ntwpen )
return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TRTFPRINT
*-------------------------
PUBLIC oPrintRTF1,oPrintRTF2,oPrintRTF3
Empty( lselect )
Empty( lpreview )
Empty( npapersize )
Empty( cprinterx )
oPrintRTF1:={}
oPrintRTF2:=10
oPrintRTF3:=llandscape
return self


*-------------------------
METHOD condendosx() CLASS TRTFPRINT
*-------------------------
return self


*-------------------------
METHOD normaldosx() CLASS TRTFPRINT
*-------------------------
return self






//////////////////////////////////////CSV
CREATE CLASS TCSVPRINT FROM TPRINTBASE


*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimage() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx()
*-------------------------

*-------------------------
method normaldosx()
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

ENDCLASS


*-------------------------
METHOD initx() CLASS TCSVPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="CSVPRINT"
return self


*-------------------------
METHOD begindocx(cdoc) CLASS TCSVPRINT
*-------------------------
Default cDoc to "List"
return self


*-------------------------
METHOD enddocx() CLASS TCSVPRINT
*-------------------------
RUTAFICRTF1:=GetCurrentFolder()
SET PRINTER TO &RUTAFICRTF1\Print.csv
SET DEVICE TO PRINTER
SETPRC(0,0)
FOR nTCSVPRINT=1 TO LEN(oPrintCSV1)
   @ PROW(),PCOL() SAY oPrintCSV1[nTCSVPRINT]
   @ PROW()+1,0 SAY ""
NEXT
SET DEVICE TO SCREEN
SET PRINTER TO
RELEASE oPrintCSV1,oPrintCSV2,oPrintCSV3
IF ShellExecute(0, "open", "soffice.exe", "Print.csv" , RUTAFICRTF1 , 1)<=32
   IF ShellExecute(0, "open", "Excel.exe", "Print.csv" , RUTAFICRTF1 , 1)<=32
      IF ShellExecute(0, "open", "Print.csv" , RUTAFICRTF1 , , 1)<=32
         MSGINFO("No se ha localizado el programa asociado a la extension CSV"+HB_OsNewLine()+HB_OsNewLine()+ ;
                 "El fichero se ha guardado en:"+HB_OsNewLine()+RUTAFICRTF1+"\Print.csv")
      ENDIF
   ENDIF
ENDIF
RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS TCSVPRINT
*-------------------------
return self


*-------------------------
METHOD endpagex() CLASS TCSVPRINT
*-------------------------
ASORT(::alincelda,,, { |x, y| STR(x[1])+STR(x[2]) < STR(y[1])+STR(y[2]) })
MiLinea1:=0
MiLinea2:=0
FOR nTCSVPRINT1=1 TO LEN(::alincelda)
   IF MiLinea1<::alincelda[nTCSVPRINT1,1]
      DO WHILE MiLinea1<::alincelda[nTCSVPRINT1,1]
         AADD(oPrintCSV1,"")
         MiLinea1++
      ENDDO
   ENDIF

   IF LEN(oPrintCSV1)>0
      IF LEN(oPrintCSV1[LEN(oPrintCSV1)])=0
         oPrintCSV1[LEN(oPrintCSV1)]:=::alincelda[nTCSVPRINT1,3]
      ELSE
         oPrintCSV1[LEN(oPrintCSV1)]:=oPrintCSV1[LEN(oPrintCSV1)]+";"+STRTRAN(::alincelda[nTCSVPRINT1,3],";",",")
      ENDIF
   ENDIF
NEXT
::alincelda:={}
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS TCSVPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self


*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TCSVPRINT
*-------------------------
Empty( data )
Empty( cfont )
Empty( lbold )
Empty( acolor )
Empty( nlen )
if ::cunitslin>1
   nlin:=round(nlin/::cunitslin,0)
endif
IF ::cunits="MM"
   ctext:=ALLTRIM(ctext)
ENDIF
AADD(::alincelda,{nlin,ncol,ctext,nsize,calign,cfont,lbold,acolor})
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TCSVPRINT
*-------------------------
Empty( nlin )
Empty( ncol )
Empty( nlinf )
Empty( ncolf )
Empty( atcolor )
Empty( ntwpen )
return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TCSVPRINT
*-------------------------
PUBLIC oPrintCSV1,oPrintCSV2,oPrintCSV3
Empty( lselect )
Empty( lpreview )
Empty( npapersize )
Empty( cprinterx )
oPrintCSV1:={}
oPrintCSV2:=10
oPrintCSV3:=llandscape
return self


*-------------------------
METHOD condendosx() CLASS TCSVPRINT
*-------------------------
return self


*-------------------------
METHOD normaldosx() CLASS TCSVPRINT
*-------------------------
return self






//////////////////////////////////////HTML
CREATE CLASS THTMLPRINT FROM TPRINTBASE


*-------------------------
METHOD initx()
*-------------------------

*-------------------------
METHOD begindocx()
*-------------------------

*-------------------------
METHOD enddocx()
*-------------------------

*-------------------------
METHOD beginpagex()
*-------------------------

*-------------------------
METHOD endpagex()
*-------------------------

*-------------------------
METHOD releasex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printdatax()
*-------------------------

*-------------------------
METHOD printimage() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printlinex()
*-------------------------

*-------------------------
METHOD printrectanglex BLOCK { || nil }
*-------------------------

*-------------------------
METHOD selprinterx()
*-------------------------

*-------------------------
METHOD getdefprinterx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setcolorx() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD setpreviewsizex() BLOCK { || nil }
*-------------------------

*-------------------------
METHOD printroundrectanglex() BLOCK { || nil }
*-------------------------

*-------------------------
method condendosx()
*-------------------------

*-------------------------
method normaldosx()
*-------------------------

*-------------------------
METHOD setunits()   // mm o rowcol , mm por renglon
*-------------------------

ENDCLASS


*-------------------------
METHOD initx() CLASS THTMLPRINT
*-------------------------
::impreview:=.F.
::cprintlibrary:="HTMLPRINT"
return self


*-------------------------
METHOD begindocx(cdoc) CLASS THTMLPRINT
*-------------------------
////Empty( cdoc )
Default cDoc to "List"

AADD(oPrintRTF1,"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN'")
AADD(oPrintRTF1,"'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>")
AADD(oPrintRTF1,"<html xmlns='http://www.w3.org/1999/xhtml'>")
AADD(oPrintRTF1,"<head>")
AADD(oPrintRTF1,"<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />")
AADD(oPrintRTF1,"<title>"+cdoc+"</title>")
AADD(oPrintRTF1,"</head>")
AADD(oPrintRTF1,"<body>")
return self


*-------------------------
METHOD enddocx() CLASS THTMLPRINT
*-------------------------
local nTHTMLPRINT
private rutaficrtf1
RUTAFICRTF1:=GetCurrentFolder()
SET PRINTER TO &RUTAFICRTF1\Print.html
SET DEVICE TO PRINTER
SETPRC(0,0)
FOR nTHTMLPRINT=1 TO LEN(oPrintRTF1)
   @ PROW(),PCOL() SAY oPrintRTF1[nTHTMLPRINT]
   @ PROW()+1,0 SAY ""
NEXT
   @ PROW()+1,0 SAY "</body>"
   @ PROW()+1,0 SAY "</html>"
   @ PROW()+1,0 SAY ""
SET DEVICE TO SCREEN
SET PRINTER TO
RELEASE oPrintRTF1,oPrintRTF2,oPrintRTF3
IF ShellExecute(0, "open", "firefox.exe", "Print.html" , RUTAFICRTF1 , 1)<=32
   IF ShellExecute(0, "open", "Print.html" , RUTAFICRTF1 , , 1)<=32
      MSGINFO("No se ha localizado el programa asociado a la extension HTML"+HB_OsNewLine()+HB_OsNewLine()+ ;
              "El fichero se ha guardado en:"+HB_OsNewLine()+RUTAFICRTF1+"\Print.html")
   ENDIF
ENDIF
RETURN self


*-------------------------
METHOD beginpagex(nWidth,nHeight) CLASS THTMLPRINT
*-------------------------
return self


*-------------------------
METHOD endpagex() CLASS THTMLPRINT
*-------------------------
local milinea,nTHTMLPRINT1,nTHTMLPRINT2,aMAXCOL2:={},nMAXCOL2:=0
ASORT(::alincelda,,, { |x, y| STR(x[1])+STR(x[2]) < STR(y[1])+STR(y[2]) })

FOR nTHTMLPRINT1=1 TO LEN(::alincelda)
   SICOL:=ASCAN(aMAXCOL2,{ |AVAL| AVAL[1]==::alincelda[nTHTMLPRINT1,1] } )
   IF SICOL=0
      Aadd( aMAXCOL2 , {::alincelda[nTHTMLPRINT1,1],1} )
   ELSE
      aMAXCOL2[SICOL,2]:=aMAXCOL2[SICOL,2]+1
   ENDIF
NEXT

FOR nTHTMLPRINT1=1 TO LEN(aMAXCOL2)
   nMAXCOL2:=MAX(nMAXCOL2,aMAXCOL2[nTHTMLPRINT1,2])
NEXT
AADD(oPrintRTF1,"")
AADD(oPrintRTF1,"<br />")
AADD(oPrintRTF1,"")
AADD(oPrintRTF1,"<table align='center' border='0'>")

MiLinea:=0
FOR nTHTMLPRINT1=1 TO LEN(::alincelda)
   IF MiLinea<::alincelda[nTHTMLPRINT1,1]
      IF MiLinea>0
         AADD(oPrintRTF1,"</tr>")
      ENDIF
      DO WHILE MiLinea<::alincelda[nTHTMLPRINT1,1]
         AADD(oPrintRTF1,"<tr><td></td></tr>")
         MiLinea++
      ENDDO
      AADD(oPrintRTF1,"<tr>")
   ENDIF
   AADD(oPrintRTF1,"<td>")
   AADD(oPrintRTF1,"")

   DO CASE
   CASE ::alincelda[nTHTMLPRINT1,5]="C"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"<div align='center'>"
   CASE ::alincelda[nTHTMLPRINT1,5]="R"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"<div align='right'>"
   ENDCASE
   IF ::alincelda[nTHTMLPRINT1,7]=.T.
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"<strong>"
   ENDIF
   IF UPPER(::alincelda[nTHTMLPRINT1,6])="COURIER"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"<font face='Courier New, Courier, monospace'>"
   ENDIF

   oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+::alincelda[nTHTMLPRINT1,3]

   IF UPPER(::alincelda[nTHTMLPRINT1,6])="COURIER"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"</font>"
   ENDIF
   IF ::alincelda[nTHTMLPRINT1,7]=.T.
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"</strong>"
   ENDIF
   IF ::alincelda[nTHTMLPRINT1,5]="C" .OR. ::alincelda[nTHTMLPRINT1,5]="R"
      oPrintRTF1[LEN(oPrintRTF1)]:=oPrintRTF1[LEN(oPrintRTF1)]+"</div>"
   ENDIF

   AADD(oPrintRTF1,"</td>")
NEXT
AADD(oPrintRTF1,"</tr>")
AADD(oPrintRTF1,"")
AADD(oPrintRTF1,"</table>")
AADD(oPrintRTF1,"")
::alincelda:={}
return self


*-------------------------
METHOD setunits(cunitsx,cunitslinx) CLASS THTMLPRINT
*-------------------------
if cunitsx="MM"
   ::cunits:="MM"
else
   ::cunits:="ROWCOL"
endif
if cunitslinx=NIL
   ::cunitslin:=1
else
   ::cunitslin:=cunitslinx
endif
RETURN self


*-------------------------
METHOD printdatax(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS THTMLPRINT
*-------------------------
Empty( data )
Empty( cfont )
Empty( lbold )
Empty( acolor )
Empty( nlen )
if ::cunitslin>1
   nlin:=round(nlin/::cunitslin,0)
endif
IF ::cunits="MM"
   ctext:=ALLTRIM(ctext)
ENDIF
AADD(::alincelda,{nlin,ncol,ctext,nsize,calign,cfont,lbold,acolor})
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS THTMLPRINT
*-------------------------
Empty( nlin )
Empty( ncol )
Empty( nlinf )
Empty( ncolf )
Empty( atcolor )
Empty( ntwpen )
return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS THTMLPRINT
*-------------------------
PUBLIC oPrintRTF1,oPrintRTF2,oPrintRTF3
Empty( lselect )
Empty( lpreview )
Empty( npapersize )
Empty( cprinterx )
oPrintRTF1:={}
oPrintRTF2:=10
oPrintRTF3:=llandscape
return self


*-------------------------
METHOD condendosx() CLASS THTMLPRINT
*-------------------------
return self


*-------------------------
METHOD normaldosx() CLASS THTMLPRINT
*-------------------------
return self






//////////////////////////////////////PDF
CREATE CLASS TPDFPRINT FROM TPRINTBASE

/////PRIVATE:
DATA oPDF        as object                     //el objeto pdf
DATA oPage       as object                     //la pagina pdf
DATA cDocument   as character init ""          //nombre del documento
DATA cPageSize   as character init ""          //tamaño de pagina.
DATA cPageOrient as numeric   init 0           //0 = portrait, 1 = Landscape
DATA nFontType   as numeric   init 1           //tipo de fuente(normal=0 o negrita=1)
DATA lPreview    as logical   init .f.         //indica si abrimos el pdf al finalizar
DATA aPaper      as array     init {} hidden   //array con los tipos de papel soportados por la clase pdf.


METHOD initx()
METHOD begindocx()
METHOD enddocx()
METHOD beginpagex()
METHOD endpagex()
METHOD releasex()
METHOD printdatax()
METHOD printimagex()
METHOD printlinex()
METHOD printrectanglex()
METHOD selprinterx()
METHOD getdefprinterx()
METHOD setcolorx()
METHOD setpreviewsizex()
METHOD printroundrectanglex()

ENDCLASS
*---------------------------------------

*-------------------------
METHOD INITx() CLASS TPDFPRINT
*-------------------------

//              nombre    ,pdf,tprint
aadd(::aPaper,{"LETTER"   , 0 , 1})
aadd(::aPaper,{"LEGAL"    , 1 , 5})
aadd(::aPaper,{"A3"       , 2 , 8})
aadd(::aPaper,{"A4"       , 3 , 9})
aadd(::aPaper,{"A5"       , 4 ,11})
aadd(::aPaper,{"B4"       , 5 ,12})
aadd(::aPaper,{"B5"       , 6 ,13})
aadd(::aPaper,{"EXECUTIVE", 7 , 7})
aadd(::aPaper,{"US4x6"    , 8 , 9})
aadd(::aPaper,{"US4x8"    , 9 , 9})
aadd(::aPaper,{"US5x7"    , 10, 9})
aadd(::aPaper,{"COMM10"   , 11, 9})
aadd(::aPaper,{"EOF"      , 12, 9})

::cprintlibrary:="PDFPRINT"

return self


*-------------------------
METHOD selprinterx( lselect , lpreview, llandscape , npapersize ,cprinterx,nres,nWidth,nHeight) CLASS TPDFPRINT
*-------------------------
local nPos := 0

Default lpreview to .f.
Default llandscape to .f.
Default npapersize to 3

empty( lselect )

/*lSelect no lo tomamos en cuenta aqui*/

***Si se va a necesitar abrir el pdf al finalizar la generacion***
::lPreview := lpreview
***Establecemos la orientacion de la hoja***
::cPageOrient := IIF(lLandScape,1,0)

***Establecemos el tamaño del papel***
IF VALTYPE(npapersize)="N"
	nPos := aScan(::aPaper,{|x|x[3] = npapersize})
	If nPos > 0
	   ::cPageSize := ::aPaper[nPos][1]
	Else
	   ::cPageSize := "A4"
	Endif
ENDIF

return self


*-------------------------
METHOD BEGINDOCx (cdoc) CLASS TPDFPRINT
*-------------------------
public cpdfname:= cdoc+".pdf"
Public cpos :={}

   ::oPdf := HPDF_New()
   if ::oPdf == NIL
      MsgAlert( 'Pdf no ha sido creado', 'error' )
      return nil
   endif
   If file ( cpdfname )
      Ferase( cpdfname )
   Endif

**	HPDF_SetPageMode(::oPdf,1) //tipo de vista miniaturas
**	HPDF_SetPageLayout(::oPdf,2) //paginas por vista
	HPDF_SetInfoAttr(::oPdf,5,cdoc)

   /* set compression mode */
**   HPDF_SetCompressionMode( ::oPdf, 0x0F ) //HPDF_COMP_ALL


return self


*-------------------------
METHOD ENDDOCx() CLASS TPDFPRINT
*-------------------------
   HPDF_SaveToFile( ::oPdf, cpdfname )
**	MsgBox(HPDF_GETERROR(),"error")
   HPDF_Free( ::oPdf )
IF ::lPreview=.T.
	IF ShellExecute(0, "open", cpdfname, GetCurrentFolder(), ,1) <=32
	     MSGINFO("No hay programa asociado a PDF"+HB_OsNewLine()+HB_OsNewLine()+ ;
	     "Fichero grabado en:"+HB_OsNewLine()+cpdfname)
	ENDIF
ENDIF
return self


*-------------------------
METHOD BEGINPAGEx(nWidth,nHeight) CLASS TPDFPRINT
*-------------------------
   ::oPage := HPDF_AddPage(::oPdf)
   aadd(cpos ,::oPage)
	IF nWidth=0 .AND. nHeight=0
	   HPDF_Page_SetSize( ::cPageSize , ::cPageOrient )
	ELSE
		HPDF_Page_SetWidth( ::oPage , nWidth*2.83 )
		HPDF_Page_SetHeight( ::oPage , nHeight*2.83 )
	ENDIF
   Public PDFheight := HPDF_Page_GetHeight(::oPage)
   Public PDFwidth  := HPDF_Page_GetWidth(::oPage)

return self


*-------------------------
METHOD ENDPAGEx() CLASS TPDFPRINT
*-------------------------
return self


*-------------------------
METHOD RELEASEx() CLASS TPDFPRINT
*-------------------------
return self


*-------------------------
METHOD PRINTDATAx(nlin,ncol,data,cfont,nsize,lbold,acolor,calign,nlen,ctext,nangle,litalic) CLASS TPDFPRINT
*-------------------------
local nType   := 0
local nlength := 0
local cColor  := Chr(253)
local I

Default cFont to ::cFontName
Default nSize to ::nFontSize
////Default lBold to .f.
default aColor to ::acolor

DO CASE
CASE AT("HELVETICA",UPPER(cfont))<>0
	DO CASE
	CASE lbold=.T. .OR. litalic=.F.
		cfont:="Helvetica-Bold"
	CASE lbold=.F. .OR. litalic=.T.
		cfont:="Helvetica-Oblique"
	CASE lbold=.T. .OR. litalic=.T.
		cfont:="Helvetica-BoldOblique"
	OTHERWISE
		cfont:="Helvetica"
	ENDCASE
CASE AT("COURIER",UPPER(cfont))<>0
	DO CASE
	CASE lbold=.T. .OR. litalic=.F.
		cfont:="Courier-Bold"
	CASE lbold=.F. .OR. litalic=.T.
		cfont:="Courier-Oblique"
	CASE lbold=.T. .OR. litalic=.T.
		cfont:="Courier-BoldOblique"
	OTHERWISE
		cfont:="Courier"
	ENDCASE
CASE AT("SYMBOL",UPPER(cfont))<>0
	cfont:="Symbol"
OTHERWISE //AT("TIMES",UPPER(cfont))<>0
	DO CASE
	CASE lbold=.T. .OR. litalic=.F.
		cfont:="Times-Bold"
	CASE lbold=.F. .OR. litalic=.T.
		cfont:="Times-Italic"
	CASE lbold=.T. .OR. litalic=.T.
		cfont:="Times-BoldItalic"
	OTHERWISE
		cfont:="Times-Roman"
	ENDCASE
ENDCASE

 	HPDF_Page_SetRGBFill( ::oPage, acolor[1]/255, acolor[2]/255, acolor[3]/255)
   def_font := HPDF_GetFont( ::oPdf, cfont, "CP1252" )
   HPDF_Page_SetFontAndSize( ::oPage, def_font, nsize )
   HPDF_Page_BeginText( ::oPage )
	IF nangle<>0
		rad1:=nangle / 180 * 3.141592
		HPDF_Page_SetTextMatrix (::oPage, cos(rad1), sin(rad1), -sin(rad1), cos(rad1), ncol*2.83, nlin*2.83)
	ENDIF

	nlin:=nlin*2.83
	ncol:=ncol*2.83
	DO CASE
	CASE calign="C"
	   tw := HPDF_Page_TextWidth( ::oPage, ctext )
		ncol:=ncol-(tw/2)
	CASE calign="R"
	   tw := HPDF_Page_TextWidth( ::oPage, ctext )
		ncol:=ncol-(tw)
	ENDCASE

	**cambiar esquina inicial**
	nlin:=PDFheight-nlin-(nsize/1.4)

   HPDF_Page_TextOut( ::oPage, ncol, nlin, ctext)
   HPDF_Page_EndText( ::oPage )

return self


*-------------------------
METHOD printimagex(nlin,ncol,nlinf,ncolf,cimage) CLASS TPDFPRINT
*-------------------------
IF FILE(cimage)
	DO CASE
	CASE UPPER(RIGHT(cimage,4))=".JPG" .OR. ;
	     UPPER(RIGHT(cimage,5))=".JPEG"
	   himage = HPDF_LoadJpegImageFromFile(::oPdf, cimage)
	CASE UPPER(RIGHT(cimage,4))=".PNG"
		himage = HPDF_LoadPngImageFromFile (::oPdf, cimage)
	OTHERWISE
		return self
	ENDCASE

	**cambiar esquina inicial**
	nlin:=(PDFheight/2.83)-nlin
	nlinf:=(PDFheight/2.83)-nlinf

   nlinf := nlinf - nlin
   ncolf := nColf - nCol

	HPDF_Page_DrawImage (::oPage, himage, ncol*2.83, nlin*2.83, ncolf*2.83, nlinf*2.83)
ENDIF
return self


*-------------------------
METHOD printlinex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TPDFPRINT
*-------------------------
default atColor to ::acolor
	**cambiar esquina inicial**
	nlin:=(PDFheight/2.83)-nlin
	nlinf:=(PDFheight/2.83)-nlinf

	HPDF_Page_SetRGBStroke( ::oPage, atcolor[1]/255, atcolor[2]/255, atcolor[3]/255)
   HPDF_Page_SetLineWidth(::oPage, ntwpen)
   HPDF_Page_MoveTo(::oPage, ncol*2.83, nlin*2.83)
   HPDF_Page_LineTo(::oPage, ncolf*2.83, nlinf*2.83)
   HPDF_Page_Stroke(::oPage )
return self

*-------------------------
METHOD printrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen,arcolor ) CLASS TPDFPRINT
*-------------------------
Default ntwpen to ::nwpen
default atColor to ::acolor
	**cambiar esquina inicial**
	nlin:=(PDFheight/2.83)-nlin
	nlinf:=(PDFheight/2.83)-nlinf

   nlinf := nlinf - nlin
   ncolf := nColf - nCol
	HPDF_Page_SetRGBStroke( ::oPage, atcolor[1]/255, atcolor[2]/255, atcolor[3]/255)
	HPDF_Page_SetRGBFill( ::oPage, arcolor[1]/255, arcolor[2]/255, arcolor[3]/255)
   HPDF_Page_SetLineWidth( ::oPage, ntwpen )
   HPDF_Page_Rectangle( ::oPage, ncol*2.83, nlin*2.83, ncolf*2.83, nlinf*2.83 )
   HPDF_Page_Fill(::oPage )
   HPDF_Page_Rectangle( ::oPage, ncol*2.83, nlin*2.83, ncolf*2.83, nlinf*2.83 )
   HPDF_Page_Stroke(::oPage )
return self


*-------------------------
METHOD printroundrectanglex(nlin,ncol,nlinf,ncolf,atcolor,ntwpen ) CLASS TPDFPRINT
*-------------------------
return self


*-------------------------
METHOD getdefprinterx() CLASS TPDFPRINT
*-------------------------
local cdefprinter
return cdefprinter


*-------------------------
METHOD setcolorx() CLASS TPDFPRINT
*-------------------------
return self


*-------------------------
METHOD setpreviewsizex( /* ntam */) CLASS TPDFPRINT
*-------------------------
return self





/*
******************************************
FUNCTION MiColorHex(aNomColor,TipColor)
DEFAULT TipColor to ""
DEFAULT aNomColor to {}
IF LEN(aNomColor)<>3
	aNomColor:={255,255,255}
ENDIF
IF TipColor="PDF"
   Numcolor:=chr(aNomColor[1]) + chr(aNomColor[2]) + chr(aNomColor[3])
ELSE
	Numcolor:="#"+PADL(DECTOHEXA(aNomColor[1]),2,"0")+PADL(DECTOHEXA(aNomColor[2]),2,"0")+PADL(DECTOHEXA(aNomColor[3]),2,"0")
ENDIF
RETURN(Numcolor)
*/
