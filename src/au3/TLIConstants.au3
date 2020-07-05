#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.1
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Type information library constants and enumerations.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
Global Const $UUID_ITypeInfo    = "{00020401-0000-0000-C000-000000000046}"
Global Const $UUID_ITypeLib     = "{00020402-0000-0000-C000-000000000046}" 

Global Const $PARAMFLAG_NONE    = 0
Global Const $PARAMFLAG_FIN     = 0x1
Global Const $PARAMFLAG_FOUT    = 0x2
Global Const $PARAMFLAG_FLCID   = 0x4
Global Const $PARAMFLAG_FRETVAL = 0x8
Global Const $PARAMFLAG_FOPT    = 0x10
Global Const $PARAMFLAG_FHASDEFAULT     = 0x20
Global Const $PARAMFLAG_FHASCUSTDATA    = 0x40

Global Enum _ ; FUNCFLAGS
    $FUNCFLAG_FRESTRICTED   = 0x1, _
    $FUNCFLAG_FSOURCE   = 0x2, _
    $FUNCFLAG_FBINDABLE = 0x4, _
    $FUNCFLAG_FREQUESTEDIT  = 0x8, _
    $FUNCFLAG_FDISPLAYBIND  = 0x10, _
    $FUNCFLAG_FDEFAULTBIND  = 0x20, _
    $FUNCFLAG_FHIDDEN   = 0x40, _
    $FUNCFLAG_FUSESGETLASTERROR = 0x80, _
    $FUNCFLAG_FDEFAULTCOLLELEM  = 0x100, _
    $FUNCFLAG_FUIDEFAULT    = 0x200, _
    $FUNCFLAG_FNONBROWSABLE = 0x400, _
    $FUNCFLAG_FREPLACEABLE  = 0x800, _
    $FUNCFLAG_FIMMEDIATEBIND    = 0x1000
Global Enum _ ; CALLCONV
    $CC_CDECL = 1, _
    $CC_MSCPASCAL = 2, _
    $CC_PASCAL = 2, _
    $CC_MACPASCAL = 3, _
    $CC_STDCALL = 4, _
    $CC_RESERVED = 5, _
    $CC_SYSCALL = 6, _
    $CC_MPWCDECL = 7, _
    $CC_MPWPASCAL = 8, _
    $CC_MAX = 9
Global Enum _ ; INVOKEKIND
    $INVOKE_FUNC = 0x01, _
    $INVOKE_PROPERTYGET = 0x02, _
    $INVOKE_PROPERTYPUT = 0x04, _
    $INVOKE_PROPERTYPUTREF = 0x08
Global Enum _ ; FUNCKIND
    $FUNC_VIRTUAL = 0, _
    $FUNC_PUREVIRTUAL = 1, _
    $FUNC_NONVIRTUAL = 2, _
    $FUNC_STATIC = 3, _
    $FUNC_DISPATCH = 4
Global Enum _ ; VARFLAGS
    $VARFLAG_FREADONLY  = 0x1, _
    $VARFLAG_FSOURCE    = 0x2, _
    $VARFLAG_FBINDABLE  = 0x4, _
    $VARFLAG_FREQUESTEDIT   = 0x8, _
    $VARFLAG_FDISPLAYBIND   = 0x10, _
    $VARFLAG_FDEFAULTBIND   = 0x20, _
    $VARFLAG_FHIDDEN    = 0x40, _
    $VARFLAG_FRESTRICTED    = 0x80, _
    $VARFLAG_FDEFAULTCOLLELEM   = 0x100, _
    $VARFLAG_FUIDEFAULT = 0x200, _
    $VARFLAG_FNONBROWSABLE  = 0x400, _
    $VARFLAG_FREPLACEABLE   = 0x800, _
    $VARFLAG_FIMMEDIATEBIND = 0x1000
Global Enum _ ; VARKIND
    $VAR_PERINSTANCE = 0, _
    $VAR_STATIC = 1, _
    $VAR_CONST = 2, _
    $VAR_DISPATCH = 3
Global Enum _ ; IMPLTYPEFLAGS
    $IMPLTYPEFLAG_FDEFAULT = 0x1, _
    $IMPLTYPEFLAG_FSOURCE = 0x2, _
    $IMPLTYPEFLAG_FRESTRICTED = 0x4, _
    $IMPLTYPEFLAG_FDEFAULTVTABLE = 0x800
Global Enum _ ; TYPEFLAGS
    $TYPEFLAG_FAPPOBJECT    = 0x1, _
    $TYPEFLAG_FCANCREATE    = 0x2, _
    $TYPEFLAG_FLICENSED = 0x4, _
    $TYPEFLAG_FPREDECLID    = 0x8, _
    $TYPEFLAG_FHIDDEN   = 0x10, _
    $TYPEFLAG_FCONTROL  = 0x20, _
    $TYPEFLAG_FDUAL = 0x40, _
    $TYPEFLAG_FNONEXTENSIBLE    = 0x80, _
    $TYPEFLAG_FOLEAUTOMATION    = 0x100, _
    $TYPEFLAG_FRESTRICTED   = 0x200, _
    $TYPEFLAG_FAGGREGATABLE = 0x400, _
    $TYPEFLAG_FREPLACEABLE  = 0x800, _
    $TYPEFLAG_FDISPATCHABLE = 0x1000, _
    $TYPEFLAG_FREVERSEBIND  = 0x2000, _
    $TYPEFLAG_FPROXY    = 0x4000
Global Enum _ ; TYPEKIND
    $TKIND_ENUM = 0, _
    $TKIND_RECORD = 1, _
    $TKIND_MODULE = 2, _
    $TKIND_INTERFACE = 3, _
    $TKIND_DISPATCH = 4, _
    $TKIND_COCLASS = 5, _
    $TKIND_ALIAS = 6, _
    $TKIND_UNION = 7, _
    $TKIND_MAX = 8
Global Enum _ ;LIBFLAGS
    $LIBFLAG_FRESTRICTED = 0x01, _
    $LIBFLAG_FCONTROL = 0x02, _
    $LIBFLAG_FHIDDEN = 0x04, _
    $LIBFLAG_FHASDISKIMAGE = 0x08
Global Enum _ ; SYSKIND
    $SYS_WIN16 = 0, _
    $SYS_WIN32 = 1, _
    $SYS_MAC = 2, _
    $SYS_WIN64 = 3
Global Enum _ ; REGKIND
    $REGKIND_DEFAULT, _
    $REGKIND_REGISTER, _
    $REGKIND_NONE

Global Const $tMEMBERID     = "long"
Global Const $tINVOKEKIND   = "int"

Global Enum $TDOC_NAME, $TDOC_DOCSTRING, $TDOC_HELPCONTEXT, $TDOC_HELPFILE 