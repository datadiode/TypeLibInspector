#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.1
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Global functions and helper interfaces for type information library.
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
Global $_TLI_debug = False
#Region ###TLI global functions###
;===============================================================================
; Function Name:   _TLI_TypeLibInfoFromFile
; Description:     _TLI_TypeLibInfoFromFile is used to create a TypeLibInfo object
;                  directly from a file (s. _ITypeLib_Load).
; Requirement(s):  ITypeLib, TLBINFTypeLibInfo.
; 
; Parameter(s):    $fileName - The file which contains the type library resource
;                  which is currently being represented. Type libraries can be
;                  either standalone files, which have the extension TLB, or
;                  contained as resources in other files (EXE, DLL, OCX, OLB, etc).
; 
; Return Value(s): New TypeLibInfo object. In case of failure 0 is returned and
;                  @error is set to HRESULT value.
; 
; Author(s):       doudou
;===============================================================================
Func _TLI_TypeLibInfoFromFile($fileName)
    Local $objTLib = _ITypeLib_Load($fileName)
    If IsObj($objTLib) Then Return TypeLibInfo_New($objTLib) 
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _TLI_TypeLibInfoFromRegistry
; Description:     _TLI_TypeLibInfoFromRegistry is used to create a TypeLibInfo
;                  object directly from the registry settings (s. _ITypeLib_LoadReg).
;                  The location of the type library to load is determined by
;                  the key HKEY_CLASSES_ROOT\TypeLib\TypeLibGuid\MajorVersion.MinorVersion\LCID\win32.
; Requirement(s):  ITypeLib, TLBINFTypeLibInfo.
; 
; Parameter(s):    $guid - The GUID of the library being loaded. Can be binary or
;                  string value.
;                  $verMajor - Major version number of the library being loaded.
;                  $verMinor - Minor version number of the library being loaded.
;                  $lcid - National language code of the library being loaded.
; 
; Return Value(s): New TypeLibInfo object. In case of failure 0 is returned and
;                  @error is set to HRESULT value.
; 
; Author(s):       doudou
;===============================================================================
Func _TLI_TypeLibInfoFromRegistry(Const ByRef $guid, $verMajor = 1, $verMinor = 0, $lcid = $LOCALE_SYSTEM_DEFAULT)
    Local $objTLib = _ITypeLib_LoadReg($guid, $verMajor, $verMinor, $lcid)
    If IsObj($objTLib) Then Return TypeLibInfo_New($objTLib) 
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _TLI_TypeInfoFromObject
; Description:     TypeInfo objects are generally retreived from a TypeLibInfo
;                  object. However, this also available directly from running
;                  objects. _TLI_TypeInfoFromObject allows you to inspect a
;                  running object, retrieving information about its properties,
;                  methods, and events (s. _IDispatch_Wrap, _IDispatch_GetTypeInfo).
; Requirement(s):  IDispatch, ITypeInfo, TLBINFTypeInfo.
; 
; Parameter(s):    $object - COM object retrived with ObjCreate() or ObjGet().
; 
; Return Value(s): New TypeInfo object. In case of failure 0 is returned and @error
;                  is set to HRESULT value.
; 
; Author(s):       doudou
;===============================================================================
Func _TLI_TypeInfoFromObject(ByRef $object)
    Local $objDisp = _IDispatch_Wrap($object)
    If IsObj($objDisp) Then
        Local $objTInfo = _IDispatch_GetTypeInfo($objDisp)
        $objDisp = 0
        If IsObj($objTInfo) Then Return _TLI_TypeInfoFromITypeInfo($objTInfo, 0)
    EndIf
    Return SetError(@error, @extended, 0)
EndFunc
;===============================================================================
; Function Name:   _TLI_TypeInfoFromITypeInfo
; Description:     Inside every TypeInfo is a reference to an ITypeInfo instance.
;                  In fact, the TypeInfo object can be viewed as a wrapper on the
;                  ITypeInfo interface which is easier to program against
;                  ITypeInfo itself. If you have an ITypeInfo reference and want
;                  to use TLI objects, then you can call _TLI_TypeInfoFromITypeInfo
;                  to generate a fully functional TypeInfo object directly from
;                  your ITypeInfo reference.
; Requirement(s):  ITypeInfo, TLBINFTypeInfo, TLBINFInterfaceInfo, TLBINFDispInterfaceInfo,
;                  TLBINFCoClassInfo, TLBINFIntrinsicAlias, TLBINFRecordInfo,
;                  TLBINFModuleInfo
; 
; Parameter(s):    $objITInfo - ITypeInfo interface wrapper.
;                  $objTLInfo - (optional) TypeLibInfo that contains type
;                  information, sumbit 0 if unknown.
; 
; Return Value(s): New TypeInfo object. In case of failure 0 is returned and @error
;                  is set to HRESULT value.
; 
; Author(s):       doudou
;===============================================================================
Func _TLI_TypeInfoFromITypeInfo(ByRef $objITInfo, Const ByRef $objTLInfo)
    If Not IsObj($objITInfo) Then Return SetError(-2, 0, 0)
    Local $result = TypeInfo_New($objITInfo, $objTLInfo)
    Switch $result.TypeKind()
        Case $TKIND_INTERFACE
            $result = InterfaceInfo_Inherit($result)
        Case $TKIND_DISPATCH
            $result = DispInterfaceInfo_Inherit($result)
        Case $TKIND_COCLASS
            $result = CoClassInfo_Inherit($result)
        Case $TKIND_ALIAS
            $result = IntrinsicAliasInfo_Inherit($result)
        Case $TKIND_RECORD, $TKIND_UNION, $TKIND_ENUM
            $result = RecordInfo_Inherit($result)
        Case $TKIND_MODULE
            $result = ModuleInfo_Inherit($result)
    EndSwitch
    Return $result
EndFunc
#EndRegion ;TLI global functions

#Region ###_TLI_StructWrapper construction/destruction###
Func _TLI_StructWrapper_New(Const ByRef $tagStruct, $pointer = 0)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_StructPtr", $ELSCOPE_PRIVATE, Number($pointer))
    _AutoItObject_AddProperty($result, "StructTag", $ELSCOPE_READONLY, $tagStruct)
    _AutoItObject_AddMethod($result, "GetPtr", "_TLI_StructWrapper_GetPtr")
    _AutoItObject_AddMethod($result, "GetData", "_TLI_StructWrapper_GetData")
;     _AutoItObject_AddDestructor($result, "_TLI_StructWrapper_Release")
    Return $result
EndFunc

; Func _TLI_StructWrapper_Release($oSelf)
;     ConsoleWrite("[DBG] _TLI_StructWrapper_Release" & @LF)
; EndFunc
#EndRegion ;_TLI_StructWrapper construction/destruction

#Region ###_TLI_StructWrapper public methods###
Func _TLI_StructWrapper_GetPtr($oSelf, $element = 0)
    If $oSelf._StructPtr Then
        Local $result = $oSelf._StructPtr
        If $element Then
            Local $t = DllStructCreate($oSelf.StructTag, $result)
            If 0 = @error Then $result = DllStructGetPtr($t, $element)
        EndIf
        Return SetError(@error, @extended, Number($result))
    EndIf
    Return SetError(1, 0, $oSelf._StructPtr)
EndFunc

Func _TLI_StructWrapper_GetData($oSelf, $element, $index = Default)
    Local $p = $oSelf.GetPtr()
    If $p Then
        Local $result = DllStructGetData(DllStructCreate($oSelf.StructTag, $p), $element, $index)
        If IsPtr($result) Then $result = Number($result)
        Return $result
    EndIf
    Return SetError(1, 0, 0)
EndFunc
#EndRegion ;_TLI_StructWrapper public methods
