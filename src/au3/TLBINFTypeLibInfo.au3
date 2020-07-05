#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.3
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Definitions for TLBINF:
;                 TypeLibInfo,
;                 TypeInfoCollection,
;                 TLibAttr
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 21 $
; $Date: 2010-05-23 21:30:22 +0200 (Sun, 23 May 2010) $
;
; ------------------------------------------------------------------------------
#include <WinAPI.au3>
#include <ITypeInfo.au3>
#include <ITypeLib.au3>
#include <AutoItObject.au3>
#include <TLBINFDocumentation.au3>
#include <SimpleCollection.au3>

Global $SYSKIND_NAMES[4] = ["win16", "win32", "mac", "win64"]

#Region ###TLibAttr construction/destruction###
Func TLibAttr_New(ByRef $objITypeLibInfo)
    Local $result = _AutoItObject_Create(_TLI_StructWrapper_New($tagTLIBATTR))
    _AutoItObject_AddProperty($result, "_objITLInfo", $ELSCOPE_PRIVATE, $objITypeLibInfo)
    _AutoItObject_AddMethod($result, "GetPtr", "TLibAttr_GetPtr")
    _AutoItObject_AddDestructor($result, "TLibAttr_Release")
    Return $result
EndFunc

Func TLibAttr_Release($oSelf)
;     ConsoleWrite("[DBG] TLibAttr_Release" & @LF)
    If $oSelf._StructPtr Then
        If $_TLI_debug Then ConsoleWrite("[DBG] Freeing TLIBATTR pointer" & @LF)
        $oSelf._objITLInfo.ReleaseTLibAttr(Number($oSelf._StructPtr))
    EndIf
    $oSelf._StructPtr = 0
    $oSelf._objITLInfo = 0
EndFunc
#EndRegion ;TLibAttr construction/destruction 

#Region ###TLibAttr publc methods###
Func TLibAttr_GetPtr($oSelf, $element = 0)
    If 0 = $oSelf._StructPtr Then
        Local $objITLInfo = $oSelf._objITLInfo
        $oSelf._StructPtr = Number(_ITypeLib_GetLibAttrPtr($objITLInfo))
    EndIf
    Return _TLI_StructWrapper_GetPtr($oSelf, $element)
EndFunc
#EndRegion ;TLibAttr public methods 

#Region ###TypeLibInfo construction/destruction###
Func TypeLibInfo_New(ByRef $objITypeLibInfo)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_objITLInfo", $ELSCOPE_READONLY, $objITypeLibInfo)
    _AutoItObject_AddProperty($result, "_docInfo", $ELSCOPE_PRIVATE, 0)
    _AutoItObject_AddProperty($result, "_tlibAttr", $ELSCOPE_PRIVATE, TLibAttr_New($objITypeLibInfo))
    _AutoItObject_AddMethod($result, "SysKind", "TypeLibInfo_SysKind")
    _AutoItObject_AddMethod($result, "SysKindString", "TypeLibInfo_SysKindString")
    _AutoItObject_AddMethod($result, "GUID", "TypeLibInfo_GUID")
    _AutoItObject_AddMethod($result, "MajorVersion", "TypeLibInfo_MajorVersion")
    _AutoItObject_AddMethod($result, "MinorVersion", "TypeLibInfo_MinorVersion")
    _AutoItObject_AddMethod($result, "AttributeMask", "TypeLibInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "LCID", "TypeLibInfo_LCID")
    _AutoItObject_AddMethod($result, "Name", "TLIObj_Name")
    _AutoItObject_AddMethod($result, "HelpString", "TLIObj_HelpString")
    _AutoItObject_AddMethod($result, "HelpContext", "TLIObj_HelpContext")
    _AutoItObject_AddMethod($result, "HelpFile", "TLIObj_HelpFile")
    _AutoItObject_AddMethod($result, "AttributeMask", "TypeLibInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "ContainingFile", "TypeLibInfo_ContainingFile")
    _AutoItObject_AddMethod($result, "TypeInfoCount", "TypeLibInfo_TypeInfoCount")
    _AutoItObject_AddMethod($result, "TypeInfos", "TypeLibInfo_TypeInfos")
    _AutoItObject_AddMethod($result, "GetTypeKind", "TypeLibInfo_GetTypeKind")
    _AutoItObject_AddMethod($result, "GetTypeInfoOfGuid", "TypeLibInfo_GetTypeInfoOfGuid")
    _AutoItObject_AddMethod($result, "_ReadDocumentation", "TypeLibInfo__ReadDocumentation", True)
    _AutoItObject_AddDestructor($result, "TypeLibInfo_Release")
    Return $result
EndFunc

Func TypeLibInfo_Release($oSelf)
;     ConsoleWrite("[DBG] TypeLibInfo_Release" & @LF)
    $oSelf._tlibAttr = 0
    $oSelf._docInfo = 0
    $oSelf._objITLInfo = 0
EndFunc
#EndRegion ;TypeLibInfo construction/destruction

#Region ###TypeLibInfo public property getters###
Func TypeLibInfo_SysKind($oSelf)
    Return $oSelf._tlibAttr.GetData("syskind")
EndFunc

Func TypeLibInfo_SysKindString($oSelf)
    Local $i = $oSelf.SysKind()
    If -1 < $i And $i < UBound($SYSKIND_NAMES) Then Return $SYSKIND_NAMES[$i]
    Return ""
EndFunc

Func TypeLibInfo_GUID($oSelf)
    Local $p = $oSelf._tlibAttr.GetPtr("guid")
    If $p Then Return _WinAPI_StringFromGUID($p)
    Return ""
EndFunc

Func TypeLibInfo_MajorVersion($oSelf)
    Return $oSelf._tlibAttr.GetData("wMajorVerNum")
EndFunc

Func TypeLibInfo_MinorVersion($oSelf)
    Return $oSelf._tlibAttr.GetData("wMinorVerNum")
EndFunc

Func TypeLibInfo_LCID($oSelf)
    Return $oSelf._tlibAttr.GetData("lcid")
EndFunc

Func TypeLibInfo_AttributeMask($oSelf)
    Return $oSelf._tlibAttr.GetData("wLibFlags")
EndFunc

Func TypeLibInfo_ContainingFile($oSelf)
    Return RegRead("HKCR\TypeLib\" & $oSelf.GUID() & "\" & $oSelf.MajorVersion() & "." & $oSelf.MinorVersion() & "\" & $oSelf.LCID() & "\win32", "")
EndFunc

Func TypeLibInfo_TypeInfoCount($oSelf)
    Local $obj = $oSelf._objITLInfo
    Return _ITypeLib_GetTypeInfoCount($obj)
EndFunc

Func TypeLibInfo_TypeInfos($oSelf)
    Return TypeInfoCollection_New($oSelf, $oSelf.TypeInfoCount())
EndFunc

Func TypeLibInfo_GetTypeKind($oSelf, $index)
    Local $obj = $oSelf._objITLInfo
    Return _ITypeLib_GetTypeInfoType($obj, $index)
EndFunc

Func TypeLibInfo_GetTypeInfoOfGuid($oSelf, $guid)
    Local $obj = $oSelf._objITLInfo
    Local $objTInfo = _ITypeLib_GetTypeInfoOfGuid($obj, $guid)
    Return _TLI_TypeInfoFromITypeInfo($objTInfo, $oSelf)
EndFunc
#EndRegion ;TypeLibInfo public property getters

#Region ###TypeLibInfo private methods###
Func TypeLibInfo__ReadDocumentation($oSelf)
    Local $result = $oSelf._docInfo
    If Not IsArray($result) Then
        If IsObj($oSelf._objITLInfo) Then
            Local $obj = $oSelf._objITLInfo
            Dim $result[4] = [1, 1, 1, 1]
            If _ITypeLib_GetDocumentation($obj, -1, $result) Then $oSelf._docInfo = $result
        EndIf
    EndIf
    Return $result
EndFunc
#EndRegion ;TypeLibInfo private methods

#Region ###TypeInfoCollection construction/destruction###
Func TypeInfoCollection_New(ByRef $objTypeLibInfo, $size)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objTypeLibInfo)
    _AutoItObject_AddProperty($result, "Count", $ELSCOPE_READONLY, $size)
    _AutoItObject_AddMethod($result, "Item", "TypeInfoCollection_Item")
    _AutoItObject_AddEnum($result, "SimpleCollection_EnumNext" ,"SimpleCollection_EnumReset")
    _AutoItObject_AddDestructor($result, "TypeInfoCollection_Release")
    Return $result
EndFunc

Func TypeInfoCollection_Release($oSelf)
;     ConsoleWrite("[DBG] TypeInfoCollection_Release" & @LF)
    $oSelf._tInfo = 0
    $oSelf.Count = 0
EndFunc
#EndRegion ;TypeInfoCollection construction/destruction

#Region ###TypeInfoCollection public property getters###
Func TypeInfoCollection_Item($oSelf, $index)
    If -1 < $index And $index < $oSelf.Count Then
        Local $obj = $oSelf._tInfo._objITLInfo
        Local $objTInfo = _ITypeLib_GetTypeInfo($obj, $index)
        Local $objParent = $oSelf._tInfo
        Return _TLI_TypeInfoFromITypeInfo($objTInfo, $objParent)
    EndIf
    Return SetError(-1, 0, 0)
EndFunc
#EndRegion ;TypeInfoCollection public property getters
