#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.3
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Definitions for TLBINF:
;                 TypeInfo,
;                 MemberInfo,
;                 MemberCollection,
;                 ParameterInfo,
;                 TypeAttr
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 20 $
; $Date: 2010-05-23 21:18:26 +0200 (Sun, 23 May 2010) $
;
; ------------------------------------------------------------------------------
#include <WinAPI.au3>
#include <ITypeInfo.au3>
#include <ITypeLib.au3>
#include <AutoItObject.au3>
#include <TLBINFDocumentation.au3>

Global Enum $DESCKIND_NONE, $DESCKIND_FUNCDESC, $DESCKIND_VARDESC
Global $TYPEKIND_NAMES[$TKIND_MAX] = ["enum", "struct", "module", "interface", "dispinterface", "coclass", "alias", "union"]

#Region ###TypeAttr construction/destruction###
Func TypeAttr_New(ByRef $objITypeInfo)
    Local $result = _AutoItObject_Create(_TLI_StructWrapper_New($tagTYPEATTR))
    _AutoItObject_AddProperty($result, "_objITInfo", $ELSCOPE_PRIVATE, $objITypeInfo)
    _AutoItObject_AddMethod($result, "GetPtr", "TypeAttr_GetPtr")
    _AutoItObject_AddDestructor($result, "TypeAttr_Release")
    Return $result
EndFunc

Func TypeAttr_Release($oSelf)
;     ConsoleWrite("[DBG] TypeAttr_Release" & @LF)
    If $oSelf._StructPtr Then
        If $_TLI_debug Then ConsoleWrite("[DBG] Freeing TYPEATTR pointer" & @LF)
        $oSelf._objITInfo.ReleaseTypeAttr(Number($oSelf._StructPtr))
    EndIf
    $oSelf._StructPtr = 0
    $oSelf._objITInfo = 0
EndFunc
#EndRegion ;TypeAttr construction/destruction 

#Region ###TypeAttr publc methods###
Func TypeAttr_GetPtr($oSelf, $element = 0)
    If 0 = $oSelf._StructPtr Then
        Local $objITInfo = $oSelf._objITInfo
        $oSelf._StructPtr = Number(_ITypeInfo_GetTypeAttrPtr($objITInfo))
    EndIf
    Return _TLI_StructWrapper_GetPtr($oSelf, $element)
EndFunc
#EndRegion ;TypeAttr public methods 

#Region ###TypeInfo construction/destruction###
Func TypeInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_objITInfo", $ELSCOPE_READONLY, $objITypeInfo)
    _AutoItObject_AddProperty($result, "_docInfo", $ELSCOPE_PRIVATE, 0)
    _AutoItObject_AddProperty($result, "_typeAttr", $ELSCOPE_PRIVATE, TypeAttr_New($objITypeInfo))
    _AutoItObject_AddProperty($result, "_tlibInfo", $ELSCOPE_PRIVATE, $objTypeLibInfo)
    _AutoItObject_AddProperty($result, "_number", $ELSCOPE_PRIVATE, $index)
    _AutoItObject_AddMethod($result, "TypeKind", "TypeInfo_TypeKind")
    _AutoItObject_AddMethod($result, "TypeKindString", "TypeInfo_TypeKindString")
    _AutoItObject_AddMethod($result, "GUID", "TypeInfo_GUID")
    _AutoItObject_AddMethod($result, "MajorVersion", "TypeInfo_MajorVersion")
    _AutoItObject_AddMethod($result, "MinorVersion", "TypeInfo_MinorVersion")
    _AutoItObject_AddMethod($result, "LCID", "TypeInfo_LCID")
    _AutoItObject_AddMethod($result, "AttributeMask", "TypeInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "Name", "TLIObj_Name")
    _AutoItObject_AddMethod($result, "HelpString", "TLIObj_HelpString")
    _AutoItObject_AddMethod($result, "HelpContext", "TLIObj_HelpContext")
    _AutoItObject_AddMethod($result, "HelpFile", "TLIObj_HelpFile")
    _AutoItObject_AddMethod($result, "Parent", "TypeInfo_Parent")
    _AutoItObject_AddMethod($result, "TypeInfoNumber", "TypeInfo_TypeInfoNumber")
    _AutoItObject_AddMethod($result, "_ReadDocumentation", "TypeInfo__ReadDocumentation", True)
    _AutoItObject_AddDestructor($result, "TypeInfo_Release")
    Return $result
EndFunc

Func TypeInfo_Release($oSelf)
;     ConsoleWrite("[DBG] TypeInfo_Release" & @LF)
    $oSelf._typeAttr = 0
    $oSelf._docInfo = 0
    $oSelf._tlibInfo = 0
    $oSelf._objITInfo = 0
EndFunc
#EndRegion ;TypeInfo construction/destruction 

#Region ###TypeInfo public property getters###
Func TypeInfo_TypeKind($oSelf)
    Return $oSelf._typeAttr.GetData("typekind")
EndFunc

Func TypeInfo_TypeKindString($oSelf)
    Local $i = $oSelf.TypeKind()
    If -1 < $i And $TKIND_MAX > $i Then Return $TYPEKIND_NAMES[$i]
    Return ""
EndFunc

Func TypeInfo_GUID($oSelf)
    Local $p = $oSelf._typeAttr.GetPtr("guid")
    If $p Then Return _WinAPI_StringFromGUID($p)
    Return ""
EndFunc

Func TypeInfo_MajorVersion($oSelf)
    Return $oSelf._typeAttr.GetData("wMajorVerNum")
EndFunc

Func TypeInfo_MinorVersion($oSelf)
    Return $oSelf._typeAttr.GetData("wMinorVerNum")
EndFunc

Func TypeInfo_LCID($oSelf)
    Return $oSelf._typeAttr.GetData("lcid")
EndFunc

Func TypeInfo_AttributeMask($oSelf)
    Return $oSelf._typeAttr.GetData("wTypeFlags")
EndFunc

Func TypeInfo_Parent($oSelf)
    If Not IsObj($oSelf._tlibInfo) Or 0 > $oSelf._number Then
        Local $index
        Local $objITInfo = $oSelf._objITInfo
        Local $objITLInfo = _ITypeInfo_GetContainingTypeLib($objITInfo, $index)
        $oSelf._number = $index
        If IsObj($objITLInfo) And Not IsObj($oSelf._tlibInfo) Then $oSelf._tlibInfo = TypeLibInfo_New($objITLInfo)
    EndIf
    Return $oSelf._tlibInfo
EndFunc

Func TypeInfo_TypeInfoNumber($oSelf)
    $oSelf.Parent()
    Return $oSelf._number
EndFunc
#EndRegion ;TypeInfo public property getters

#Region ###TypeInfo private methods###
Func TypeInfo__ReadDocumentation($oSelf)
    Local $result = $oSelf._docInfo
    If Not IsArray($result) Then
        $oSelf.Parent()
        If IsObj($oSelf._tlibInfo) Then
            Local $obj = $oSelf._tlibInfo._objITLInfo
            Dim $result[4] = [1, 1, 1, 1]
            If _ITypeLib_GetDocumentation($obj, $oSelf._number, $result) Then $oSelf._docInfo = $result
        EndIf
    EndIf
    Return $result
EndFunc
#EndRegion ;TypeInfo private methods

#Region ###MemberCollection construction/destruction###
Func MemberCollection_New(ByRef $objTypeInfo, $size, $memberType)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objTypeInfo)
    _AutoItObject_AddProperty($result, "_membType", $ELSCOPE_PRIVATE, $memberType)
    _AutoItObject_AddProperty($result, "Count", $ELSCOPE_READONLY, $size)
    _AutoItObject_AddMethod($result, "Item", "MemberCollection_Item")
    _AutoItObject_AddEnum($result, "SimpleCollection_EnumNext" ,"SimpleCollection_EnumReset")
    _AutoItObject_AddDestructor($result, "MemberCollection_Release")
    Return $result
EndFunc

Func MemberCollection_Release($oSelf)
;     ConsoleWrite("[DBG] MemberCollection_Release" & @LF)
    $oSelf._tInfo = 0
    $oSelf.Count = 0
EndFunc
#EndRegion ;MemberCollection construction/destruction

#Region ###MemberCollection public property getters###
Func MemberCollection_Item($oSelf, $index)
    If -1 < $index And $index < $oSelf.Count Then
        Local $obj = $oSelf._tInfo
        Switch $oSelf._membType
            Case $DESCKIND_FUNCDESC
                Return MethodInfo_New($obj, $index)
            Case $DESCKIND_VARDESC
                Return PropertyInfo_New($obj, $index)
            Case Else
                Return MemberInfo_New($obj, $index)
        EndSwitch
    EndIf
    Return SetError(-1, 0, 0)
EndFunc
#EndRegion ;MemberCollection public property getters

#Region ###MemberInfo construction/destruction###
Func MemberInfo_New(ByRef $objTypeInfo, $index)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_READONLY, $objTypeInfo)
    _AutoItObject_AddProperty($result, "_docInfo", $ELSCOPE_PRIVATE, 0)
    _AutoItObject_AddProperty($result, "_desc", $ELSCOPE_PRIVATE, 0)
    _AutoItObject_AddProperty($result, "DescKind", $ELSCOPE_READONLY, $DESCKIND_NONE)
    _AutoItObject_AddProperty($result, "MemberNumber", $ELSCOPE_READONLY, $index)
    _AutoItObject_AddMethod($result, "MemberId", "MemberInfo_MemberId")
    _AutoItObject_AddMethod($result, "AttributeMask", "MemberInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "ReturnType", "MemberInfo_ReturnType")
    _AutoItObject_AddMethod($result, "Name", "TLIObj_Name")
    _AutoItObject_AddMethod($result, "HelpString", "TLIObj_HelpString")
    _AutoItObject_AddMethod($result, "HelpContext", "TLIObj_HelpContext")
    _AutoItObject_AddMethod($result, "HelpFile", "TLIObj_HelpFile")
    _AutoItObject_AddMethod($result, "_ReadDocumentation", "MemberInfo__ReadDocumentation", True)
    _AutoItObject_AddDestructor($result, "MemberInfo_Release")
    Return $result
EndFunc

Func MemberInfo_Release($oSelf)
;     ConsoleWrite("[DBG] MemberInfo_Release" & @LF)
    $oSelf._docInfo = 0
    $oSelf._desc = 0
    $oSelf._tInfo = 0
EndFunc
#EndRegion ;MemberInfo construction/destruction

#Region ###MemberInfo public property getters###
Func MemberInfo_MemberId($oSelf)
    If IsObj($oSelf._desc) Then Return $oSelf._desc.GetData("MemID")
    Return -1
EndFunc

Func MemberInfo_AttributeMask($oSelf)
    Return 0
EndFunc

Func MemberInfo_ReturnType($oSelf)
    Return 0
EndFunc
#EndRegion ;MemberInfo public property getters

#Region ###MemberInfo private methods###
Func MemberInfo__ReadDocumentation($oSelf)
    Local $result = $oSelf._docInfo
    If Not IsArray($result) And IsObj($oSelf._tInfo) Then
        Local $obj = $oSelf._tInfo._objITInfo
        Dim $result[4] = [1, 1, 1, 1]
        If _ITypeInfo_GetDocumentation($obj, $oSelf.MemberId(), $result) Then $oSelf._docInfo = $result
    EndIf
    Return $result
EndFunc
#EndRegion ;MemberInfo private methods