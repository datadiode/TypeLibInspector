#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.2
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Definitions for TLBINF:
;                 InterfaceInfo  (specialization of TypeInfo),
;                 MethodInfo (specialization of MemberInfo),
;                 ParameterCollection,
;                 ParameterInfo,
;                 FuncDesc,
;                 ParamDesc
; Link:           http://support.microsoft.com/kb/224331/en-us/
; $Revision: 19 $
; $Date: 2010-05-16 20:24:46 +0200 (Sun, 16 May 2010) $
;
; ------------------------------------------------------------------------------
#include <SimpleCollection.au3>

Global $INVOKEKIND_NAMES[4] = ["function", "propget", "propput", "propputref"]

#Region ###InterfaceInfo construction/destruction###
Func InterfaceInfo_New(ByRef $objITypeInfo, Const ByRef $objTypeLibInfo, $index = -1)
    Local $parent = TypeInfo_New($objITypeInfo, $objTypeLibInfo, $index)
    Return InterfaceInfo_Inherit($parent)
EndFunc

Func InterfaceInfo_Inherit(ByRef $objTypeInfo)
    Local $result = _AutoItObject_Create($objTypeInfo)
    _AutoItObject_AddMethod($result, "Methods", "InterfaceInfo_Methods")
    _AutoItObject_AddMethod($result, "Interfaces", "InterfaceInfo_Interfaces")
;     _AutoItObject_AddDestructor($result, "InterfaceInfo_Release")
    Return $result
EndFunc

; Func InterfaceInfo_Release($oSelf)
;     ConsoleWrite("[DBG] InterfaceInfo_Release" & @LF)
; EndFunc
#EndRegion ;InterfaceInfo construction/destruction

#Region ###InterfaceInfo public property getters###
Func InterfaceInfo_Methods($oSelf)
    Local $c = $oSelf._typeAttr.GetData("cFuncs")
    Return MemberCollection_New($oSelf, $c, $DESCKIND_FUNCDESC)
EndFunc

Func InterfaceInfo_Interfaces($oSelf)
    Local $c = $oSelf._typeAttr.GetData("cImplTypes")
    Return InterfaceCollection_New($oSelf, $c)
EndFunc
#EndRegion ;InterfaceInfo public property getters

#Region ###InterfaceCollection construction/destruction###
Func InterfaceCollection_New(ByRef $objInterfaceInfo, $size)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objInterfaceInfo)
    _AutoItObject_AddProperty($result, "Count", $ELSCOPE_READONLY, $size)
    _AutoItObject_AddMethod($result, "Item", "InterfaceCollection_Item")
    _AutoItObject_AddEnum($result, "SimpleCollection_EnumNext" ,"SimpleCollection_EnumReset")
    _AutoItObject_AddDestructor($result, "InterfaceCollection_Release")
    Return $result
EndFunc

Func InterfaceCollection_Release($oSelf)
;     ConsoleWrite("[DBG] InterfaceCollection_Release" & @LF)
    $oSelf._tInfo = 0
    $oSelf.Count = 0
EndFunc
#EndRegion ;InterfaceCollection construction/destruction

#Region ###InterfaceCollection public property getters###
Func InterfaceCollection_Item($oSelf, $index)
    If -1 < $index And $index < $oSelf.Count Then
        Local $obj = $oSelf._tInfo._objITInfo
        Local $objImplTInfo = _ITypeInfo_GetTypeOfImplType($obj, $index)
        Return _TLI_TypeInfoFromITypeInfo($objImplTInfo, 0)
    EndIf
    Return SetError(-1, 0, 0)
EndFunc
#EndRegion ;InterfaceCollection public property getters

#Region ###FuncDesc construction/destruction###
Func FuncDesc_New(ByRef $objITypeInfo, $index)
    Local $result = _AutoItObject_Create(_TLI_StructWrapper_New($tagFUNCDESC))
    _AutoItObject_AddProperty($result, "_objITInfo", $ELSCOPE_PRIVATE, $objITypeInfo)
    _AutoItObject_AddProperty($result, "_index", $ELSCOPE_PRIVATE, $index)
    _AutoItObject_AddMethod($result, "GetPtr", "FuncDesc_GetPtr")
    _AutoItObject_AddDestructor($result, "FuncDesc_Release")
    Return $result
EndFunc

Func FuncDesc_ForMember(ByRef $objMemberInfo)
    Local $objITInfo = $objMemberInfo._tInfo._objITInfo
    Return FuncDesc_New($objITInfo, $objMemberInfo.MemberNumber)
EndFunc

Func FuncDesc_Release($oSelf)
;     ConsoleWrite("[DBG] FuncDesc_Release" & @LF)
    If $oSelf._StructPtr Then
        If $_TLI_debug Then ConsoleWrite("[DBG] Freeing FUNCDESC pointer" & @LF)
        $oSelf._objITInfo.ReleaseFuncDesc(Number($oSelf._StructPtr))
    EndIf
    $oSelf._StructPtr = 0
    $oSelf._objITInfo = 0
EndFunc
#EndRegion ;FuncDesc construction/destruction 

#Region ###FuncDesc publc methods###
Func FuncDesc_GetPtr($oSelf, $element = 0)
    If 0 = $oSelf._StructPtr Then
        Local $objITInfo = $oSelf._objITInfo
        $oSelf._StructPtr = Number(_ITypeInfo_GetFuncDescPtr($objITInfo, $oSelf._index))
    EndIf
    Return _TLI_StructWrapper_GetPtr($oSelf, $element)
EndFunc
#EndRegion ;FuncDesc public methods 

#Region ###MethodInfo construction/destruction###
Func MethodInfo_New(ByRef $objInterfaceInfo, $index)
    Local $parent = MemberInfo_New($objInterfaceInfo, $index)
    Return MethodInfo_Inherit($parent)
EndFunc

Func MethodInfo_Inherit(ByRef $objMemberInfo)
    Local $result = _AutoItObject_Create($objMemberInfo)
    _AutoItObject_AddProperty($result, "_desc", $ELSCOPE_READONLY, FuncDesc_ForMember($objMemberInfo))
    _AutoItObject_AddProperty($result, "DescKind", $ELSCOPE_READONLY, $DESCKIND_FUNCDESC)
    _AutoItObject_AddMethod($result, "AttributeMask", "MethodInfo_AttributeMask")
    _AutoItObject_AddMethod($result, "ReturnType", "MethodInfo_ReturnType")
    _AutoItObject_AddMethod($result, "InvokeKind", "MethodInfo_InvokeKind")
    _AutoItObject_AddMethod($result, "InvokeKindString", "MethodInfo_InvokeKindString")
    _AutoItObject_AddMethod($result, "CallConv", "MethodInfo_CallConv")
    _AutoItObject_AddMethod($result, "FuncKind", "MethodInfo_FuncKind")
    _AutoItObject_AddMethod($result, "VTableOffset", "MethodInfo_VTableOffset")
    _AutoItObject_AddMethod($result, "Parameters", "MethodInfo_Parameters")
    _AutoItObject_AddDestructor($result, "MethodInfo_Release")
    Return $result
EndFunc

Func MethodInfo_Release($oSelf)
;     ConsoleWrite("[DBG] MethodInfo_Release" & @LF)
    $oSelf._desc = 0
EndFunc
#EndRegion ;MethodInfo construction/destruction

#Region ###MethodInfo public property getters###
Func MethodInfo_AttributeMask($oSelf)
    Return $oSelf._desc.GetData("wFuncFlags")
EndFunc

Func MethodInfo_ReturnType($oSelf)
    Local $p = $oSelf._desc.GetPtr("lpItemDesc")
    If $p Then
        Local $obj = $oSelf._tInfo
        Local $t = DllStructCreate($tagELEMDESC, $p)
        Return VarTypeInfo_New($obj, $t)
    EndIf
    Return 0
EndFunc

Func MethodInfo_InvokeKind($oSelf)
    Return $oSelf._desc.GetData("InvKind")
EndFunc

Func MethodInfo_InvokeKindString($oSelf)
    Local $i = $oSelf.InvokeKind()
    If 1 > $i Or $i > 8 Then $i = 1
    Return $INVOKEKIND_NAMES[Int(Log($i) / Log(2))]
EndFunc

Func MethodInfo_CallConv($oSelf)
    Return $oSelf._desc.GetData("callconv")
EndFunc

Func MethodInfo_FuncKind($oSelf)
    Return $oSelf._desc.GetData("funckind")
EndFunc

Func MethodInfo_VTableOffset($oSelf)
    Return $oSelf._desc.GetData("oVft")
EndFunc

Func MethodInfo_Parameters($oSelf)
    Local $c = $oSelf._desc.GetData("cParams")
    Return ParameterCollection_New($oSelf, $c)
EndFunc
#EndRegion ;MethodInfo public property getters

#Region ###ParameterCollection construction/destruction###
Func ParameterCollection_New(ByRef $objMethodInfo, $size)
    Local $result = _AutoItObject_Create()
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objMethodInfo)
    _AutoItObject_AddProperty($result, "Count", $ELSCOPE_READONLY, $size)
    _AutoItObject_AddProperty($result, "_names", $ELSCOPE_PRIVATE, 0)
    _AutoItObject_AddMethod($result, "Item", "ParameterCollection_Item")
    _AutoItObject_AddMethod($result, "_ReadParamNames", "ParameterCollection__ReadParamNames", True)
    _AutoItObject_AddEnum($result, "SimpleCollection_EnumNext", "SimpleCollection_EnumReset")
    _AutoItObject_AddDestructor($result, "ParameterCollection_Release")
    Return $result
EndFunc

Func ParameterCollection_Release($oSelf)
;     ConsoleWrite("[DBG] ParameterCollection_Release" & @LF)
    $oSelf.Count = 0
    $oSelf._names = 0
    $oSelf._tInfo = 0
EndFunc
#EndRegion ;ParameterCollection construction/destruction

#Region ###ParameterCollection public property getters###
Func ParameterCollection_Item($oSelf, $index)
    Local $names = $oSelf._ReadParamNames()
    If Not IsInt($index) Then
        For $i = 1 To Ubound($names) - 1
            If String($index) = $names[$i] Then
                $index = $i - 1
                ExitLoop
            EndIf
        Next
    EndIf
    If -1 < $index And $index < $oSelf.Count Then
        Local $obj = $oSelf._tInfo
        Return ParameterInfo_New($obj, StringStripWS($names[$index + 1], 3), $index)
    EndIf
    Return SetError(-1, 0, 0)
EndFunc
#EndRegion ;ParameterCollection public property getters

#Region ###ParameterCollection private methods###
Func ParameterCollection__ReadParamNames($oSelf)
    Local $result = $oSelf._names
    If Not IsArray($result) Then
        Dim $result[$oSelf.Count + 1]
        Local $obj = $oSelf._tInfo._tInfo._objITInfo
        If 0 < _ITypeInfo_GetNames($obj, $oSelf._tInfo.MemberId(), $result) Then $oSelf._names = $result
    EndIf
    Return $result
EndFunc
#EndRegion ;ParameterCollection private methods

#Region ###ParamDesc construction/destruction###
Func ParamDesc_New(ByRef $objFuncDesc, $index)
    Local $p = _carray_ptr($objFuncDesc.GetData("lprgElemDescParam"), $index, "tagELEMDESC")
    Local $result = _AutoItObject_Create(_TLI_StructWrapper_New($tagELEMDESC, $p))
    Return $result
EndFunc
#EndRegion ;ParamDesc construction/destruction 

#Region ###ParameterInfo construction/destruction###
Func ParameterInfo_New(ByRef $objMethodInfo, $name, $index)
    Local $result = _AutoItObject_Create()
    If 0 = StringLen($name) Then
        Local $ik = $objMethodInfo.InvokeKind()
        If 0 < BitAnd($INVOKE_PROPERTYPUT, $ik) Or 0 < BitAnd($INVOKE_PROPERTYPUTREF, $ik) Then
            If 0 < BitAnd($INVOKE_PROPERTYPUTREF, $ik) Then
                $name = $objMethodInfo.Name & "Ref"
            Else
                $name = $objMethodInfo.Name & "Val"
            EndIf
        Else
            $name = "arg" & ($index + 1)
        EndIf
    EndIf
    _AutoItObject_AddProperty($result, "_tInfo", $ELSCOPE_PRIVATE, $objMethodInfo)
    Local $funcDesc = $objMethodInfo._desc
    _AutoItObject_AddProperty($result, "_desc", $ELSCOPE_PRIVATE, ParamDesc_New($funcDesc, $index))
    _AutoItObject_AddProperty($result, "Name", $ELSCOPE_READONLY, $name)
    _AutoItObject_AddProperty($result, "ParameterNumber", $ELSCOPE_READONLY, $index)
    _AutoItObject_AddMethod($result, "Default", "ParameterInfo_Default")
    _AutoItObject_AddMethod($result, "Optional", "ParameterInfo_Optional")
    _AutoItObject_AddMethod($result, "HasCustomData", "ParameterInfo_HasCustomData")
    _AutoItObject_AddMethod($result, "Flags", "ParameterInfo_Flags")
    _AutoItObject_AddMethod($result, "DefaultValue", "ParameterInfo_DefaultValue")
    _AutoItObject_AddMethod($result, "VarTypeInfo", "ParameterInfo_VarTypeInfo")
    _AutoItObject_AddDestructor($result, "ParameterInfo_Release")
    Return $result
EndFunc

Func ParameterInfo_Release($oSelf)
;     ConsoleWrite("[DBG] ParameterInfo_Release" & @LF)
    $oSelf._desc = 0
    $oSelf._tInfo = 0
EndFunc
#EndRegion ;ParameterInfo construction/destruction

#Region ###ParameterInfo public property getters###
Func ParameterInfo_Default($oSelf)
    Local $f = $oSelf.Flags()
    Return 0 < BitAnd($PARAMFLAG_FHASDEFAULT, $f)
EndFunc

Func ParameterInfo_Optional($oSelf)
    Local $f = $oSelf.Flags()
    Return 0 < BitAnd($PARAMFLAG_FOPT, $f)
EndFunc

Func ParameterInfo_HasCustomData($oSelf)
    Local $f = $oSelf.Flags()
    Return 0 < BitAnd($PARAMFLAG_FHASCUSTDATA, $f)
EndFunc

Func ParameterInfo_DefaultValue($oSelf)
    Local $f = $oSelf.Flags()
    If 0 < BitAnd($PARAMFLAG_FOPT, $f) Or 0 < BitAnd($PARAMFLAG_FHASDEFAULT, $f) Then
        Local $p = $oSelf._desc.GetData("lpParamDescEx")
        If $p Then
            Local $t = DllStructCreate($tagPARAMDESCEX, $p)
            Local $result = _AutoItObject_VariantRead(DllStructGetPtr($t, "vt"))
            If IsPtr($result) Then $result = Number($result)
            Return $result
        EndIf
    EndIf
    Return 0
EndFunc

Func ParameterInfo_Flags($oSelf)
    Return $oSelf._desc.GetData("wParamFlags")
EndFunc

Func ParameterInfo_VarTypeInfo($oSelf)
    Local $p = $oSelf._desc.GetPtr()
    If $p Then
        Local $tElemDesc = DllStructCreate($tagELEMDESC, $p)
        Local $obj = $oSelf._tInfo._tInfo
        Return VarTypeInfo_New($obj, $tElemDesc)
    EndIf
    Return 0 
EndFunc
#EndRegion ;ParameterInfo public property getters
