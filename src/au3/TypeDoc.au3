#include-once
; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    DOM manipulation functions for TypeLibInspector.
; $Revision: 1.4 $
; $Date: 2010/07/29 14:21:57 $
;
; ------------------------------------------------------------------------------
Global Const $CHARS_ALPHA   = "abcdefghijklmnopqrstuvwxyz"
Global Const $CHARS_ALPHAU  = StringUpper($CHARS_ALPHA)
Global Const $TYPEDOC_QFMT_IGNORECASE   = "translate(%s, '%s', '%s')"
Global Const $TYPEDOC_QFMT_CMP[4] = ["%s = '%s'", "contains(%s, '%s')", "starts-with(%s, '%s')", "substring(%s, string-length(%s) - %d + 1, %d) = '%s'"]
Global Enum $TYPEDOC_QEQUALS, $TYPEDOC_QCONTAINS, $TYPEDOC_QSTARTS, $TYPEDOC_QENDS

Global Const $TYPEDOC_TKIND[$TKIND_MAX] = ["Enum", "Record", "Module", "Interface", "DispInterface", "CoClass", "Alias", "Union"]
Global Const $TYPEDOC_STDMEMB_IGNORE   = "|1610612736|1610612737|1610612738|1610678272|1610678273|1610678274|1610678275|"
Global Enum $TYPEDOC_DUALMODE_OLEAUT, $TYPEDOC_DUALMODE_VTABLE

Global $typeDoc_objDOM = 0
Global $typeDoc_cTypes = 0
Global $typeDoc_objTypeColl = 0
Global $typeDoc_optLazyExpand = False
Global $typeDoc_optDualIfaceMode = $TYPEDOC_DUALMODE_OLEAUT
Global $typeDoc_optCollBase = 0
Global $typeDoc_optCollUBDecr = 1

Func TypeDoc_Create($progID = "MSXML.DOMDocument")
    Local $result = False
    $typeDoc_objDOM = ObjCreate($progID)
    If IsObj($typeDoc_objDOM) Then
        $typeDoc_objDOM.async = False
        $typeDoc_objDOM.validateOnParse = False
        $typeDoc_objDOM.setProperty("SelectionLanguage", "XPath")
        ;$typeDoc_objDOM.setProperty("SelectionNamespaces", "xmlns:ms='urn:schemas-microsoft-com:xslt'")
        $result = $typeDoc_objDOM.loadXML("<?xml version=""1.0"" encoding=""UTF-8""?><TypeLib/>")
    EndIf
    Return $result
EndFunc

Func TypeDoc_Init(ByRef $objTypeLibInfo, $lazyExpand = False)
    $typeDoc_optLazyExpand = $lazyExpand
    If IsObj($app_objTLI) Then
        $typeDoc_optCollBase = 1
        $typeDoc_optCollUBDecr = 0
    Else
        $typeDoc_optCollBase = 0
        $typeDoc_optCollUBDecr = 1
    EndIf
    If IsObj($typeDoc_objDOM) Then
        Local $result = $typeDoc_objDOM.documentElement
        TypeDoc_InitInfoNode($result, $objTypeLibInfo)
        $result.setAttribute("sysKind", $objTypeLibInfo.SysKind)
        $result.setAttribute("lcid", $objTypeLibInfo.LCID)
        $result.setAttribute("path", $objTypeLibInfo.ContainingFile)
        $typeDoc_objTypeColl = $objTypeLibInfo.TypeInfos
        Return $result
    EndIf
    Return 0
EndFunc

Func TypeDoc_Destroy()
    $typeDoc_cTypes = 0
    $typeDoc_objDOM = 0
    $typeDoc_objTypeColl = 0
EndFunc

Func TypeDoc_PrettyPrint()
    Local $reader = ObjCreate("MSXML2.SAXXMLReader.4.0")
    If IsObj($reader) Then
        Local $writer = ObjCreate("MSXML2.MXXMLWriter.4.0")
        $writer.indent = True
        $writer.standalone = True
        $writer.encoding = "UTF-8"
        $writer.output = $typeDoc_objDOM
        $reader.contentHandler = $writer
        $reader.parse($typeDoc_objDOM)
    EndIf
EndFunc

Func TypeDoc_MakeQuery($search, $tests, $ignoreCase = True)
    Local $result = ""
    If $ignoreCase Then $search = StringLower($search)
    Local $fmt = $TYPEDOC_QCONTAINS
    If "*" = StringLeft($search, 1) And "*" = StringRight($search, 1) Then
        $search = StringMid($search, 2, StringLen($search) - 2)
    ElseIf "*" = StringLeft($search, 1) Then
        $search = StringRight($search, StringLen($search) - 1)
        $fmt = $TYPEDOC_QENDS
    ElseIf "*" = StringRight($search, 1) Then
        $search = StringLeft($search, StringLen($search) - 1)
        $fmt = $TYPEDOC_QSTARTS
    ElseIf """" = StringLeft($search, 1) And """" = StringRight($search, 1) Then
        $search = StringMid($search, 2, StringLen($search) - 2)
        $fmt = $TYPEDOC_QEQUALS
    EndIf
    Local $parts = StringSplit($tests, ",")
    For $i = 1 To $parts[0]
        If 0 < StringLen($result) Then $result &= " or "
        If $ignoreCase Then $parts[$i] = StringFormat($TYPEDOC_QFMT_IGNORECASE, $parts[$i], $CHARS_ALPHAU, $CHARS_ALPHA)
        Switch $fmt
            Case $TYPEDOC_QENDS
                $result &= StringFormat($TYPEDOC_QFMT_CMP[$fmt], $parts[$i], $parts[$i], StringLen($search), StringLen($search), $search)
            Case Else
                $result &= StringFormat($TYPEDOC_QFMT_CMP[$fmt], $parts[$i], $search)
        EndSwitch
    Next
    Return $result
EndFunc

Func TypeDoc_ParseTypeLibInfo($feedback = False)
    If Not IsObj($typeDoc_objDOM) Or Not IsObj($typeDoc_objTypeColl) Then Return 0
    Local $coll = $typeDoc_objDOM.createElement("Types")
;     For $objTypeInfo In $app_objTypeLib.TypeInfos
    Local $objTypeInfo
    For $i = 0 To $typeDoc_objTypeColl.Count - 1
        If IsObj($app_objTLI) Then
            $objTypeInfo = $typeDoc_objTypeColl.IndexedItem($i)
        Else
            $objTypeInfo = $typeDoc_objTypeColl.Item($i)
        EndIf
        If $feedback Then ProgressSet($i * 100 / $typeDoc_objTypeColl.Count, $objTypeInfo.Name)
        $coll.appendChild(TypeDoc_CreateTypeInfoNode($objTypeInfo))
    Next
    $objTypeInfo = 0
    Local $result = $typeDoc_objTypeColl.Count
    If Not $typeDoc_optLazyExpand Then $typeDoc_objTypeColl = 0
    $typeDoc_objDOM.documentElement.appendChild($coll)
    Return $result
EndFunc

Func TypeDoc_InitInfoNode(ByRef $node, Const ByRef $info, $nodeType = 0)
    With $info
        _GUICtrlStatusBar_SetText($frmMain_status, .Name, 0)
        $node.setAttribute("name", .Name)
        Local $kid = $typeDoc_objDOM.createElement("Attributes")
        $kid.setAttribute("mask", .AttributeMask)
        $node.appendChild($kid)

        If 0 < .AttributeMask Then
            Switch $nodeType
                Case 0
                    TypeDoc_ExplodeFlags(.AttributeMask, $LIBFLAG_FRESTRICTED, $LIBFLAG_FHASDISKIMAGE, $kid)
                Case 1
                    TypeDoc_ExplodeFlags(.AttributeMask, $TYPEFLAG_FAPPOBJECT, $TYPEFLAG_FPROXY, $kid)
                Case 2
                    TypeDoc_ExplodeFlags(.AttributeMask, $FUNCFLAG_FRESTRICTED, $FUNCFLAG_FIMMEDIATEBIND, $kid)
                Case 3
                    TypeDoc_ExplodeFlags(.AttributeMask, $VARFLAG_FREADONLY, $VARFLAG_FIMMEDIATEBIND, $kid)
            EndSwitch
        EndIf
        If 1 < $nodeType Then
            $node.setAttribute("dispid", .MemberId)
            If Not IsObj($app_objTLI) Then $node.setAttribute("number", .MemberNumber)
        Else
            $node.setAttribute("guid", .GUID)
            $kid = $typeDoc_objDOM.createElement("Version")
            $kid.setAttribute("major", .MajorVersion)
            $kid.setAttribute("minor", .MinorVersion)
            $node.appendChild($kid)
        EndIf
        $kid = $typeDoc_objDOM.createElement("Help")
        If .HelpFile Then $kid.setAttribute("file", .HelpFile)
        $kid.setAttribute("context", .HelpContext)
        If .HelpString Then $kid.appendChild($typeDoc_objDOM.createTextNode(.HelpString))
        $node.appendChild($kid)
    EndWith
EndFunc

Func TypeDoc_CreateTypeInfoNode(Const ByRef $info)
    Local $result = 0
    Local $typeKind = $info.TypeKind
    $typeDoc_cTypes += 1
    If -1 < $typeKind And $TKIND_MAX > $typeKind Then
        $result = $typeDoc_objDOM.createElement($TYPEDOC_TKIND[$typeKind])
    Else
        $result = $typeDoc_objDOM.createElement("UnknownType")
    EndIf
    TypeDoc_InitInfoNode($result, $info, 1)
    $result.setAttribute("kind", $typeKind)
    $result.setAttribute("number", $info.TypeInfoNumber)
    If $typeDoc_optLazyExpand Then
        $result.setAttribute("tldapp-expand", "1")
        Return $result
    EndIf
    TypeDoc_ExpandTypeInfoNode($info, $node)
    Return $result
EndFunc

Func TypeDoc_UnexpandTypeInfoNode(ByRef $node)
    If "1" = $node.getAttribute("tldapp-expand") Then Return
    Local $cPreserved = 0
    Local $i = 0
    Local $kid
    While $cPreserved < $node.childNodes.length()
        $kid = $node.childNodes.item($i)
        If "Attributes" = $kid.nodeName Or "Help" = $kid.nodeName Then
            $cPreserved += 1
            $i += 1
        Else
            $node.removeChild($kid)
        EndIf
    WEnd
    $node.setAttribute("tldapp-expand", "1")
EndFunc

Func TypeDoc_ExpandTypeInfoNode(Const ByRef $oInfo, ByRef $result, $feedback = False)
    If "1" <> $result.getAttribute("tldapp-expand") Then Return
    Local $info = $oInfo
    If Not IsObj($info) Then
        If IsObj($app_objTLI) Then
            $info = $typeDoc_objTypeColl.IndexedItem(Number($result.getAttribute("number")))
        Else
            $info = $typeDoc_objTypeColl.Item(Number($result.getAttribute("number")))
        EndIf
    EndIf
    If Not IsObj($info) Then Return

    If $feedback Then GUISetCursor(15, 1, $frmMain)

    Local $typeKind = $info.TypeKind
    Local $coll, $collInfos, $objSubinfo
    If $TKIND_DISPATCH = $typeKind And BitAnd($TYPEFLAG_FDUAL, $info.AttributeMask) Then
        $objSubinfo = $info.VTableInterface
        If IsObj($objSubinfo) And $TYPEDOC_DUALMODE_VTABLE = $typeDoc_optDualIfaceMode Then
            $info = $objSubinfo
            $typeKind = $TKIND_INTERFACE
        Else
            TypeDoc_ListBaseInterfaces($objSubinfo, $result)
        EndIf
    EndIf
    Switch $typeKind
        Case $TKIND_COCLASS
            If BitAnd($TYPEFLAG_FCANCREATE, $info.AttributeMask) Then $result.setAttribute("progid", _COM_ProgIDFromCLSID($info.GUID))
            $coll = $typeDoc_objDOM.createElement("Interfaces")
            $collInfos = $info.Interfaces
            For $i = $typeDoc_optCollBase To $collInfos.Count - $typeDoc_optCollUBDecr
                $objSubinfo = $collInfos.Item($i)
                Local $kid = $typeDoc_objDOM.createElement("Impl")
                Local $attr = $typeDoc_objDOM.createElement("Attributes")
                Local $mask = 0
                If IsObj($app_objTLI) Then
                    $mask = $objSubinfo.AttributeMask
                Else
                    $mask = $info.ImplTypeFlags($i)
                EndIf
                $attr.setAttribute("mask", $mask)
                $kid.appendChild($attr)
                TypeDoc_ExplodeFlags($mask, $IMPLTYPEFLAG_FDEFAULT, $IMPLTYPEFLAG_FDEFAULTVTABLE, $attr)
                $kid.appendChild(TypeDoc_CreateTypeRefNode($objSubinfo))
                $coll.appendChild($kid)
            Next
            $objSubinfo = 0
            $result.appendChild($coll)
        Case $TKIND_INTERFACE
            TypeDoc_ListBaseInterfaces($info, $result)
            If IsObj($app_objTLI) Then $info = $info.VTableInterface
        Case $TKIND_ALIAS
            Local $objVarType = $info.ResolvedType
            If IsObj($objVarType) Then
                $coll = $typeDoc_objDOM.createElement("Resolved")
                $coll.appendChild(TypeDoc_CreateVarInfoNode($objVarType))
                $result.appendChild($coll)
            EndIf
    EndSwitch
    Switch $typeKind
        Case $TKIND_DISPATCH, $TKIND_MODULE, $TKIND_RECORD, $TKIND_UNION, $TKIND_ENUM
            $coll = $typeDoc_objDOM.createElement("Properties")
            If IsObj($app_objTLI) Then
                $collInfos = $info.Members
            Else
                $collInfos = $info.Properties
            EndIf
;             For $objPropInfo In $info.Properties
            For $i = $typeDoc_optCollBase To $collInfos.Count - $typeDoc_optCollUBDecr
                $objSubinfo = $collInfos.Item($i)
                If $DESCKIND_VARDESC = $objSubinfo.DescKind Then $coll.appendChild(TypeDoc_CreatePropertyInfoNode($objSubinfo, $typeKind))
            Next
            $objSubinfo = 0
            $result.appendChild($coll)
    EndSwitch
    Switch $typeKind
        Case $TKIND_INTERFACE, $TKIND_DISPATCH, $TKIND_MODULE
            $coll = $typeDoc_objDOM.createElement("Methods")
            If IsObj($app_objTLI) Then
                $collInfos = $info.Members
            Else
                $collInfos = $info.Methods
            EndIf
;             For $objMethodInfo In $info.Methods
            For $i = $typeDoc_optCollBase To $collInfos.Count - $typeDoc_optCollUBDecr
                $objSubinfo = $collInfos.Item($i)
                If "IUnknown" <> $info.Name And "IDispatch" <> $info.Name And StringInStr($TYPEDOC_STDMEMB_IGNORE, "|" & $objSubinfo.MemberId & "|") Then ContinueLoop
                If $DESCKIND_FUNCDESC = $objSubinfo.DescKind Then $coll.appendChild(TypeDoc_CreateMethodInfoNode($objSubinfo))
            Next
            $objSubinfo = 0
            $result.appendChild($coll)
    EndSwitch
    $result.removeAttribute("tldapp-expand")

    If $feedback Then GUISetCursor(-1, 0, $frmMain)
EndFunc

Func TypeDoc_CreateTypeRefNode(Const ByRef $info)
    Local $result = $typeDoc_objDOM.createElement("TypeRef")
    $result.setAttribute("name", $info.Name)
    $result.setAttribute("guid", $info.GUID)
    $result.setAttribute("number", $info.TypeInfoNumber)
    Local $objParent = $info.Parent
    If IsObj($objParent) Then $result.setAttribute("typelib", $objParent.GUID)
    Return $result
EndFunc

Func TypeDoc_CreatePropertyInfoNode(Const ByRef $info, $parentKind)
    Local $result = $typeDoc_objDOM.createElement("Property")
    TypeDoc_InitInfoNode($result, $info, 3)
    Local $kind = $VAR_PERINSTANCE
    If IsObj($app_objTLI) Then
        If BitAnd(0x20, $info.InvokeKind) Or $TKIND_ENUM = $parentKind Then $kind = $VAR_CONST
    Else
        $kind = $info.VarKind
    EndIf
    $result.setAttribute("kind", $kind)

    Local $objVarInfo = $info.ReturnType
    $result.appendChild(TypeDoc_CreateVarInfoNode($objVarInfo))

    If $VAR_CONST = $kind Then
        Local $kid = $typeDoc_objDOM.createElement("Value")
        $kid.appendChild($typeDoc_objDOM.createTextNode("" & $info.Value))
        $result.appendChild($kid)
    EndIf

    Return $result
EndFunc

Func TypeDoc_CreateMethodInfoNode(Const ByRef $info)
    Local $result
    Local $invKind = BitAnd($info.InvokeKind, BitNot(0x30))
    Switch $invKind
        Case $INVOKE_FUNC
            $result = $typeDoc_objDOM.createElement("Function")
        Case $INVOKE_PROPERTYGET
            $result = $typeDoc_objDOM.createElement("PropertyGet")
        Case $INVOKE_PROPERTYPUT
            $result = $typeDoc_objDOM.createElement("PropertyPut")
        Case $INVOKE_PROPERTYPUTREF
            $result = $typeDoc_objDOM.createElement("PropertyPutRef")
        Case Else
            $result = $typeDoc_objDOM.createElement("Method")
    EndSwitch
    TypeDoc_InitInfoNode($result, $info, 2)
    With $info
        If Not IsObj($app_objTLI) Then $result.setAttribute("kind", .FuncKind)
        $result.setAttribute("call", .CallConv)
        $result.setAttribute("invoke", $invKind)
        $result.setAttribute("vtable", .VTableOffset)

        Local $coll = $typeDoc_objDOM.createElement("Parameters")
        Local $collInfos = .Parameters
        Local $objSubinfo, $objVarInfo, $param, $kid
        For $i = $typeDoc_optCollBase To $collInfos.Count - $typeDoc_optCollUBDecr
            $objSubinfo = $collInfos.Item($i)
            $param = $typeDoc_objDOM.createElement("Parameter")
            Local $name = $objSubinfo.Name
            If Not $name Then $name = $info.Name & "Val"
            $param.setAttribute("name", $name)
            ; .Optional property in TLBINF is arbitrary
            $param.setAttribute("optional", 0 < BitAnd($PARAMFLAG_FOPT, $objSubinfo.Flags) ? 1 : 0)

            $kid = $typeDoc_objDOM.createElement("Flags")
            $kid.setAttribute("mask", $objSubinfo.Flags)
            TypeDoc_ExplodeFlags($objSubinfo.Flags, $PARAMFLAG_FIN, $PARAMFLAG_FHASCUSTDATA, $kid)
            $param.appendChild($kid)

            $objVarInfo = $objSubinfo.VarTypeInfo
            $param.appendChild(TypeDoc_CreateVarInfoNode($objVarInfo))
            $coll.appendChild($param)

            ; .Default property in TLBINF is arbitrary
            If BitAnd($PARAMFLAG_FHASDEFAULT, $objSubinfo.Flags) Then
                Local $val = $objSubinfo.DefaultValue
;                 If $val Or IsBool($val) Or TypeDoc_IsNumericVarType($objVarInfo) Or $__Au3Obj_VT_USERDEFINED = $objVarInfo.VarType Or $__Au3Obj_VT_EMPTY = $objVarInfo.VarType Then
                $kid = $typeDoc_objDOM.createElement("Value")
                $kid.appendChild($typeDoc_objDOM.createTextNode("" & $val))
                $param.appendChild($kid)
            EndIf
        Next
        $objVarInfo = 0
        $objSubinfo = 0

        Local $objRetInfo = .ReturnType
        If IsObj($app_objTLI) Then
            ; repair crapy "VB" view in TLBINF
            Switch $invKind
                Case $INVOKE_PROPERTYPUT, $INVOKE_PROPERTYPUTREF
                    If $objRetInfo.VarType <> $__Au3Obj_VT_HRESULT And $objRetInfo.VarType <> $__Au3Obj_VT_VOID Then
                        $param = $typeDoc_objDOM.createElement("Parameter")
                        If $INVOKE_PROPERTYPUTREF = $invKind Then
                            $param.setAttribute("name", $info.Name & "Ref")
                        Else
                            $param.setAttribute("name", $info.Name & "Val")
                        EndIF
                        $param.appendChild(TypeDoc_CreateVarInfoNode($objRetInfo))
                        $coll.appendChild($param)
                        $objRetInfo = 0

                        $kid = $typeDoc_objDOM.createElement("VarType")
                        $kid.setAttribute("vt", $__Au3Obj_VT_VOID)
                        $result.appendChild($kid)
                    EndIf
            EndSwitch
        EndIf
        If IsObj($objRetInfo) Then $result.appendChild(TypeDoc_CreateVarInfoNode($objRetInfo))

        $result.appendChild($coll)
    EndWith
    Return $result
EndFunc

Func TypeDoc_CreateVarInfoNode(Const ByRef $info)
    Local $result = $typeDoc_objDOM.createElement("VarType")
    Local $vt = $info.VarType
    $result.setAttribute("vt", $vt)
    $result.setAttribute("indirection", $info.PointerLevel)
    Switch $vt
        Case $__Au3Obj_VT_USERDEFINED, $__Au3Obj_VT_EMPTY
            Local $objSubinfo = $info.TypeInfo
            If IsObj($objSubinfo) Then
                $result.setAttribute("vt", $__Au3Obj_VT_USERDEFINED)
                $result.appendChild(TypeDoc_CreateTypeRefNode($objSubinfo))
            EndIf
        Case Else
            ; VT_CARRAY, VT_SAFEARRAY is never returned by TLBINFO32
            If $__Au3Obj_VT_SAFEARRAY = $vt Or $__Au3Obj_VT_CARRAY = $vt Or BitAnd($__Au3Obj_VT_ARRAY, $vt) Or BitAnd($__Au3Obj_VT_VECTOR, $vt) Then
                Local $baseVT = BitAnd(BitNot(BitOr($__Au3Obj_VT_ARRAY, $__Au3Obj_VT_VECTOR)), $vt)
                If $baseVT <> $vt Then $result.setAttribute("vt", BitAnd(BitNot($baseVT), $vt))
                Local $kid = $typeDoc_objDOM.createElement("Array")
                If IsObj($app_objTLI) Then
                    Local $elmType = $typeDoc_objDOM.createElement("VarType")
                    $elmType.setAttribute("vt", $baseVT)
                    $elmType.setAttribute("indirection", $info.ElementPointerLevel)
                    If $baseVT = $__Au3Obj_VT_USERDEFINED Or $baseVT = $__Au3Obj_VT_EMPTY Then
                        Local $objSubinfo = $info.TypeInfo
                        If IsObj($objSubinfo) Then
                            $elmType.setAttribute("vt", $__Au3Obj_VT_USERDEFINED)
                            $elmType.appendChild(TypeDoc_CreateTypeRefNode($objSubinfo))
                        EndIf
                    EndIf
                    $kid.appendChild($elmType)
                    ; how we get array of Long here?
                    Local $bounds[1] = [0]
                    Local $dims = $info.ArrayBounds($bounds)
                    For $i = 0 To $dims - 1
                        Local $d = $typeDoc_objDOM.createElement("Dim")
                        $d.setAttribute("lbound", $bounds[$i][0])
                        $d.setAttribute("ubound", $bounds[$i][1])
                        $kid.appendChild($d)
                    Next
                Else
                    Local $elmInfo = $info.ArrayElementInfo
                    $kid.appendChild(TypeDoc_CreateVarInfoNode($elmInfo))
                    Local $dims = $info.ArrayBounds
                    If IsArray($dims) Then
                        For $i = 1 To $dims[0]
                            Local $d = $typeDoc_objDOM.createElement("Dim")
                            Local $bounds = $dims[$i]
                            $d.setAttribute("lbound", $bounds[0])
                            $d.setAttribute("ubound", $bounds[1])
                            $kid.appendChild($d)
                        Next
                    EndIf
                EndIf
                $result.appendChild($kid)
            EndIf
    EndSwitch
    Return $result
EndFunc

Func TypeDoc_ListBaseInterfaces(ByRef $info, ByRef $result)
    Local $coll = $typeDoc_objDOM.createElement("Base")
    Local $collInfos
    If IsObj($app_objTLI) Then
        $collInfos = $info.ImpliedInterfaces
    Else
        $collInfos = $info.Interfaces
    EndIf
    For $i = $typeDoc_optCollBase To $collInfos.Count - $typeDoc_optCollUBDecr
        $objSubinfo = $collInfos.Item($i)
        $coll.appendChild(TypeDoc_CreateTypeRefNode($objSubinfo))
    Next
    $objSubinfo = 0
    $result.appendChild($coll)
EndFunc

Func TypeDoc_IsNumericVarType(ByRef $varTypeInfo)
    Local $ind = 0
    If IsObj($app_objTLI) Then
        $ind = $varTypeInfo.PointerLevel
    Else
        $ind = $varTypeInfo.ElementPointerLevel
    EndIf
    Return TypeDoc_IsNumericType($varTypeInfo.VarType, $ind)
EndFunc

Func TypeDoc_IsNumericType($vt, $indirection = 0)
    If 0 < $indirection Then Return True
    Switch $vt
        Case $__Au3Obj_VT_HRESULT, $__Au3Obj_VT_PTR, $__Au3Obj_VT_INT_PTR, $__Au3Obj_VT_UINT_PTR, _
            $__Au3Obj_VT_NULL To $__Au3Obj_VT_DATE, $__Au3Obj_VT_DECIMAL To $__Au3Obj_VT_UINT
            Return True
    EndSwitch
    Return False
EndFunc

Func TypeDoc_ExplodeFlags($mask, $min, $max, ByRef $node)
    If 0 < $mask Then
        Local $attr
        Local $d = $min
        While $d <= $max
            If BitAnd($d, $mask) Then
                $attr = $typeDoc_objDOM.createElement("Flag")
                $attr.setAttribute("val", $d)
                $node.appendChild($attr)
            EndIf
            $d = BitShift($d, -1)
        WEnd
    EndIf
EndFunc
