; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Main executable for TypeLibInspector.
; $Revision: 1.4 $
; $Date: 2010/07/29 14:24:18 $
;
; ------------------------------------------------------------------------------
#NoTrayIcon
Opt("TrayIconHide", 1)

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GuiStatusBar.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <GuiMenu.au3>
#include <TLI.au3>
#include "Product.au3"
#include "Utils.au3"
#include "TypeDoc.au3"
#include "FrmMain.au3"
#include "Plugin.au3"
#include "PagingStack.au3"

Global Const $PROGID_MSXMLDOC   = "MSXML2.DOMDocument"
Global Const $OBJ_NULL  = 0

Global Const $APP_PARAM_SELTYPE     = "--selectType="

Global Const $APPMB_FATAL   = BitOr($MB_OK, $MB_ICONHAND)
Global Const $APPMB_ERROR   = BitOr($MB_OK, $MB_ICONEXCLAMATION)
Global Const $APPMB_INFO    = BitOr($MB_OK, $MB_ICONASTERISK)
Global Const $APPMB_YESNO   = BitOr($MB_YESNO, $MB_ICONQUESTION)

Global $app_tviRoot = 0
Global $app_hImlTypes = 0
Global $app_objTLI = 0
Global $app_objTypeLib = 0
Global $app_objXSLProc = 0
Global $app_objFoundNodes = 0
Global $app_iFindNext = 0
Global $app_strFind =""
Global $app_objError = ObjEvent("AutoIt.Error", "App_OnError")
Global $app_objIE = 0
Global $app_objIEEvent = 0
Global $app_history = 0
Global $app_historyCurr = -1
Global $app_historyNav = False

Global $app_optResDir = @ScriptDir & "\..\..\res\"
If @compiled Then $app_optResDir = @ScriptDir & "\"
Global $app_optXSLDir = $app_optResDir & "xsl\"
Global $app_optIconDir = $app_optResDir & "icons\"
Global $app_optUseTLBINF = 1
Global $app_optHistoryMax = 10

Main()

Func Main()
    Local $opt = RegRead("HKCU\Software\" & $VENDOR & "\" & $PRODUCT, "UseTLBINF")
    If 0 = @error Then $app_optUseTLBINF = $opt
    
    $opt = RegRead("HKCU\Software\" & $VENDOR & "\" & $PRODUCT, "DualIfaceMode")
    If 0 = @error Then $typeDoc_optDualIfaceMode = $opt
    
    $app_objIE = ObjCreate("Shell.Explorer.2")
    If Not IsObj($app_objIE) Then App_MsgBox("Failed to create HTML view.", $APPMB_FATAL)
    
    FrmMain_Create()
    If 0 < $CmdLine[0] Then _GUICtrlStatusBar_SetText($frmMain_status, "Busy...", 0)
    
    OnAutoItExitRegister("App_OnExit")

    If IsObj($app_objIE) Then $app_objIE.navigate("file:///" & $app_optResDir & "html\tldoc-view.html")

    If 1 > $app_optUseTLBINF Then GUICtrlSetState($app_cmdUseTLBINF, $GUI_UNCHECKED)
    If $TYPEDOC_DUALMODE_OLEAUT <> $typeDoc_optDualIfaceMode Then
        GUICtrlSetState($app_cmdViewDualVTable, $GUI_CHECKED)
        GUICtrlSetState($app_cmdViewDualOLEAut, $GUI_UNCHECKED)
    EndIf
    _AutoItObject_Startup()
    
    $app_objXSLProc = App_CreateXSLProcessor($app_optXSLDir & "tldoc-view.xsl")
    
    $app_hImlTypes = _GUIImageList_Create(16, 16, 5, 1)
    _GUIImageList_AddIcon($app_hImlTypes, $app_optIconDir & "any.ico")
    _GUIImageList_AddIcon($app_hImlTypes, $app_optIconDir & "TypeLib.ico")
    _GUICtrlTreeView_SetNormalImageList(_GUICtrl_GetHandle($frmMain_tvwLib), $app_hImlTypes)
    
    TLDPlugin_LoadList($frmMain_mnuFileExport, "export")
    
    If 0 < $CmdLine[0] Then
        Local $query = "", $tlbToLoad = $CmdLine[1]
        If 1 < $CmdLine[0] And $APP_PARAM_SELTYPE = StringLeft($CmdLine[2], StringLen($APP_PARAM_SELTYPE)) Then
            $query = StringRight($CmdLine[2], StringLen($CmdLine[2]) - StringLen($APP_PARAM_SELTYPE))
        EndIf
        If "-" = $tlbToLoad And "@progid=" = StringLeft($query, 8) Then
            Local $progID = StringMid($query, 10, StringLen($query) - 10)
            Local $obj = ObjCreate($progID)
            If IsObj($obj) Then
                Local $tinf = _TLI_TypeInfoFromObject($obj)
                If IsObj($tinf) Then
                    $tlbToLoad = $tinf.Parent.ContainingFile
                    $query &= " or (@guid='" & $tinf.GUID & "' and @name='" & $tinf.Name & "')"
                EndIf
            Else
                App_MsgBox($progID & " is not valid.", $APPMB_ERROR)
                $tlbToLoad = ""
            EndIf
            $obj = 0
            $tinf = 0
        EndIf
        If "-" = $tlbToLoad Then
            App_MsgBox("Some program arguments are either missing or invalid.", $APPMB_ERROR)
        ElseIf 0 < StringLen($tlbToLoad) Then
            App_LoadTypeLib($tlbToLoad)
        EndIf
        If 0 < StringLen($query) And IsObj($typeDoc_objDOM) Then
            Local $node = $typeDoc_objDOM.selectSingleNode("/TypeLib/Types/*[" & $query & "]")
            If IsObj($node) Then
                App_SelectAndDisplayNode($node.getAttribute("tldapp-ptr"))
            Else
                App_MsgBox("Type by " & $query & " is not found in this library." & @CRLF & "Some types (like OLE stock types) are inaccessible through TypeLib Information API.")
            EndIf
        EndIf
    EndIf
    FrmMain_EnableInterface()
    If IsObj($app_objIE) Then $app_objIEEvent = ObjEvent($app_objIE, "IEEvent_", "DWebBrowserEvents")
    
    EnvSet("tldresdir", $app_optResDir)
    EnvSet("tldiconsdir", $app_optIconDir)
    EnvSet("tldxsldir", $app_optXSLDir)
    While App_CheckGUIMsg()
        ;
    WEnd
EndFunc

Func App_CheckGUIMsg()
    Local $msg = GUIGetMsg(1)
    If $msg[1] <> _GUICtrl_GetHandle($frmMain_tvwLib) Then
        Switch $msg[0]
            Case $app_cmdAbout
                App_AboutBox()
            Case $app_cmdOpen
                App_LoadTypeLib()
            Case $app_cmdSave
                App_SaveTypeLibXML()
            Case $app_cmdFind
                App_FindString()
            Case $app_cmdFindNext
                App_FindNext()
            Case $app_cmdUseTLBINF
                App_ToggleTLIBackEnd()
            Case $app_cmdNavHome
                App_SelectAndDisplayNode($app_tviRoot)
            Case $app_cmdNavUp
                App_SelectAndDisplayNode(_GUICtrlTreeView_GetParentParam(_GUICtrl_GetHandle($frmMain_tvwLib), _GUICtrl_GetHandle(GUICtrlRead($frmMain_tvwLib))))
            Case $app_cmdNavBack, $app_cmdNavForward
                App_HistoryNavigate($msg[0])
            Case $app_cmdViewDualOLEAut, $app_cmdViewDualVTable
                App_ToggleDualView($msg[0])
            Case $app_cmdsSelect[0] To $app_cmdsSelect[$TKIND_MAX]
                App_ToggleSelectFilter($msg[0])
            Case $GUI_EVENT_CLOSE, $app_cmdExit
                Return False
            Case Else
                TLDPlugin_IfExec($msg)
        EndSwitch
    EndIf
    Return True
EndFunc

Func App_ToggleTLIBackEnd()
    If BitAnd($GUI_CHECKED, GUICtrlRead($app_cmdUseTLBINF)) Then
        $app_optUseTLBINF = 0
        GUICtrlSetState($app_cmdUseTLBINF, $GUI_UNCHECKED)
    Else
        $app_optUseTLBINF = 1
        GUICtrlSetState($app_cmdUseTLBINF, $GUI_CHECKED)
    EndIf
    RegWrite("HKCU\Software\" & $VENDOR & "\" & $PRODUCT, "UseTLBINF", "REG_DWORD", $app_optUseTLBINF)
    If IsObj($app_objTypeLib) Then
        If 6 = App_MsgBox("Changes will be applied the next time a TypeLib is loaded. Reload current TypeLib?", $APPMB_YESNO) Then
            Local $TLID[4] = [$app_objTypeLib.GUID, $app_objTypeLib.MajorVersion, $app_objTypeLib.MinorVersion, $app_objTypeLib.LCID]
            App_UnloadTypeLib()
            App_LoadTypeLib($TLID[0], $TLID[1], $TLID[2], $TLID[3])
        EndIf
    EndIf
EndFunc

Func App_CreateXSLProcessor($path, $respath = $app_optResDir)
    Local $result = 0
    Local $xsl = ObjCreate("MSXML2.FreeThreadedDOMDocument")
    If IsObj($xsl) Then
        $xsl.async = False
        If $xsl.load($path) Then
            Local $xslt = ObjCreate("MSXML2.XSLTemplate")
            $xslt.stylesheet = $xsl
            $result = $xslt.createProcessor()
            If 0 < StringLen($respath) Then $result.addParameter("respath", $respath)
        EndIf
        If Not IsObj($result) Then App_MsgBox("Failed to load XSL " & $path, $APPMB_ERROR)
    Else
        App_MsgBox("MSXML2 is not installed.", $APPMB_ERROR)
    EndIf
    Return $result
EndFunc

Func App_SaveXSLOutput(ByRef $xslProc, $path)
    Local $stream = ObjCreate("ADODB.Stream")
    If IsObj($stream) Then
        $stream.Open
        $stream.Type = 1
        
        $xslProc.output = $stream
        $xslProc.transform()
        $stream.SaveToFile($path, 2)
        $stream.Close()
    Else
        ConsoleWrite("WARNING: ADODB.Stream not found, encoding of " & $path & " will be UTF-16 LE" & @lf)
        $xslProc.transform()
        Local $hf = FileOpen($path, BitOr(2, 32))
        FileWrite($hf, $xslProc.output)
        FileFlush($hf)
        FileClose($hf)
    EndIf
EndFunc

Func App_MsgBox($text, $flag = $APPMB_INFO, $timeout = 0, $title = $PRODUCT_TITLE)
    ProgressOff()
    Return MsgBox($flag, $title, $text, $timeout)
EndFunc

Func App_SaveTypeLibXML($path = "")
    If Not IsObj($typeDoc_objDOM) Then Return
    If 0 = StringLen($path) Then $path = FileSaveDialog("Save XML", "", "XML Files (*.xml)|All Files (*.*)", 18, $app_objTypeLib.Name & "-" & $app_objTypeLib.MajorVersion & "." & $app_objTypeLib.MinorVersion & ".xml")
    If 0 = StringLen($path) Then Return
    App_ExpandAll()
    TypeDoc_PrettyPrint()
    $typeDoc_objDOM.save($path)
EndFunc

Func App_InitTLIBackend()
    If 0 < $app_optUseTLBINF And Not IsObj($app_objTLI) Then $app_objTLI = ObjCreate("TLI.TLIApplication")
    If IsObj($app_objTLI) Then
        $app_objTLI.ResolveAliases = False
        _GUICtrlStatusBar_SetText($frmMain_status, "Using: TLBINF32.dll", 1)
    Else
        _GUICtrlStatusBar_SetText($frmMain_status, "Using: TLI.au3", 1)
    EndIf
EndFunc

Func App_LoadTypeLib($pathOrGUID = "", $verMajor = -1, $verMinor = 0, $lcid = $LOCALE_SYSTEM_DEFAULT)
    If 0 = StringLen($pathOrGUID) Then $pathOrGUID = FileOpenDialog("Open Type Library", "", "TypeLib Files (*.tlb;*.olb;*.dll;*.ocx;*.exe)|All Files (*.*)", 1)
    If 0 = StringLen($pathOrGUID) Then Return
    App_UnloadTypeLib()
    App_InitTLIBackend()
    If "{" = StringLeft($pathOrGUID, 1) And "}" = StringRight($pathOrGUID, 1) Then
        If 0 > $verMajor Then
            Local $iTLib
            For $i = 0 To 100
                $iTLib = _ITypeLib_LoadReg($pathOrGUID, $i, $verMinor, $lcid)
                If IsObj($iTLib) Then ExitLoop
            Next
            If IsObj($iTLib) Then
                Local $tAttr = _ITypeLib_GetLibAttr($iTLib)
                $verMajor = DllStructGetData($tAttr, "wMajorVerNum")
                $verMinor = DllStructGetData($tAttr, "wMinorVerNum")
                $lcid = DllStructGetData($tAttr, "lcid")
                $iTLib = 0
            Else
                $verMajor = 1
            EndIf
        EndIf
        If IsObj($app_objTLI) Then
            $app_objTypeLib = $app_objTLI.TypeLibInfoFromRegistry($pathOrGUID, $verMajor, $verMinor, $lcid)
        Else
            $app_objTypeLib = _TLI_TypeLibInfoFromRegistry($pathOrGUID, $verMajor, $verMinor, $lcid)
        EndIf
    Else
        If IsObj($app_objTLI) Then
            $app_objTypeLib = $app_objTLI.TypeLibInfoFromFile($pathOrGUID)
        Else
            $app_objTypeLib = _TLI_TypeLibInfoFromFile($pathOrGUID)
        EndIf
    EndIf
    If Not IsObj($app_objTypeLib) Then
        App_MsgBox("Failed to load TypeLib " & $pathOrGUID, $APPMB_ERROR)
        Return
    EndIf
    App_ParseTypeLib()
EndFunc

Func App_ParseTypeLib()
    If Not IsObj($app_objTypeLib) Then Return

    $app_tviRoot = GUICtrlCreateTreeViewItem($app_objTypeLib.Name & " (parsing...)", $frmMain_tvwLib)
    Local $htvwLib = _GUICtrl_GetHandle($frmMain_tvwLib)
    _GUICtrlTreeView_SetImageIndex($htvwLib, GUICtrlGetHandle($app_tviRoot), 1)
    _GUICtrlTreeView_SetSelectedImageIndex($htvwLib, GUICtrlGetHandle($app_tviRoot), 1)
    Local $success = TypeDoc_Create($PROGID_MSXMLDOC)
    If Not $success Then
        App_MsgBox("Failed to initialize " & $PROGID_MSXMLDOC, $APPMB_ERROR)
        Return
    EndIf
    
    ProgressOn("Parsing TypeLib", $app_objTypeLib.Name)
    GUISetCursor(15, 1, $frmMain)
    
    Local $nodePtr = $app_tviRoot
    Local $node = TypeDoc_Init($app_objTypeLib, True)

;     $node.setAttribute("path", $path)
    $node.setAttribute("tldapp-ptr", $nodePtr)
    _GUICtrlTreeView_SetItemParam($htvwLib, GUICtrlGetHandle($app_tviRoot), $nodePtr)
    
    TypeDoc_ParseTypeLibInfo(True)
    
    ProgressSet(100, "")
    
    For $typeKind = 0 To $TKIND_MAX - 1
        Local $nl = $typeDoc_objDOM.selectNodes("/TypeLib/Types/*[@kind='" & $typeKind & "']")
        If IsObj($nl) And 0 < $nl.length() Then
            For $i = 0 To $nl.length() - 1
                $node = $nl.item($i)
                Local $tvi = App_InsertNode($app_tviRoot, $node, (0 = $typeKind And 0 = $i))
                If $typeDoc_optLazyExpand Then
                    Local $tviDummy = GUICtrlCreateTreeViewItem("parsing...", $tvi)
                    _GUICtrlTreeView_SetItemParam($htvwLib, GUICtrlGetHandle($tviDummy), -1)
                Else
                    App_MakeChildren($nodePtr, $node)
                EndIf
            Next
        EndIf
    Next
    GUICtrlSetData($app_tviRoot, $app_objTypeLib.Name)
    _GUICtrlStatusBar_SetText($frmMain_status, "TypeLib " & $app_objTypeLib.Name, 0)
    GUICtrlSetState($app_tviRoot, BitOr($GUI_EXPAND, $GUI_CHECKED, $GUI_FOCUS))

    GUISetCursor(-1, 0, $frmMain)
    ProgressOff()
    GUICtrlSetState($app_cmdSave, $GUI_ENABLE)
    GUICtrlSetState($app_cmdFind, $GUI_ENABLE)
    GUICtrlSetState($app_cmdFindNext, $GUI_ENABLE)
    GUICtrlSetState($frmMain_mnuFileExport, $GUI_ENABLE)
    GUICtrlSetState($frmMain_mnuEditSelect, $GUI_ENABLE)
    WinSetTitle($frmMain, "", $app_objTypeLib.Name & " - " & $PRODUCT_TITLE)
EndFunc

Func App_UnloadTypeLib($final = False)
    If Not $final Then
        GuiCtrlSetData($frmMain_txtInfo, "")
        GUICtrlSetState($frmMain_mnuFileExport, $GUI_DISABLE)
        GUICtrlSetState($frmMain_mnuEditSelect, $GUI_DISABLE)
        GUICtrlSetState($app_cmdSave, $GUI_DISABLE)
        GUICtrlSetState($app_cmdFind, $GUI_DISABLE)
        GUICtrlSetState($app_cmdFindNext, $GUI_DISABLE)
        If $app_tviRoot Then GUICtrlDelete($app_tviRoot)
        $app_history = 0
        $app_historyCurr = -1
    EndIf
    $app_objFoundNodes = 0
    TypeDoc_Destroy()
    $app_objTypeLib = 0
    If 1 > $app_optUseTLBINF Then $app_objTLI = 0
    If Not $final Then
        $app_objIE.navigate("file:///" & $app_optResDir & "html\tldoc-view.html")
        _GUICtrlStatusBar_SetText($frmMain_status, "Ready", 0)
    EndIf
EndFunc

Func App_GetNodeIconIndex(ByRef $node)
    Local $name = $node.nodeName
    Switch $node.nodeName
        Case "TypeRef"
            $name = $node.parentNode.nodeName
        Case "Property"
            If $VAR_CONST = Number($node.getAttribute("kind")) Then
                $name = "Const"
            ElseIf BitAnd($VARFLAG_FREADONLY, $node.selectSingleNode("Attributes").getAttribute("mask")) Then
                $name &= "RO" 
            EndIf
    EndSwitch
    Return App_GetNamedIconIndex($name)
EndFunc

Func App_GetNamedIconIndex($name)
    Local $varName = "app_imlInd" & $name
    If IsDeclared($varName) Then Return Eval($varName)
    
    Local $result = _GUIImageList_AddIcon($app_hImlTypes, $app_optIconDir & $name & ".ico")
    If @error Or 0 > $result Then $result = 0
    If 1 < $result Then Assign($varName, $result, 2)
    Return $result
EndFunc

Func App_FindString()
    If Not IsObj($typeDoc_objDOM) Then Return
    $app_strFind = InputBox($PRODUCT_TITLE, "Find type or member" & @lf & "(*text* or text - find partial string, text* - find leading string, *text - find trailing string, ""text"" - find string exact):", $app_strFind)
    If "*" = $app_strFind Then $app_strFind = ""
    If 0 = StringLen($app_strFind) Then Return

    App_ExpandAll(True)
    Local $query = "//*[" & TypeDoc_MakeQuery($app_strFind, "@name,@progid,@guid,Help") & "]"
;     Local $query = "//*[contains(@name, '" & $text & "')]"
    ConsoleWrite("Searching " & $query & @lf)
    $app_iFindNext = 0
    $app_objFoundNodes = $typeDoc_objDOM.selectNodes($query)
    If IsObj($app_objFoundNodes) And 0 < $app_objFoundNodes.length() Then
        App_FindNext()
    Else
        App_MsgBox("'" & $app_strFind & "' not found.")
        $app_iFindNext = 0
;         $app_objFoundNodes = 0
;         $app_strFind = ""
    EndIf
EndFunc

Func App_FindNext()
    If Not IsObj($app_objFoundNodes) Then
        App_FindString()
        Return
    EndIf
    Local $nodePtr = 0
    Local $node
    For $i = $app_iFindNext To $app_objFoundNodes.length() - 1
        $node = $app_objFoundNodes.item($i)
        If "TypeRef" = $node.nodeName Then ContinueLoop
        $nodePtr = Number($node.getAttribute("tldapp-ptr"))
        If 0 = $nodePtr Then $node = $node.selectSingleNode("../..")
        $nodePtr = Number($node.getAttribute("tldapp-ptr"))
        If $nodePtr = GUICtrlRead($frmMain_tvwLib) Then ContinueLoop
        If 0 <> $nodePtr Then ExitLoop
    Next
    $app_iFindNext =  $i + 1
    If 0 <> $nodePtr Then
        App_SelectAndDisplayNode($nodePtr)
    Else
        App_MsgBox("No more occurences of '" & $app_strFind & "' not found.")
        $app_iFindNext = 0
;         $app_objFoundNodes = 0
;         $app_strFind = ""
    EndIf
EndFunc

Func App_HistoryNavigate($cmdID)
    If 0 < _PagingStack_GetSize($app_history) Then
        If $app_cmdNavForward = $cmdID Then
            $app_historyCurr += 1
        ElseIf $app_cmdNavBack = $cmdID Then
            $app_historyCurr -= 1
        ElseIf  1000 <= $cmdID Then
            $app_historyCurr = _GUICtrlMenu_GetItemData(_GUICtrl_GetHandle($frmMain_mnuViewNav), $cmdID, False) - 1
        EndIf
        GUICtrlSetState($app_cmdNavForward, $GUI_DISABLE)
        GUICtrlSetState($app_cmdNavBack, $GUI_DISABLE)
        If -1 < $app_historyCurr And $app_historyCurr < _PagingStack_GetSize($app_history) Then
            $app_historyNav = True
            App_SelectAndDisplayNode(_PagingStack_Get($app_history, $app_historyCurr))
            $app_historyNav = False
            If 0 < $app_historyCurr Then GUICtrlSetState($app_cmdNavBack, $GUI_ENABLE)
            If _PagingStack_GetSize($app_history) - 1 > $app_historyCurr Then GUICtrlSetState($app_cmdNavForward, $GUI_ENABLE)
            Return True
        Else
            $app_historyCurr = -1
        EndIf
    EndIf
    Return False
EndFunc

Func App_HistoryAdd($nodePtr)
    If Not $app_historyNav Then
        If IsArray($app_history) Then
            If $nodePtr = _PagingStack_GetTail($app_history) Then Return
        Else
            $app_history = _PagingStack_Create($app_optHistoryMax, True)
        EndIf
        _PagingStack_Push($app_history, $nodePtr)
        $app_historyCurr = _PagingStack_GetSize($app_history) - 1
        If 1 < _PagingStack_GetSize($app_history) Then GUICtrlSetState($app_cmdNavBack, $GUI_ENABLE)
    EndIf
EndFunc

Func App_ToggleDualView($cmdID)
    If ($TYPEDOC_DUALMODE_OLEAUT = $typeDoc_optDualIfaceMode And BitAnd($GUI_CHECKED, GUICtrlRead($app_cmdViewDualOLEAut))) _
        Or ($TYPEDOC_DUALMODE_VTABLE = $typeDoc_optDualIfaceMode And BitAnd($GUI_CHECKED, GUICtrlRead($app_cmdViewDualVTable))) Then Return
    
    If BitAnd($GUI_CHECKED, GUICtrlRead($app_cmdViewDualVTable)) Then
        $typeDoc_optDualIfaceMode = $TYPEDOC_DUALMODE_VTABLE
    Else
        $typeDoc_optDualIfaceMode = $TYPEDOC_DUALMODE_OLEAUT
    EndIf
    RegWrite("HKCU\Software\" & $VENDOR & "\" & $PRODUCT, "DualIfaceMode", "REG_DWORD", $typeDoc_optDualIfaceMode)
    Local $htvw = _GUICtrl_GetHandle($frmMain_tvwLib)
    Local $nodePtr = 0
    Local $nl = $typeDoc_objDOM.selectNodes("/TypeLib/Types/DispInterface[Attributes/Flag/@val='" & Number($TYPEFLAG_FDUAL) & "']")
    If IsObj($nl) And 0 < $nl.length() Then
        For $i = 0 To $nl.length() - 1
            $node = $nl.item($i)
            If "1" <> $node.getAttribute("tldapp-expand") Then
                $nodePtr = Number($node.getAttribute("tldapp-ptr"))
                If $nodePtr Then _GUICtrlTreeView_DeleteChildren($htvw, _GUICtrl_GetHandle($nodePtr))
                TypeDoc_UnexpandTypeInfoNode($node)
                Local $tviDummy = GUICtrlCreateTreeViewItem("parsing...", $nodePtr)
                _GUICtrlTreeView_SetItemParam($htvw, GUICtrlGetHandle($tviDummy), -1)
            EndIf
        Next
    EndIf
    App_SelectAndDisplayNode($app_tviRoot)
EndFunc

Func App_ToggleSelectFilter($filterID)
    Local $check = $GUI_UNCHECKED
    If 0 = BitAnd($GUI_CHECKED, GUICtrlRead($filterID)) Then $check = $GUI_CHECKED
    GUICtrlSetState($filterID, $check)
    Local $filterIdx = $filterID - $app_cmdsSelect[0]
    If $GUI_UNCHECKED = $check And $TKIND_MAX > $filterIdx Then GUICtrlSetState($app_cmdsSelect[$TKIND_MAX], $check)
    If Not IsObj($typeDoc_objDOM) Then Return
    
    Local $filter = "/TypeLib/Types/"
    If $TKIND_MAX <= $filterIdx Then
        $filter &= "*"
        For $i = 0 To $TKIND_MAX - 1
            GUICtrlSetState($app_cmdsSelect[$i], $check)
        Next
    Else
        $filter &= $TYPEDOC_TKIND[$filterIdx]
    EndIf
    Local $nl = $typeDoc_objDOM.selectNodes($filter)
    If IsObj($nl) And 0 < $nl.length() Then
        Local $nodePtr = 0
        For $i = 0 To $nl.length() - 1
            $nodePtr = Number($nl.item($i).getAttribute("tldapp-ptr"))
            If $nodePtr Then
                GUICtrlSetState($nodePtr, $check)
                TVWLib_CheckChildren(_GUICtrl_GetHandle($nodePtr), ($GUI_CHECKED = $check))
            EndIf
        Next
    EndIf
EndFunc

Func App_SelectAndDisplayNode($nodePtr)
    Local $render = GUICtrlRead($frmMain_tvwLib) = $nodePtr
    Local $iCheck = _GUICtrlTreeView_GetStateImageIndex(_GUICtrl_GetHandle($frmMain_tvwLib), _GUICtrl_GetHandle($nodePtr))
    GUICtrlSetState($nodePtr, BitOr(GUICtrlRead($nodePtr), $GUI_FOCUS))
    If 1 > $iCheck Then _GUICtrlTreeView_SetStateImageIndexMod(_GUICtrl_GetHandle($frmMain_tvwLib), _GUICtrl_GetHandle($nodePtr), 0)
    If $render Then App_DisplayNode($nodePtr)
EndFunc

Func App_DisplayNode($nodePtr)
    Local $node = $typeDoc_objDOM.selectSingleNode("//*[@tldapp-ptr='" & $nodePtr & "']")
    If Not IsObj($node) Then
        _GUICtrlStatusBar_SetText($frmMain_status, "[Invalid selection]", 0)
        GuiCtrlSetData($frmMain_txtInfo, "")
        $app_objIE.document.body.innerHTML = ""
        Return
    EndIf
    If "TypeRef" = $node.nodeName Then
        Local $orig = $typeDoc_objDOM.selectSingleNode("/TypeLib/Types/*[@guid='" & $node.getAttribute("guid") & "' and @name='" & $node.getAttribute("name") & "']")
        If IsObj($orig) Then
            $nodePtr = Number($orig.getAttribute("tldapp-ptr"))
            _WinAPI_PostMessage($frmMain, $TVM_REDIRECT, 0, $nodePtr)
            Return
        EndIf
        App_DisplayExternal($node.getAttribute("typelib"), $node.getAttribute("name"))
    EndIf
    App_HistoryAdd($nodePtr)
    App_ExpandNode($nodePtr)
    _GUICtrlStatusBar_SetText($frmMain_status, $node.nodeName & " " & $node.getAttribute("name"), 0)
    GuiCtrlSetData($frmMain_txtInfo, StringReplace($node.XML, @tab, "    "))
    
    $app_objXSLProc.input = $typeDoc_objDOM
    $app_objXSLProc.addParameter("ptr", $nodePtr)
    $app_objXSLProc.transform()
    
    $app_objIE.document.body.innerHTML = "" & $app_objXSLProc.output
    ;ConsoleWrite("XSLT: " & $app_objXSLProc.output & @lf)
;     $dummy = 0
EndFunc

Func App_DisplayExternal($typelib, $name, $guid = "", $num = -1)
    If 6 = App_MsgBox("This type references external library, would you like to open it?", $APPMB_YESNO) Then
        If @compiled Then
            ShellExecute(@ScriptFullPath, $typelib & " " & $APP_PARAM_SELTYPE & "@name='" & $name & "'")
        Else
            ShellExecute(@AutoItExe, """" & @ScriptFullPath & """ " & $typelib & " " & $APP_PARAM_SELTYPE & "@name='" & $name & "'")
        EndIf
    EndIf
EndFunc

Func App_ExpandAll($withPtr = False)
    If $typeDoc_optLazyExpand Then
        Local $nl = $typeDoc_objDOM.selectNodes("//*[@tldapp-expand='1']")
        If IsObj($nl) And 0 < $nl.length() Then
            Local $nodePtr = GUICtrlRead($frmMain_tvwLib)
            
            Local $start = TimerInit()
            ProgressOn("Expanding", $app_objTypeLib.Name)
            GUISetCursor(15, 1, $frmMain)
            Local $c = $nl.length()
            For $i = 0 To $c - 1
                Local $node = $nl.item($i)
                ProgressSet($i * 100 / $c, $node.getAttribute("name"))
                If $withPtr Then
                    App_ExpandNode($node.getAttribute("tldapp-ptr"))
                Else
                    TypeDoc_ExpandTypeInfoNode($OBJ_NULL, $node)
                EndIf
            Next
            ConsoleWrite("---------- Load time: " & TimerDiff($start) & @lf)
            GUISetCursor(-1, 0, $frmMain)
            ProgressOff()
            If 0 < $nodePtr Then
                Local $node = $typeDoc_objDOM.selectSingleNode("//*[@tldapp-ptr='" & $nodePtr & "']")
                If IsObj($node) Then
                    _GUICtrlStatusBar_SetText($frmMain_status, $node.nodeName & " " & $node.getAttribute("name"), 0)
                Else
                    $nodePtr = 0
                EndIf
            EndIf
            If 0 >= $nodePtr Then _GUICtrlStatusBar_SetText($frmMain_status, "Ready", 0)
        EndIf
    EndIf
EndFunc

Func App_ExpandNode($nodePtr)
    Local $htvwLib = _GUICtrl_GetHandle($frmMain_tvwLib)
    Local $htiOld = _GUICtrlTreeView_GetFirstChild($htvwLib, GUICtrlGetHandle($nodePtr))
    If -1 < _GUICtrlTreeView_GetItemParam($htvwLib, $htiOld) Then Return
    Local $node = $typeDoc_objDOM.selectSingleNode("//*[@tldapp-ptr='" & $nodePtr & "']")
    If Not IsObj($node) Then
        _GUICtrlTreeView_SetText($htvwLib, $htiOld, "failed")
        Return
    EndIf
    TypeDoc_ExpandTypeInfoNode($OBJ_NULL, $node, True)
    App_MakeChildren($nodePtr, $node)
    _GUICtrlTreeView_Delete($htvwLib, $htiOld)
EndFunc

Func App_MakeChildren($tvi, ByRef $node)
    App_InsertChildren($tvi, "Resolved/VarType/TypeRef", $node)
    App_InsertChildren($tvi, "VTable/*", $node)
    App_InsertChildren($tvi, "Base/*", $node)
    App_InsertChildren($tvi, "Interfaces/Impl/TypeRef", $node)
    App_InsertChildren($tvi, "Properties/*", $node)
    App_InsertChildren($tvi, "Methods/*[@invoke='" & $INVOKE_PROPERTYGET & "' or @invoke='" & $INVOKE_PROPERTYPUT & "' or @invoke='" & $INVOKE_PROPERTYPUTREF & "']", $node)
    App_InsertChildren($tvi, "Methods/*[@invoke='" & $INVOKE_FUNC & "']", $node)
EndFunc

Func App_InsertChildren($tviParent, $childType, ByRef $parent)
    Local $nl = $parent.selectNodes($childType)
    If IsObj($nl) And 0 < $nl.length() Then
        For $i = 0 To $nl.length() - 1
            Local $node = $nl.item($i)
            App_InsertNode($tviParent, $node)
        Next
    EndIf
EndFunc

Func App_InsertNode($tviParent, ByRef $node, $expand = False)
    Local $name = $node.getAttribute("name")
    If "Method" = $node.nodeName Or "Function" = $node.nodeName Then $name &= "()"
    Local $htvw = _GUICtrl_GetHandle($frmMain_tvwLib)
    
    Local $tvi = GUICtrlCreateTreeViewItem($name, $tviParent)
    Local $hti = GUICtrlGetHandle($tvi)
    Local $iType = App_GetNodeIconIndex($node)
    _GUICtrlTreeView_SetImageIndex($htvw, $hti, $iType)
    _GUICtrlTreeView_SetSelectedImageIndex($htvw, $hti, $iType)
    
    Local $nodePtr = Number($tvi)
    $node.setAttribute("tldapp-ptr", "" & $nodePtr)
    _GUICtrlTreeView_SetItemParam($htvw, $hti, $nodePtr)
    Local $state = $GUI_CHECKED
    If 0 < StringLen($node.getAttribute("dispid")) Then
        Local $f = $node.selectSingleNode("Attributes/Flag[@val='" & Number($FUNCFLAG_FDEFAULTBIND) & "']")
        If IsObj($f) Then $state = BitOr($state, $GUI_DEFBUTTON)
    EndIf
    GUICtrlSetState($tvi, $state)
    If 0 = StringLen($node.getAttribute("guid")) Or "TypeRef" = $node.nodeName Then _GUICtrlTreeView_SetStateImageIndexMod($htvw, $hti, 0)
    If $expand Then GUICtrlSetState($tviParent, BitOr(GUICtrlRead($tviParent), $GUI_EXPAND))
    Return $tvi
EndFunc

Func App_AboutBox()
    DlgAbout_Create()
EndFunc

Func App_OnError()
    ConsoleWrite("! COM Error !  Number: 0x" & Hex($app_objError.number, 8) & " Source: " & $app_objError.source & "   ScriptLine: " & $app_objError.scriptline & " - " & $app_objError.windescription & @CRLF)
EndFunc

Func App_OnExit()
    _GUIImageList_Destroy($app_hImlTypes)
    If IsObj($app_objIEEvent) Then $app_objIEEvent.Stop
    $app_objIEEvent = 0
    $app_objXSLProc = 0
    App_UnloadTypeLib(True)
    _AutoItObject_Shutdown()
    $app_objError = 0
    FrmMain_Destroy()
EndFunc

Func IEEvent_DownloadComplete()
    If IsObj($typeDoc_objDOM) And IsObj($app_objFoundNodes) And 0 < StringLen($app_strFind) Then
        While Not (String($app_objIE.readyState) = "complete" Or $app_objIE.readyState = 4)
            Sleep(100)
        WEnd

        Local $text = StringRegExpReplace(StringRegExpReplace($app_strFind, "^\*", ""), "\*$", "")
        Local $rng = $app_objIE.document.body.createTextRange()
        Local $nextFnd = $rng.findText($text)
        While $nextFnd
            ;$rng.select()
            $rng.execCommand("BackColor", False, "#FFFF99")
            $rng = $rng.duplicate()
            $rng.collapse(false)
            $nextFnd = $rng.findText($text, 10000)
        WEnd
    EndIf
EndFunc

Func IEEvent_NavigateComplete($url)  
    If Not IsObj($typeDoc_objDOM) Or Not IsObj($app_objTypeLib) Then Return
    
    Local $nodePtr = 0
    Local $params = _URL_GetParameters($url)
    If 3 < $params[0] And "showNode" = $params[2] And 0 < StringLen($params[4]) Then
        $nodePtr = Number($params[4])
    ElseIf 5 < $params[0] And "showExternal" = $params[2] And 0 < StringLen($params[4]) And 0 < StringLen($params[6]) Then
        App_DisplayExternal($params[4], $params[6])
    EndIf
    If 0 = $nodePtr  Then $nodePtr = GUICtrlRead($frmMain_tvwLib)
    App_SelectAndDisplayNode($nodePtr)
EndFunc

Func IEEvent_BeforeNavigate($url, $Flags, $TargetFrameName, $PostData, $Headers, $Cancel)
    
    Local $params = _URL_GetParameters($url)
    If 3 < $params[0] And "showHelp" = $params[2] And 0 < StringLen($params[4]) Then
        $params[4] = StringStripWS($params[4], 3)
        Local $helpExe = "winhlp32.exe"
        Local $helpParams = ""
        If ".chm" = StringLower(StringRight($params[4], 4)) Then
            $helpExe = "hh.exe"
            $helpParams = "ms-its:"
            If 5 < $params[0] And 0 <> Number($params[6]) Then $helpParams = "-mapid " & $params[6] & " " & $helpParams
        Else
            $params[4] = StringReplace($params[4], "/", "\")
            If 5 < $params[0] And 0 <> Number($params[6]) Then $helpParams = "-n" & $params[6] & " "
        EndIf
        ;ConsoleWrite("Executing: " & $helpExe & " " & $helpParams & $params[4] & @lf)
        ShellExecute($helpExe, $helpParams & $params[4])
    EndIf
EndFunc
