#include-once
; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Main form and info dialog for TypeLibInspector.
; $Revision: 1.3 $
; $Date: 2010/07/29 14:28:47 $
;
; ------------------------------------------------------------------------------

Global Const $TVN_ITEMCHECKED = $WM_USER + 0x100
Global Const $TVM_REDIRECT = $WM_USER + 0x101
Global Const $WM_CLOSECHILD = $WM_USER + 0x102

Func FrmMain_Create()
    Global $dlgAbout = 0
#Region ### START Koda GUI section ### Form=FrmMain.kxf
    Global $frmMain = GUICreate("Type Library Inspector", 654, 515, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_MAXIMIZEBOX,$WS_SIZEBOX,$WS_THICKFRAME,$WS_CLIPCHILDREN,$WS_TABSTOP))
    Global $frmMain_mnuFile = GUICtrlCreateMenu("&File")
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $app_cmdOpen = GUICtrlCreateMenuItem("&Open..."&@TAB&"Ctrl+O", $frmMain_mnuFile)
    Global $app_cmdSave = GUICtrlCreateMenuItem("&Save"&@TAB&"Ctrl+S", $frmMain_mnuFile)
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $frmMain_mnuFileExport = GUICtrlCreateMenu("E&xport", $frmMain_mnuFile)
    GUICtrlSetState(-1, $GUI_DISABLE)
    GUICtrlCreateMenuItem("", $frmMain_mnuFile)
    Global $app_cmdExit = GUICtrlCreateMenuItem("E&xit", $frmMain_mnuFile)
    Global $frmMain_mnuView = GUICtrlCreateMenu("&View")
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $frmMain_mnuViewDual = GUICtrlCreateMenu("&Dual Interfaces In", $frmMain_mnuView)
    Global $app_cmdViewDualOLEAut = GUICtrlCreateMenuItem("&OLE Automation Mode", $frmMain_mnuViewDual, 0, 1)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Global $app_cmdViewDualVTable = GUICtrlCreateMenuItem("&VTable Mode", $frmMain_mnuViewDual, 1, 1)
    Global $frmMain_mnuViewNav = GUICtrlCreateMenu("&Navigate", $frmMain_mnuView)
    Global $app_cmdNavHome = GUICtrlCreateMenuItem("&Home"&@TAB&"Alt+Home", $frmMain_mnuViewNav)
    Global $app_cmdNavUp = GUICtrlCreateMenuItem("&Up"&@TAB&"Alt+Up Arrow", $frmMain_mnuViewNav)
    Global $app_cmdNavBack = GUICtrlCreateMenuItem("&Back"&@TAB&"Alt+Left Arrow", $frmMain_mnuViewNav)
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $app_cmdNavForward = GUICtrlCreateMenuItem("&Forward"&@TAB&"Alt+Right Arrow", $frmMain_mnuViewNav)
    GUICtrlSetState(-1, $GUI_DISABLE)
    GUICtrlCreateMenuItem("", $frmMain_mnuViewNav)
    Global $frmMain_mnuEdit = GUICtrlCreateMenu("&Edit")
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $app_cmdFind = GUICtrlCreateMenuItem("&Find"&@TAB&"Ctrl+F", $frmMain_mnuEdit)
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $app_cmdFindNext = GUICtrlCreateMenuItem("Find &Next"&@TAB&"F3", $frmMain_mnuEdit)
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $frmMain_mnuTools = GUICtrlCreateMenu("&Tools")
    GUICtrlSetState(-1, $GUI_DISABLE)
    Global $app_cmdUseTLBINF = GUICtrlCreateMenuItem("&Use TLBINF32 (if available)", $frmMain_mnuTools)
    GUICtrlSetState(-1, $GUI_CHECKED)
    Global $frmMain_mnuHelp = GUICtrlCreateMenu("&?")
    Global $app_cmdAbout = GUICtrlCreateMenuItem("About...", $frmMain_mnuHelp)
    Global $frmMain_tvwLib = GUICtrlCreateTreeView(8, 8, 249, 457, BitOR($GUI_SS_DEFAULT_TREEVIEW,$WS_CLIPSIBLINGS,$TVS_CHECKBOXES), $WS_EX_STATICEDGE);
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM+$GUI_DOCKWIDTH)
    Global $frmMain_htmInfo = GUICtrlCreateObj($app_objIE, 264, 8, 385, 458);321)
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
    Global $frmMain_txtInfo = GUICtrlCreateEdit("", 264, 336, 385, 129, BitOR($GUI_SS_DEFAULT_EDIT,$ES_READONLY,$WS_CLIPSIBLINGS), $WS_EX_STATICEDGE)
    GUICtrlSetState(-1, $GUI_HIDE)
    GUICtrlSetData(-1, "")
    GUICtrlSetFont(-1, 10, 400, 0, "Courier New")
    GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
    Global $frmMain_status = _GUICtrlStatusBar_Create($frmMain, -1, "", BitOR($SBARS_SIZEGRIP,$SBARS_TOOLTIPS,$WS_VISIBLE,$WS_CHILD))
    Global $frmMain_status_PartsWidth[2] = [300, -1]
    _GUICtrlStatusBar_SetParts($frmMain_status, $frmMain_status_PartsWidth)
    _GUICtrlStatusBar_SetText($frmMain_status, "Ready", 0)
    _GUICtrlStatusBar_SetMinHeight($frmMain_status, 22)
    GUISetState(@SW_SHOW)
    Global $frmMain_AccelTable[8][2] = [["^o", $app_cmdOpen],["^s", $app_cmdSave],["^f", $app_cmdFind],["{F3}", $app_cmdFindNext] _
        , ["!{HOME}", $app_cmdNavHome], ["!{UP}", $app_cmdNavUp], ["!{LEFT}", $app_cmdNavBack], ["!{RIGHT}", $app_cmdNavForward]]
    GUISetAccelerators($frmMain_AccelTable)
#EndRegion ### END Koda GUI section ###

    GuiSetIcon($app_optIconDir & $PRODUCT & ".ico")
    
    Global $app_cmdsSelect[$TKIND_MAX + 1]
    GUICtrlCreateMenuItem("", $frmMain_mnuEdit)
    Global $frmMain_mnuEditSelect = GUICtrlCreateMenu("&Select By Type", $frmMain_mnuEdit)
    GUICtrlSetState(-1, $GUI_DISABLE)
    For $i = 0 To $TKIND_MAX - 1
        $app_cmdsSelect[$i] = GUICtrlCreateMenuItem($TYPEDOC_TKIND[$i], $frmMain_mnuEditSelect)
        GUICtrlSetState(-1, $GUI_CHECKED)
    Next
    $app_cmdsSelect[$TKIND_MAX] = GUICtrlCreateMenuItem("All/None", $frmMain_mnuEditSelect)
    GUICtrlSetState(-1, $GUI_CHECKED)
    
    GUIRegisterMsg($WM_CLOSECHILD, "FrmMain_On_WM_CLOSECHILD")
    GUIRegisterMsg($WM_SIZE, "FrmMain_On_WM_SIZE")
    GUIRegisterMsg($WM_NOTIFY, "FrmMain_On_WM_NOTIFY")
    GUIRegisterMsg($WM_COMMAND, "FrmMain_On_WM_COMMAND")
    GUIRegisterMsg($WM_INITMENUPOPUP, "FrmMain_On_WM_INITMENUPOPUP")
    GUIRegisterMsg($TVN_ITEMCHECKED, "TVWLib_On_TVN_ITEMCHECKED")
    GUIRegisterMsg($TVM_REDIRECT, "TVWLib_On_TVM_REDIRECT")
EndFunc

Func FrmMain_Destroy()
    DlgAbout_Destroy(True)
EndFunc

Func FrmMain_EnableInterface($enable = True)
    Local $state = $GUI_ENABLE
    If Not $enable Then $state = $GUI_DISABLE
    GUICtrlSetState($frmMain_mnuFile, $state)
    GUICtrlSetState($frmMain_mnuEdit, $state)
    GUICtrlSetState($frmMain_mnuTools, $state)
    GUICtrlSetState($frmMain_mnuView, $state)
EndFunc

Func FrmMain_On_WM_CLOSECHILD($hWnd, $msg, $wParam, $lParam)
    If $dlgAbout = $wParam Then
        DlgAbout_Destroy()
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func FrmMain_On_WM_SIZE($hWnd, $msg, $wParam, $lParam)
    _GUICtrlStatusBar_Resize($frmMain_status)
    Return $GUI_RUNDEFMSG
EndFunc

Func FrmMain_On_WM_INITMENUPOPUP($hWnd, $msg, $wParam, $lParam)
    If $wParam = _GUICtrl_GetHandle($frmMain_mnuViewNav) Then
        Local $hMenu = $wParam
        Local $c = _GUICtrlMenu_GetItemCount($hMenu)
        Local $cmd, $nodePtr, $node
        Local $iCurr = -1
        For $i = 0 To _PagingStack_GetSize($app_history) - 1
            $nodePtr = _PagingStack_Get($app_history, $i)
            $node = $typeDoc_objDOM.selectSingleNode("//*[@tldapp-ptr='" & $nodePtr & "']")
            If IsObj($node) Then
                If $i > $c - 6 Then 
                    $cmd = GUICtrlCreateMenuItem($node.nodeName & " " & $node.getAttribute("name"), $frmMain_mnuViewNav)
                Else
                    _GUICtrlMenu_SetItemText($hMenu, 5 + $i, $node.nodeName & " " & $node.getAttribute("name"))
                EndIf
                _GUICtrlMenu_SetItemData($hMenu, 5 + $i, $i + 1)
                If $i = $app_historyCurr Then $iCurr = 5 + $i
                _GUICtrlMenu_SetItemID($hMenu, 5 + $i, $i + 1000)
            EndIf
        Next
        _GUICtrlMenu_SetMenuDefaultItem($hMenu, $iCurr)
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func FrmMain_On_WM_COMMAND($hWnd, $msg, $wParam, $lParam)
    If 0 = $lParam Then
        Local $nID = _WinAPI_LoWord($wParam)
        If 1000 <= $nID And 1010 > $nID Then
            App_HistoryNavigate($nID)
            Return 0
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func FrmMain_On_WM_NOTIFY($hWnd, $msg, $wParam, $lParam)
    Local $nmhdr = DllStructCreate($tagNMHDR, $lParam)
    Local $hWndFrom = HWnd(DllStructGetData($nmhdr, "hWndFrom"))
    Local $iCode = DllStructGetData($nmhdr, "Code")
    If _GUICtrl_GetHandle($frmMain_tvwLib) = $hWndFrom Then
        Switch $iCode
            Case $NM_CLICK
                Local $pos = _GUICtrl_GetMessagePos($hWndFrom)
                Local $tHitTest = _GUICtrlTreeView_HitTestEx($hWndFrom, $pos[0], $pos[1])
                If BitAnd($TVHT_ONITEMSTATEICON, DllStructGetData($tHitTest, "Flags")) Then
                    Local $hti = DllStructGetData($tHitTest, "Item")
                    _WinAPI_PostMessage($frmMain, $TVN_ITEMCHECKED, Not _GUICtrlTreeView_GetChecked($hWndFrom, $hti), DllStructGetData($tHitTest, "Item"))
                    Return 0
                EndIf
            Case $TVN_KEYDOWN
                Local $nmtvkd = DllStructCreate($tagNMTVKEYDOWN, $lParam)
                If 32 = DllStructGetData($nmtvkd, "VKey") Then
                    Local $hti = _GUICtrlTreeView_GetSelection($hWndFrom)
                    Local $check = Number(Not _GUICtrlTreeView_GetChecked($hWndFrom, $hti))
                    If 0 = _GUICtrlTreeView_GetStateImageIndex($hWndFrom, $hti) Then $check = 3
                    If $hti Then _WinAPI_PostMessage($frmMain, $TVN_ITEMCHECKED, $check, $hti)
                    Return 0
                EndIf
            Case $TVN_SELCHANGEDA, $TVN_SELCHANGEDW
                Local $nmtv = DllStructCreate($tagNMTREEVIEW, $lParam)
                If DllStructGetData($nmtv, "NewhItem") And DllStructGetData($nmtv, "NewhItem") <> DllStructGetData($nmtv, "OldhItem") Then App_DisplayNode(DllStructGetData($nmtv, "NewParam"))
                Return True
            Case $TVN_ITEMEXPANDINGA, $TVN_ITEMEXPANDINGW
                Local $nmtv = DllStructCreate($tagNMTREEVIEW, $lParam)
                If DllStructGetData($nmtv, "NewhItem") Then App_ExpandNode(DllStructGetData($nmtv, "NewParam"))
        EndSwitch
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func DlgAbout_Create()
    If 0 = $dlgAbout Then
        Local Const $CLABELS = 4
        Local Const $RESLABELS[$CLABELS] = ["Description", "Version", "Copyright", "License"]
        Local $infoText = ""
        If @compiled Then
            Local Const $RESFIELDS[$CLABELS] = ["Comments", "FileVersion", "LegalCopyright", "LegalTrademarks"]
            For $i = 0 To UBound($RESFIELDS) - 1
                If 1 = $i Then $infoText &= @lf
                $infoText &= @lf & $RESLABELS[$i] & ": " & FileGetVersion(@ScriptFullPath, $RESFIELDS[$i])
            Next
        Else
            Local Const $AIWFIELDS[$CLABELS] = ["Comment=", "Fileversion=", "LegalCopyright=", "Field=LegalTrademarks\|"]
            Local $hf = FileOpen(@ScriptFullPath)
            Local $line = FileReadLine($hf)
            Local $matches
            While 0 = @error
                If "#endregion" = StringLeft($line, 10) Then ExitLoop
                Local $lbl = ""
                For $i = 0 To UBound($AIWFIELDS) - 1
                    $matches = StringRegExp($line, "#(?i)AutoIt3Wrapper_Res_" & $AIWFIELDS[$i] & "(?i)(.*)$", 1)
                    If @error Then ContinueLoop
                    $lbl = $RESLABELS[$i]
                    If 1 = $i Then $infoText &= @lf
                    ExitLoop
                Next
                If 0 < StringLen($lbl) Then $infoText &= @lf & $lbl & ": " & $matches[0]
                $line = FileReadLine($hf)
            WEnd
            FileClose($hf)
        EndIf
        $infoText &= @lf & "Author: doudou"
        $dlgAbout = SplashTextOn("About", $PRODUCT_TITLE & @lf & $infoText, 240, 180, -1, -1, 4 + 16, "Tahoma", 10)
        _GUICtrl_Subclass("dlgAbout")
        GUISetState(@SW_DISABLE, $frmMain)
    EndIf
    WinActivate($dlgAbout)
EndFunc

Func DlgAbout_Destroy($final = False)
    If 0 <> $dlgAbout Then
        _GUICtrl_Unsubclass("dlgAbout")
        $dlgAbout = 0
        SplashOff()
        If Not $final Then
            GUISetState(@SW_ENABLE, $frmMain)
            WinActivate($frmMain)
        EndIf
    EndIf
EndFunc

Func DlgAbout_WindowProc($hWnd, $uMsg, $wParam, $lParam)
    Switch $uMsg
        Case $WM_KEYDOWN
            If 27 = $wParam Then _WinAPI_PostMessage($frmMain, $WM_CLOSECHILD, $dlgAbout, 0)
            Return 0
        Case $WM_LBUTTONUP, $WM_RBUTTONUP
            _WinAPI_PostMessage($frmMain, $WM_CLOSECHILD, $dlgAbout, 0)
            Return 0
    EndSwitch
    Return _GUICtrl_SubclassedProc("dlgAbout", $hWnd, $uMsg, $wParam, $lParam)
EndFunc

Func TVWLib_On_TVM_REDIRECT($hWnd, $msg, $wParam, $lParam)
    App_SelectAndDisplayNode(Number($lParam))
    Return True
EndFunc

Func TVWLib_On_TVN_ITEMCHECKED($hWnd, $msg, $wParam, $lParam)
    If 0 = $lParam Then Return True
    Local $hWndFrom = _GUICtrl_GetHandle($frmMain_tvwLib)
    If 3 = $wParam Then
        _GUICtrlTreeView_SetStateImageIndexMod($hWndFrom, $lParam, 0)
    Else
        TVWLib_CheckChildren($lParam, 0 < $wParam)
    EndIf
    Return True
EndFunc

Func TVWLib_CheckChildren($hItem, $fChecked = True)
    Local $hWnd = _GUICtrl_GetHandle($frmMain_tvwLib)
    Local $hti = _GUICtrlTreeView_GetFirstChild($hWnd, $hItem)
    While $hti
        If 0 < _GUICtrlTreeView_GetStateImageIndex($hWnd, $hti) Then
            _GUICtrlTreeView_SetChecked($hWnd, $hti, $fChecked)
            TVWLib_CheckChildren($hti, $fChecked)
        EndIf
        $hti = _GUICtrlTreeView_GetNextSibling($hWnd, $hti)
    WEnd
EndFunc

Func _GUICtrlTreeView_SetStateImageIndexMod($hWnd, $hItem, $iIndex)
	Local $tItem = DllStructCreate($tagTVITEMEX)
	DllStructSetData($tItem, "Mask", $TVIF_STATE)
	DllStructSetData($tItem, "hItem", $hItem)
	DllStructSetData($tItem, "State", BitShift($iIndex, -12))
	DllStructSetData($tItem, "StateMask", $TVIS_STATEIMAGEMASK)
	Return __GUICtrlTreeView_SetItem($hWnd, $tItem)
EndFunc