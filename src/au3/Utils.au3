#include-once
; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Utility functions for TypeLibInspector.
; $Revision: 1.2 $
; $Date: 2010/07/29 14:22:43 $
;
; ------------------------------------------------------------------------------

Func _GUICtrl_GetHandle($control)
    If IsHWnd($control) Then Return $control
    Return GUICtrlGetHandle($control)
EndFunc

Func _GUICtrl_GetMessagePos($control)
    Local $p = DllCall("user32.dll", "DWORD", "GetMessagePos")
    Local $pt = DllStructCreate($tagPOINT)
    DllStructSetData($pt, "X", _WinAPI_LoWord($p[0]))
    DllStructSetData($pt, "Y", _WinAPI_HiWord($p[0]))
    _WinAPI_ScreenToClient(_GUICtrl_GetHandle($control), $pt)
    Local $result[2] = [DllStructGetData($pt, "X"), DllStructGetData($pt, "Y")]
    Return $result
EndFunc

Func _GUICtrl_Subclass($controlName)
    Local $hlvw = _GUICtrl_GetHandle(Eval($controlName))
    If $hlvw Then
        Local $hProc = DllCallbackRegister($controlName & "_WindowProc", "int", "hwnd;uint;wparam;lparam")
        If $hProc Then
            Assign($controlName & "_hProc", $hProc, 2)
            Assign($controlName & "_hProcOld", _WinAPI_SetWindowLong($hlvw, $GWL_WNDPROC, DllCallbackGetPtr($hProc)), 2)
        EndIf
        Return $hProc
    EndIf
    Return 0
EndFunc

Func _GUICtrl_Unsubclass($controlName)
    Local $hProc = Eval($controlName & "_hProc")
    If $hProc Then
        _WinAPI_SetWindowLong(_GUICtrl_GetHandle(Eval($controlName)), $GWL_WNDPROC, Eval($controlName & "_hProcOld"))
        DllCallbackFree($hProc)
        Assign($controlName & "_hProc", 0)
        Assign($controlName & "_hProcOld", 0)
    EndIf
EndFunc

Func _GUICtrl_SubclassedProc($controlName, $hWnd, $uMsg, $wParam, $lParam)
    Local $hProc = Eval($controlName & "_hProcOld")
    If $hProc Then Return _WinAPI_CallWindowProc($hProc, $hWnd, $uMsg, $wParam, $lParam)
    Return 0
EndFunc

Func _COM_ProgIDFromCLSID($strGUID)
    Local $clsid = _AutoItObject_CLSIDFromString($strGUID)
    Local $progid = DllStructCreate("ptr") 
    Local $hr = DllCall("ole32.dll", "ulong", "ProgIDFromCLSID", $tREFGUID, DllStructGetPtr($clsid), "ptr", DllStructGetPtr($progid, 1))
    If _COM_Succeded($hr) Then
        Local $iLen = DllCall("kernel32.dll", "int", "lstrlenW", "ptr", DllStructGetData($progid, 1))
        Local $result = DllStructGetData(DllStructCreate("wchar[" & $iLen[0] & "]", DllStructGetData($progid, 1)), 1)
        __Au3Obj_CoTaskMemFree(DllStructGetData($progid, 1))
        Return $result
    EndIf
    Return _COM_SetError($hr, "")
EndFunc

Func _String_Capitalize($s)
    Return StringUpper(StringLeft($s, 1)) & StringRight($s, StringLen($s) - 1)
EndFunc

Func _URL_GetParameters($uri)
    Local $result[1] = [0]
    Local $parts = StringSplit($uri, "?")
    If 0 < $parts[0] Then
        $result = StringSplit($parts[$parts[0]], "&=")
    EndIf
    Return $result
EndFunc

Func _File_GetParentPath($path)
    Local $result = ""
    Local $opt = AutoItSetOption("ExpandEnvStrings", 1)
    Local $i = StringInStr($path, "\", 0, -1)
    If 1 < $i Then $result = StringLeft($path, $i - 1)
    If 0 = StringLen($result) Then $result = "."
    AutoItSetOption("ExpandEnvStrings", $opt)
    Return $result
EndFunc

Func _File_GetAbsolutePath($path, $relative = @WorkingDir)
    Local $result = $relative
    Local $opt = AutoItSetOption("ExpandEnvStrings", 1)
    Local $test = StringRegExp($path, "^([a-zA-Z]:)|(\\\\)|(/)|(\w+:)", 0)
    If $test Then
        $result = $path
        AutoItSetOption("ExpandEnvStrings", $opt)
        Return $result
    EndIf
    If "\" <> StringLeft($path, 1) Then $result &= "\"
    AutoItSetOption("ExpandEnvStrings", $opt)
    Return $result & $path
EndFunc
