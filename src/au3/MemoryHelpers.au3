#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Memory helper functions.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
Func _sizeof($type)
    If IsDeclared("_SIZEOF_" & $type) Then Return Eval("_SIZEOF_" & $type)
    Local $t
    If IsDeclared($type) Then
        $t = DllStructCreate(Eval($type))
    Else
        $t = DllStructCreate($type)
    EndIf
    Local $result = DllStructGetSize($t)
    Assign("_SIZEOF_" & $type, $result, 2)
    Return $result
EndFunc

Func _carray_ptr($arrayPtr, $index, $typeName)
    Return $arrayPtr + ($index * _sizeof($typeName))
EndFunc