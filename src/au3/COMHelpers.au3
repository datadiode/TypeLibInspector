#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    COM helper functions.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
Func _COM_Succeded(Const ByRef $status)
    If IsArray($status) Then Return $status[0] >= 0
    Return False
EndFunc

Func _COM_SetError(Const ByRef $status, $result = 0)
    If IsArray($status) Then Return SetError($status[0], 0, $result)
    Return SetError(-1, 0, $result)
EndFunc
