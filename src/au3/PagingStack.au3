#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Stack-like storage with moving fixed size page.
; $Revision: 1.1 $
; $Date: 2010/05/23 20:43:15 $
;
; ------------------------------------------------------------------------------
Func _PagingStack_Create($pageSize, $allocate = False)
    Local $size = 3
    If $allocate Then $size += $pageSize
    Local $result[$size] = [$pageSize, 0, 3]
    Return $result
EndFunc

Func _PagingStack_Clear(ByRef $stack)
    If 2 < UBound($stack) Then
        Local $realloc = (UBound($stack) = $stack[0] + 3)
        Redim $stack[3]
        If $realloc Then ReDim $stack[$stack[0] + 3]
        $stack[1] = 0
        $stack[2] = 3
    Else
        Return SetError(1, 0, False)
    EndIf
    Return True
EndFunc

Func _PagingStack_Push(ByRef $stack, $value)
    Local $size = UBound($stack)
    If 3 > $size Then Return SetError(1, 0, -1)
    If $size < $stack[0] + 3 Then Redim $stack[$size + 1]
    If $stack[1] < $stack[0] Then
        $stack[$stack[2] + $stack[1]] = $value
        $stack[1] += 1
    Else
        $stack[$stack[2]] = $value
        If $stack[2] >= 2 + $stack[1] Then
            $stack[2] = 3
        Else
            $stack[2] += 1
        EndIf
    EndIf
    Return $stack[1]
EndFunc

Func _PagingStack_GetTail(ByRef $stack)
    If 3 > UBound($stack) Then Return SetError(1, 0, Default)
    If $stack[2] > $stack[1] + 2 Then Return SetError(1, 1, Default)
    If 3 >= $stack[2] Then Return $stack[$stack[1] + 2]
    Return $stack[$stack[2] - 1]
EndFunc

Func _PagingStack_GetSize(ByRef $stack)
    If 3 > UBound($stack) Then Return SetError(1, 0, 0)
    Return $stack[1]
EndFunc

Func _PagingStack_Get(ByRef $stack, $index)
    If 3 > UBound($stack) Then Return SetError(1, 0, Default)
    If $index >= $stack[1] Then Return SetError(2, 0, Default)
    $index += $stack[2]
    If $index >= $stack[1] + 3 Then $index -= $stack[1] 
    Return $stack[$index]
EndFunc