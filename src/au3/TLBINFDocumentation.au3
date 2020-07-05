#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Common members for documentable TLBINF objects.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#Region ###TLIObj common property getters###
Func TLIObj_Name($oSelf)
    Local $di = $oSelf._ReadDocumentation()
    If IsArray($di) Then Return $di[$TDOC_NAME]
    Return "_"
EndFunc

Func TLIObj_HelpString($oSelf)
    Local $di = $oSelf._ReadDocumentation()
    If IsArray($di) Then Return $di[$TDOC_DOCSTRING]
    Return ""
EndFunc

Func TLIObj_HelpContext($oSelf)
    Local $di = $oSelf._ReadDocumentation()
    If IsArray($di) Then Return $di[$TDOC_HELPCONTEXT]
    Return 0
EndFunc

Func TLIObj_HelpFile($oSelf)
    Local $di = $oSelf._ReadDocumentation()
    If IsArray($di) Then Return $di[$TDOC_HELPFILE]
    Return ""
EndFunc
#EndRegion ;TLIObj common property getters
