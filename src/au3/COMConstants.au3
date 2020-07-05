#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Frequently used COM constants.
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#include <Constants.au3>
#include <WinAPI.au3>

Global Const $LOCALE_USER_DEFAULT = _WinAPI_MAKELCID(_WinAPI_MAKELANGID($LANG_NEUTRAL, $SUBLANG_DEFAULT), $SORT_DEFAULT)
Global Const $LOCALE_SYSTEM_DEFAULT = _WinAPI_MAKELCID(_WinAPI_MAKELANGID($LANG_NEUTRAL, $SUBLANG_SYS_DEFAULT), $SORT_DEFAULT)

Global Const $tREFIID       = "ptr"
Global Const $tREFGUID      = $tREFIID
Global Const $tLCID         = "ulong"
