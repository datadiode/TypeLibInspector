#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Definitions for basic COM interfaces to use with _AutoItObject_WrapperCreate().
; Requirements:   COMConstants
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#include <COMConstants.au3>
Global $tagIUnknown = _
    "QueryInterface long(ptr;ptr;ptr);" & _
    "AddRef ulong();" & _
    "Release ulong();"
Global $tagIDispatch = $tagIUnknown & _
    "GetTypeInfoCount long(ptr);" & _
    "GetTypeInfo long(uint;" & $tLCID & ";ptr);" & _
    "GetIDsOfNames long(" & $tREFIID & ";ptr;uint;" & $tLCID & ";ptr);" & _
    "Invoke long(long;" & $tREFIID & ";" & $tLCID & ";word;ptr;ptr;ptr;ptr);"
