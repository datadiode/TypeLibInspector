#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Type information library interface definitions.
; Requirements:   TLIConstants, COMInterfaces
; $Revision: 17 $
; $Date: 2010-05-03 00:59:28 +0200 (Mon, 03 May 2010) $
;
; ------------------------------------------------------------------------------
#include <TLIConstants.au3>
#include <COMInterfaces.au3>

Global $tagITypeInfo = $tagIUnknown & _
    "GetTypeAttr long(ptr);" & _
    "GetTypeComp long(ptr);" & _
    "GetFuncDesc long(uint;ptr);" & _
    "GetVarDesc long(uint;ptr);" & _
    "GetNames long(" & $tMEMBERID & ";ptr;uint;ptr);" & _
    "GetRefTypeOfImplType long(uint;ptr);" & _
    "GetImplTypeFlags long(uint;ptr);" & _
    "GetIDsOfNames long(ptr;uint;ptr);" & _
    "Invoke long(ptr;" & $tMEMBERID & ";short;ptr;ptr;ptr);" & _
    "GetDocumentation long(" & $tMEMBERID & ";ptr;ptr;ptr;ptr);" & _
    "GetDllEntry long(" & $tMEMBERID & ";" & $tINVOKEKIND & ";ptr;ptr;ptr);" & _
    "GetRefTypeInfo long(long;ptr);" & _
    "AddressOfMember long(" & $tMEMBERID & ";" & $tINVOKEKIND & ";ptr);" & _
    "CreateInstance long(ptr;" & $tREFIID & ";ptr);" & _
    "GetMops long(" & $tMEMBERID & ";ptr);" & _
    "GetContainingTypeLib long(ptr;ptr);" & _
    "ReleaseTypeAttr long(ptr);" & _
    "ReleaseFuncDesc long(ptr);" & _
    "ReleaseVarDesc long(ptr);"
Global $tagITypeLib = $tagIUnknown & _
    "GetTypeInfoCount long();" & _
    "GetTypeInfo long(uint;ptr);" & _
    "GetTypeInfoType long(uint;ptr);" & _
    "GetTypeInfoOfGuid long(" & $tREFGUID & ";ptr);" & _
    "GetLibAttr long(ptr);" & _
    "GetTypeComp long(ptr);" & _
    "GetDocumentation long(int;ptr;ptr;ptr;ptr);" & _
    "IsName long(ptr;ulong;ptr);" & _
    "FindName long(ptr;ulong;ptr;ptr;ptr);" & _
    "ReleaseTLibAttr long(ptr);"
