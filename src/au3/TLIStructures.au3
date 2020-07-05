#include-once
; ------------------------------------------------------------------------------
;
; Version:        1.0.3
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Type information library structure definitions.
; Requirements:   COMConstants, TLIConstants
; $Revision: 24 $
; $Date: 2010-07-23 03:04:51 +0200 (Fri, 23 Jul 2010) $
;
; ------------------------------------------------------------------------------
#include <COMConstants.au3>
#include <TLIConstants.au3>

Global Const $_PADSHORT_    = "word;"

Global Const $tagTYPEDESC   = "ulong_ptr lpItemDesc;ushort vt;" & $_PADSHORT_
Global Const $tagARRAYDESC  = $tagTYPEDESC & "ushort cDims;ulong cElements; long lLbound;" & $_PADSHORT_
Global Const $tagIDLDESC    = "ulong dwReserved;ushort wIDLFlags;" & $_PADSHORT_
Global Const $tagPARAMDESCEX    = "ulong_ptr cBytes;ptr;ushort vt;ushort r1;ushort r2;ushort r3;ptr data;ptr;"
Global Const $tagPARAMDESC  = "ptr lpParamDescEx;ushort wParamFlags;" & $_PADSHORT_
Global Const $tagELEMDESC   = $tagTYPEDESC & $tagPARAMDESC
Global Const $tagVARDESC    = $tMEMBERID & " MemID;ptr lpstrSchema;ulong_ptr lpVarValue;" & $tagELEMDESC & _
    "ushort wVarFlags;int varkind;"
Global Const $tagFUNCDESC   = $tMEMBERID & " MemID;ptr lprgSCode;ptr lprgElemDescParam;int funckind;" & $tINVOKEKIND & " InvKind;" & _
    "int callconv;short cParams;short cParamsOpt;short oVft;short cScodes;" & $tagELEMDESC & "ushort wFuncFlags;"
Global Const $tagTYPEATTR   = "byte guid[16];" & $tLCID & " lcid;ulong dwReserved;" & $tMEMBERID & " memidConstructor;" & _
    $tMEMBERID & " memidDestructor;ptr lpstrSchema;ulong cbSizeInstance;int typekind;ushort cFuncs;" & _
    "ushort cVars;ushort cImplTypes;ushort cbSizeVft;ushort cbAlignment;ushort wTypeFlags;ushort wMajorVerNum;" & _
    "ushort wMinorVerNum;" & $tagTYPEDESC & $tagIDLDESC
Global Const $tagTLIBATTR   = "byte guid[16];" & $tLCID & " lcid;int syskind;ushort wMajorVerNum;ushort wMinorVerNum;ushort wLibFlags;"
