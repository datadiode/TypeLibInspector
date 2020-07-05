#region - Compiler Directives
#AutoIt3Wrapper_outfile=..\..\build\TypeLibInspector-sfx.exe
#AutoIt3Wrapper_icon=..\..\res\icons\TypeLibInspector.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Comment=Self-extracting application package
#AutoIt3Wrapper_Res_Description=Type Library Inspector Distribution Package
#AutoIt3Wrapper_Res_Fileversion=0.1.0
#AutoIt3Wrapper_Res_LegalCopyright=© 2010 diVISION
#AutoIt3Wrapper_Res_Field=ProductName|Type Library Inspector
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Run_Before=%autoitdir%\AutoIt3.exe "%scriptdir%\UpdatePackage.au3"
#AutoIt3Wrapper_Allow_Decompile=y
#endregion
; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Distribution packager for TypeLibInspector.
; $Revision: 1.3 $
; $Date: 2010/07/29 14:27:17 $
;
; ------------------------------------------------------------------------------

If Not @compiled Then
    MsgBox(16, @ScriptName, "This is packager script, use only compiled.")
    Exit -1
EndIf

#include "..\au3\Product.au3"
Global $productEXE = $PRODUCT & ".exe"
If @AutoItX64 Then $productEXE = $PRODUCT & "64.exe"

Dim $destDir = FileSelectFolder("Select folder where " & $PRODUCT_TITLE & " will be stored", "", 7)
If 0 = StringLen($destDir) Then Exit 0

DirCreate($destDir & "\plugins\")
DirCreate($destDir & "\icons\")
DirCreate($destDir & "\html\")
DirCreate($destDir & "\css\")
DirCreate($destDir & "\xsl\")

#include "PackageContents.au3"

If 6 = MsgBox(4 + 32, $PRODUCT_TITLE, "Installation complete. Start " & $PRODUCT_TITLE & " now?") Then ShellExecute($destDir & "\" & $productEXE)
