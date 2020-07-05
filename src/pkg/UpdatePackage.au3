; ------------------------------------------------------------------------------
;
; Version:        0.0.1
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Helper script for distribution packager (executed by AutoIt3Wrapper).
; $Revision: 1.1.1.1 $
; $Date: 2010/05/16 19:58:03 $
;
; ------------------------------------------------------------------------------
#include "..\au3\Product.au3"
Global $productEXE = $PRODUCT & ".exe"
If @AutoItX64 Then $productEXE = $PRODUCT & "64.exe"

Dim $fname = "..\..\build\" & $productEXE
If Not FileExists($fname) Then
    MsgBox(48, @ScriptName, $fname & " doesn't exsists, please compile " & $PRODUCT & " first.")
    Exit -1
EndIf

Global $hContents = FileOpen(@ScriptDir & "\PackageContents.au3", BitOr(2, 256))
If -1 = $hContents Then
    MsgBox(48, @ScriptName, "Failed to open " @ScriptDir & "\PackageContents.au3 for writing.")
    Exit -1
EndIf
FileWriteLine($hContents, ";**** Automatically generated by " & @ScriptName & " on " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC)
FileWriteLine($hContents, "FileInstall(""" & $fname & """, $destDir & ""\"")")
Pkg_WtiteContents("..\..\conf\plugins", "*.*", "plugins")
Pkg_WtiteContents("..\..\res\icons", "*.*", "icons")
Pkg_WtiteContents("..\..\res\xsl", "*.*", "xsl")
Pkg_WtiteContents("..\..\res\html", "*.*", "html")
Pkg_WtiteContents("..\..\res\css", "*.*", "css")
FileFlush($hContents)
FileClose($hContents)

Func Pkg_WtiteContents($dir, $mask, $dest)
    If "\" <> StringRight($dir, 1) Then $dir &= "\"
    Local $fname
    Local $hf = FileFindFirstFile($dir & $mask)
    While 0 = @error
        $fname = FileFindNextFile($hf)
        If @error Then ExitLoop
        If 1 = @extended Then ContinueLoop
        FileWriteLine($hContents, "FileInstall(""" & $dir & $fname & """, $destDir & ""\" & $dest & "\"")")
    WEnd
    FileClose($hf)
EndFunc