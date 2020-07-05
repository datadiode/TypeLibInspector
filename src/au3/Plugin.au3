#include-once
; ------------------------------------------------------------------------------
;
; Version:        0.1.0
; AutoIt Version: 3.3.4.0
; Language:       English
; Author:         doudou
; License:        GNU/GPL, s. LICENSE.txt
; Description:    Plugin functions for TypeLibInspector.
; $Revision: 1.3 $
; $Date: 2010/07/29 14:28:47 $
;
; ------------------------------------------------------------------------------

Global Const $TLDPLUGIN_DEF = 0
Global Const $TLDPLUGIN_CMD = 1

Global $app_plugins

Func TLDPlugin_LoadList($menu, $category = "tools")
    Local $dir = @ScriptDir
    If Not @compiled Then $dir &= "\..\..\conf"
    $dir &= "\plugins\"
    Local $hf = FileFindFirstFile($dir & "*.ini")
    Local $plgDef
    While 0 = @error
        $plgDef = FileFindNextFile($hf)
        If @error Then ExitLoop
        If 1 = @extended Then ContinueLoop
        If $category = IniRead($dir & $plgDef, "Plugin", "category", "tools") Then
            Local $count = UBound($app_plugins)
            If 0 = $count Then
                Dim $app_plugins[1][2]
            Else
                Redim $app_plugins[$count + 1][2]
            EndIf
            $app_plugins[$count][$TLDPLUGIN_DEF] = $dir & $plgDef
            $app_plugins[$count][$TLDPLUGIN_CMD] = GUICtrlCreateMenuItem(IniRead($dir & $plgDef, "Plugin", "name", $plgDef), $menu)
        EndIf
    WEnd
    FileClose($hf)
    Return Ubound($app_plugins)
EndFunc

Func TLDPlugin_IfExec($msg)
    If Not IsArray($app_plugins) Then Return False
    For $i = 0 To UBound($app_plugins) - 1
        Switch $msg[0]
            Case $app_plugins[$i][$TLDPLUGIN_CMD]
                Call("TLDPlugin_Exec" & IniRead($app_plugins[$i][$TLDPLUGIN_DEF], "Plugin", "type", "Unknown"), $i)
                If 0xDEAD = @error And 0xBEEF = @extended Then App_MsgBox("Unknown plugin type in " & $app_plugins[$i][$TLDPLUGIN_DEF], $APPMB_ERROR)
                Return True
        EndSwitch
    Next
    Return False
EndFunc

Func TLDPlugin_ExecXSLExport($pluginNo)
    If Not IsObj($typeDoc_objDOM) Then Return
    Local $split = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "splitOnBoundaries", 0)
    Local $path = ""
    Local $plgName = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "Plugin", "name", "files") 
    Local $ext = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "ext", "xml")
    If 0 < $split Then
        $path = FileSelectFolder("Select folder where " & $plgName & " will be saved", "", 7)
    Else
        $path = FileSaveDialog("Save " & $plgName, "", "Target Files (*." & $ext & ")", 18, $app_objTypeLib.Name & "-" & $app_objTypeLib.MajorVersion & "." & $app_objTypeLib.MinorVersion & "." & $ext)
    EndIf
    If 0 = StringLen($path) Then Return
    
    Local $wd = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "Plugin", "workingDir", _File_GetParentPath($app_plugins[$pluginNo][$TLDPLUGIN_DEF]))
    Local $xslProc = App_CreateXSLProcessor(_File_GetAbsolutePath(IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "xslt", ""), $wd), "")
    If Not IsObj($xslProc) Then Return
    
    App_ExpandAll()
    
    ProgressOn("Exporting", $app_objTypeLib.Name)
    $xslProc.input = $typeDoc_objDOM
    
    Local $xsltParams = IniReadSection($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLTParams")
    If 0 = @error Then
        For $i = 1 To $xsltParams[0][0]
            $xslProc.addParameter($xsltParams[$i][0], $xsltParams[$i][1])
        Next
    EndIf
    
    Local $nl = 0
    Switch $split
        Case 0
            ProgressSet(10, $path)
            App_SaveXSLOutput($xslProc, $path)
        Case 1
            $nl = $typeDoc_objDOM.selectNodes("/TypeLib | /TypeLib/Types/*")
        Case 2
            $nl = $typeDoc_objDOM.selectNodes("/TypeLib | /TypeLib/Types/* | /TypeLib/Types/*/Properties/* | /TypeLib/Types/*/Methods/*")
    EndSwitch
    
    If IsObj($nl) And 0 < $nl.length() Then
        Local $node, $name
        Local $c = $nl.length()
        For $i = 0 To $c - 1
            Local $process = True
            $node = $nl.item($i)
            $name = $node.nodeName & "-" & $node.getAttribute("name")
            Local $member = ("Properties" = $node.parentNode.nodeName Or "Methods" = $node.parentNode.nodeName)
            If $member Then
                Local $type = $node.selectSingleNode("../..")
                If 0 = BitAnd(GuiCtrlRead($type.getAttribute("tldapp-ptr")), $GUI_CHECKED) Then $process = False
                
                $name = $type.nodeName & "-" & $type.getAttribute("name") & "-" & $name
                If $process Then $xslProc.addParameter("key", "member")
            Else
                If 0 = BitAnd(GuiCtrlRead($node.getAttribute("tldapp-ptr")), $GUI_CHECKED) Then $process = False
                If $process Then $xslProc.addParameter("key", "type")
            EndIf
            ProgressSet(60 * $i / $c, $name)
            
            If $process Then
                $xslProc.addParameter("keyVal", $name)
                Local $fileName = "$name"
                If $member Then
                    $fileName = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "memberFilenameTpl", $fileName)
                Else
                    $fileName = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "typeFilenameTpl", $fileName)
                    If "TypeLib" = $node.nodeName Then $fileName = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "libFilenameTpl", $fileName)
                EndIf
                $fileName = Execute($fileName)
                If 0 = StringLen($fileName) Then $fileName = $name
                If StringInStr($fileName, "\") Then
                    Local $dir = $path & "\" & _File_GetParentPath($fileName)
                    If Not FileExists($dir) And Not DirCreate($dir) Then
                        App_MsgBox("Cannot create directory " & $dir, $APPMB_ERROR)
                        ExitLoop
                    EndIf
                EndIf
                App_SaveXSLOutput($xslProc, $path & "\" & $fileName & "." & $ext)
            EndIf
        Next
    ElseIf 0 < $split Then
        App_MsgBox("Nothing to do.")
    EndIf
    
    ProgressSet(60, "Completing...")
    
    If FileExists($path) Then
        Local $copyRes = StringSplit(IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "copyResources", ""), ",")
        Local $destDir = $path
        If 0 = $split Then $destDir = _File_GetParentPath($path)
        Local $src = ""
        Local $dest = ""
        For $i = 1 To $copyRes[0]
            ProgressSet(60 + (40 * $i / $copyRes[0]), "Copying " & $copyRes[$i])
            $src = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "Copy" & $copyRes[$i], "src", "")
            If 0 = StringLen($src) Then ContinueLoop
            $dest = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "Copy" & $copyRes[$i], "dest", "")
            If 0 = StringLen($dest) Then ContinueLoop
            If Not FileCopy(_File_GetAbsolutePath($src, $wd), _File_GetAbsolutePath($dest, $destDir), 9) Then App_MsgBox("Failed to copy " & $copyRes[$i] & " to " & $dest, $APPMB_ERROR)
        Next
        Local $show = IniRead($app_plugins[$pluginNo][$TLDPLUGIN_DEF], "XSLExport", "displayResult", 0)
        If 1 = $show Then
            ShellExecute($path)
        ElseIf 0 < StringLen($show) Then
            ShellExecute($path, "", "", $show)
        EndIf
    EndIf
    ProgressOff()
EndFunc
