Const UNREGISTER_COMMAND = "%REGASM% /verbose /unregister %FILENAME%"
Const REGISTER_COMMAND = "%REGASM% /verbose /codebase %FILENAME% /tlb:%TLBFILENAME%"
Const DEBUG_MODE = False
Dim regAsm

fUNCTION DummyCall()
End function

Function FindRegASM()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    sFramework = FSO.GetSpecialFolder(0) & "\Microsoft.NET\Framework\"
    NETVersions = Array("v3.5", "v2.0.50727", "v3.0", "v1.0.3705", "v1.1.4322")

    For Each dotNetVersion In NetVersions
        If FSO.FileExists(sFrameWork & dotNetVersion & "\RegAsm.exe") Then
            RegAsm = DQ(sFrameWork & dotNetVersion & "\RegAsm.exe")
            Exit For
        End If
    Next

    If DEBUG_MODE then Msgbox RegAsm

    Set FSO = Nothing
End Function

Function UninstallScreenCapture(folderName)
    Set WShell = CreateObject("WScript.Shell")
    fileName = DQ(folderName & "KnowledgeInbox.ScreenCapture.dll")
    TLBFileName = DQ(folderName & "KnowledgeInbox.ScreenCapture.tlb")
    cmd = Replace(UNREGISTER_COMMAND, "%REGASM%", regAsm)
    cmd = Replace(cmd, "%FILENAME%", fileName)
    cmd = Replace(cmd, "%TLBFILENAME%", TLBfileName)

    If DEBUG_MODE Then Msgbox cmd
    WSHell.Run cmd, , True
End Function

Function InstallScreenCapture(folderName)
    If DEBUG_MODE Then Msgbox folderName

    Set WShell = CreateObject("WScript.Shell")
    fileName = DQ(folderName & "KnowledgeInbox.ScreenCapture.dll")
    TLBFileName = DQ(folderName & "KnowledgeInbox.ScreenCapture.tlb")

    cmd = Replace(REGISTER_COMMAND, "%REGASM%", regAsm)
    cmd = Replace(cmd, "%FILENAME%", fileName)
    cmd = Replace(cmd, "%TLBFILENAME%", TLBfileName)
    If DEBUG_MODE Then Msgbox cmd
    WSHell.Run cmd, , True
End Function

Function DQ(byVal strText)
    DQ = Chr(34) & strText & Chr(34)
End Function

on error resume next

if DEBUG_MODE Then Msgbox "ENTER"
Dim foldername

If IsObject(Session) Then
    if DEBUG_MODE Then Msgbox "in if"
    If DEBUG_MODE Then Msgbox Session.Property("APPDIR")	
    folderName = Session.Property("APPDIR")
Else
    if DEBUG_MODE Then Msgbox "in Else"
    If DEBUG_MODE Then Msgbox WScript.ScriptFullName
    If DEBUG_MODE Then Msgbox WScript.ScriptName
    folderName = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
End If

Call findRegAsm()
Call UninstallScreenCapture(folderName)
Call InstallScreenCapture (folderName)
if DEBUG_MODE Then Msgbox "EXIT"