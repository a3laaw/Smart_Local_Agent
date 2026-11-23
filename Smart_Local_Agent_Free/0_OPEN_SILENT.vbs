Option Explicit

Dim fso, sh, scriptDir, userProfile, localApp, installedApp, appPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
userProfile = sh.ExpandEnvironmentStrings("%USERPROFILE%")
localApp = scriptDir & "\doc_tools\project_dashboard.py"
installedApp = userProfile & "\.gemini\doc_tools\project_dashboard.py"

If fso.FileExists(localApp) Then
    appPath = localApp
ElseIf fso.FileExists(installedApp) Then
    appPath = installedApp
Else
    MsgBox "Smart Projects Hub is not installed yet. Please run 1_INSTALL_FIRST_TIME.bat first.", 16, "Smart Projects Hub"
    WScript.Quit 1
End If

Function ProbePort(portNum)
    On Error Resume Next
    Dim http
    ProbePort = "free"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then
        Err.Clear
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    http.SetTimeouts 300, 300, 400, 600
    http.Open "GET", "http://127.0.0.1:" & portNum & "/api/state", False
    http.Send
    If Err.Number = 0 Then
        If http.Status = 200 And InStr(http.ResponseText, "2026-11-18-v10.13.2") > 0 Then
            ProbePort = "v10.13.2"
        ElseIf http.Status > 0 Then
            ProbePort = "other"
        End If
    End If
    Err.Clear
End Function

Function FindPythonW()
    Dim localAppData, candidates, c, tmpFile, ts, line, cmdLine
    FindPythonW = ""
    localAppData = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    candidates = Array( _
        localAppData & "\Programs\Python\Python313\pythonw.exe", _
        localAppData & "\Programs\Python\Python312\pythonw.exe", _
        localAppData & "\Programs\Python\Python311\pythonw.exe", _
        localAppData & "\Programs\Python\Python310\pythonw.exe", _
        localAppData & "\Programs\Python\Python39\pythonw.exe", _
        "C:\Python313\pythonw.exe", _
        "C:\Python312\pythonw.exe", _
        "C:\Python311\pythonw.exe" _
    )
    For Each c In candidates
        If fso.FileExists(c) Then
            FindPythonW = c
            Exit Function
        End If
    Next

    tmpFile = fso.GetSpecialFolder(2) & "\sph_pyw_detect.txt"
    cmdLine = "cmd /c python -c ""import sys,os; d=os.path.dirname(sys.executable); pw=os.path.join(d,'pythonw.exe'); print(pw if os.path.exists(pw) else sys.executable)"" > """ & tmpFile & """ 2>nul"
    sh.Run cmdLine, 0, True
    If Not fso.FileExists(tmpFile) Or fso.GetFile(tmpFile).Size = 0 Then
        cmdLine = "cmd /c py -3 -c ""import sys,os; d=os.path.dirname(sys.executable); pw=os.path.join(d,'pythonw.exe'); print(pw if os.path.exists(pw) else sys.executable)"" > """ & tmpFile & """ 2>nul"
        sh.Run cmdLine, 0, True
    End If
    If fso.FileExists(tmpFile) Then
        If fso.GetFile(tmpFile).Size > 0 Then
            Set ts = fso.OpenTextFile(tmpFile, 1)
            line = Trim(ts.ReadLine)
            ts.Close
            If fso.FileExists(line) Then FindPythonW = line
        End If
        On Error Resume Next
        fso.DeleteFile tmpFile, True
        On Error GoTo 0
    End If
End Function

Function FindAppBrowser()
    Dim pf, pf86, la, bCandidates, b
    FindAppBrowser = ""
    pf = sh.ExpandEnvironmentStrings("%ProgramFiles%")
    pf86 = sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
    la = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    bCandidates = Array( _
        pf86 & "\Microsoft\Edge\Application\msedge.exe", _
        pf & "\Microsoft\Edge\Application\msedge.exe", _
        pf & "\Google\Chrome\Application\chrome.exe", _
        pf86 & "\Google\Chrome\Application\chrome.exe", _
        la & "\Google\Chrome\Application\chrome.exe", _
        la & "\Microsoft\Edge\Application\msedge.exe" _
    )
    For Each b In bCandidates
        If fso.FileExists(b) Then
            FindAppBrowser = b
            Exit Function
        End If
    Next
End Function

Sub OpenAppWindow(appUrl)
    Dim browserExe, useTab
    useTab = False
    If WScript.Arguments.Count > 0 Then
        If LCase(WScript.Arguments(0)) = "--browser-tab" Then useTab = True
    End If
    If Not useTab Then
        browserExe = FindAppBrowser()
        If browserExe <> "" Then
            sh.Run """" & browserExe & """ --app=""" & appUrl & """", 1, False
            Exit Sub
        End If
    End If
    sh.Run appUrl, 1, False
End Sub

' 1) If already running Smart Projects Hub on 8765..8772, open standalone app window immediately without starting another server
Dim pCheck, targetPort, pywExe, i
For pCheck = 8765 To 8772
    If ProbePort(pCheck) = "v10.13.2" Then
        OpenAppWindow "http://127.0.0.1:" & pCheck
        WScript.Quit 0
    End If
Next

targetPort = 8765
For pCheck = 8765 To 8772
    If ProbePort(pCheck) = "free" Then
        targetPort = pCheck
        Exit For
    End If
Next

pywExe = FindPythonW()
If pywExe = "" Then
    MsgBox "Python 3 was not found. Please run 1_INSTALL_FIRST_TIME.bat first.", 16, "Smart Projects Hub"
    WScript.Quit 1
End If

' 2) Launch dashboard silently in background (window style 0 + pythonw.exe = zero black CMD window)
sh.Run """" & pywExe & """ """ & appPath & """ --host 127.0.0.1 --port " & targetPort & " --no-browser", 0, False

' 3) Wait briefly for local server readiness, then open standalone app window
For i = 1 To 32
    WScript.Sleep 250
    If ProbePort(targetPort) <> "free" Then Exit For
Next

OpenAppWindow "http://127.0.0.1:" & targetPort
WScript.Quit 0
