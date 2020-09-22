Attribute VB_Name = "MainSub"
'---------VERSION TYPES---------
Private Const BetaVer = "BETA"
Private Const AlphaVer = "ALPHA"
'-------------------------------
Public Const ProgRelease = AlphaVer
'-------------------------------

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type


Sub Main()
    On Error Resume Next
    Debug.Print "ComTalk Command Parameters: " & Command
    
    Temp = INI.GetExtra("LockDown", "UseLockDown")
    If Temp = "FALSE" Then
        NoLockDown = True
        Options.OptionsPage(16).Enabled = False
        AddLogText "INI COMMAND - Disable LockDown"
    End If
    Temp = INI.GetExtra("CusMenus", "UseNewMenu")
    If Temp = "FALSE" Then
        PopupDLLFail = True
        AddLogText "INI COMMAND - Disable Custom Menus"
    Else
        If HasCommand("/Nm") = True Then ' Parameter to disable XP popup menus
            PopupDLLFail = True
            AddLogText "PARAMETER COMMAND - Disable Custom Menus"
        End If
    End If
    If HasCommand("/Nd") = True Then ' Parameter to disable Disk Space checking (if ComTalk crashes on startup)
        NoDiskSpaceCheck = True
        CMDDDC = True
        Options.dsc.Enabled = False
        AddLogText "PARAMETER COMMAND - Disable disk space check"
    End If
    If HasCommand("/Nl") = True Then ' Parameter to disable LockDown (if user doesn't want this option)
        NoLockDown = True
        Options.OptionsPage(16).Enabled = False
        AddLogText "PARAMETER COMMAND - Disable LockDown"
    End If
    If HasCommand("/Ns") = False Then
        Load frmSplash
    Else
        AddLogText "PARAMETER COMMAND - Disable Splash"
        Load mainfrm
    End If
End Sub

Function HasCommand(LookFor As String) As Boolean
    HasCommand = False
    If InStr(1, Command, LookFor, vbBinaryCompare) <> 0 Then HasCommand = True
End Function

Public Function ExecCmd(cmdline As String) As Long ' Runs a program with parameters and wait's for it to finish
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim lngRC As Long
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    start.wShowWindow = 0
    start.dwXSize = 1
    start.dwYSize = 1
    ' Start the shelled application:
    lngRC = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
        HIGH_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    lngRC = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, lngRC)
    Call CloseHandle(proc.hProcess)
    ExecCmd = lngRC
End Function

Private Function FileOrPathExists(Path As String, Optional File As String) As Boolean
    If File = "" Then
        FileOrPathExists = Not (Dir(Path) = "")
    Else
        FileOrPathExists = Not (Dir(Path & File) = "")
    End If
End Function

Private Function NeedSlash() As String ' Check if the path needs a slash at the end
    If Right(App.Path, 1) = "\" Then
        NeedSlash = ""
    Else
        NeedSlash = "\"
    End If
End Function

