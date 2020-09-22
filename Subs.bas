Attribute VB_Name = "LoggingSubs"
Dim DoNext As Boolean
Dim IAmSilent As Boolean
Public Enum LogMyText
    YesLog = 0
    NoLog = 1
End Enum

Public Enum OverideSilentOption
    NoOveride = 0
    YesSilent = 1
    YesNotSilent = 2
End Enum

Public Sub Req(RequestName As String) ' Menu click - Add to log file
    If GoCharacter = False Then Exit Sub
    
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, "User Request " & Chr(34) & RequestName & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Debug.Print "User Request " & Chr(34) & RequestName & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Close #1
    End If
End Sub

Public Sub Action(ActionName) ' Character action - Add to log file and play action
    If GoCharacter = False Then Exit Sub
    
    mainfrm.Agent1.Characters("Genie").Play ActionName
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, "Character Action " & Chr(34) & ActionName & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Debug.Print "Character Action " & Chr(34) & ActionName & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Close #1
    End If
End Sub

Public Sub AddLogText(Text) ' Add text to the log file
    If GoCharacter = False Then Exit Sub
    
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, Text & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Debug.Print Text & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Close #1
    End If
End Sub

Public Sub SpeakText(tts, Optional OverideSILENT As OverideSilentOption) ' Speaks text and adds text to the log file
    Static texttopass As String

    If GoCharacter = False Then Exit Sub
    
    texttopass = tts
    tts = PlugIns.CheckPassText(texttopass)
    SpVol = GetSetting("ComTalk", "Options", "SpeechVol", 65535)
    sptype = GetSetting("ComTalk", "Options", "SpeechType", 0)
    Select Case sptype
    Case 0
        sptype = "Normal"
    Case 1
        sptype = "Monotone"
    Case 2
        sptype = "Whisper"
    End Select
    
    If OverideSILENT = 0 Then
        IAmSilent = GetSetting("ComTalk", "Options", "Silent", False)
    ElseIf OverideSILENT = 1 Then
        IAmSilent = False
    Else
        IAmSilent = True
    End If
    
    If IAmSilent = True Then
        mainfrm.Agent1.Characters("Genie").Think tts
    Else
        mainfrm.Agent1.Characters("Genie").Speak "\rst\"
        mainfrm.Agent1.Characters("Genie").Speak "\vol=" & SpVol & "\ \chr=" & Chr(34) & sptype & Chr(34) & "\ \spd=" & speednum & "\ \pit=" & pitchnum & "\" & tts
    End If
    
    If UseLog = 1 Then
        Open "ComTalkLog.txt" For Append As #1
        Print #1, "Character Spoke " & Chr(34) & OneLine(tts) & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Debug.Print "Character Spoke " & Chr(34) & OneLine(tts) & Chr(34) & " (" & Format(Now, "HH:MM AM/PM") & ")"
        Close #1
    End If
    
    tts = ""
End Sub

Public Sub BeginLog() ' Starts the log file
    If UseLog = 1 Then
        logexists = Not (Dir("ComTalkLog.txt") = "")
        If logexists = True Then
            Open "ComTalkLOG.txt" For Append As #1
            Print #1, " "
            Print #1, " "
            Print #1, "SESSION STARTS - " & Format(Now, "DD/MM/YYYY HH:MM AM/PM")
            Debug.Print " "
            Debug.Print " "
            Debug.Print "SESSION STARTS - " & Format(Now, "DD/MM/YYYY HH:MM AM/PM")
            Close #1
        Else
            Open "ComTalkLOG.txt" For Append As #1
            Print #1, "----------------ComTalk Log File----------------"
            Print #1, " "
            Print #1, "SESSION STARTS - " & Now
            Debug.Print "----------------ComTalk Log File----------------"
            Debug.Print " "
            Debug.Print "SESSION STARTS - " & Now
            Close #1
        End If
    End If
End Sub

Public Sub EndLog() ' Ends the log file's current session
    SaveSetting "ComTalk", "Program", "IsOpen", 0
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, "SESSION ENDS - " & Format(Now, "DD/MM/YYYY HH:MM AM/PM")
        Debug.Print "SESSION ENDS - " & Format(Now, "DD/MM/YYYY HH:MM AM/PM")
        Close #1
    End If
    SaveSetting "ComTalk", "Program", "IsOpen", 0
End Sub

Public Sub ChangingCharacter() ' Logs character change
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, "Character Change " & Chr(34) & mainfrm.Character.filename & Chr(34) & " (" & Format(Now, "HH:MM") & ")"
        Debug.Print "Character Change " & Chr(34) & mainfrm.Character.filename & Chr(34) & " (" & Format(Now, "HH:MM") & ")"
        Close #1
    End If
End Sub

Public Sub IdlingChar() ' Logs character inactivity
    On Error Resume Next
    If UseLog = 1 Then
        Open "ComTalkLOG.txt" For Append As #1
        Print #1, "Idling (" & Format(Now, "HH:MM AM/PM") & ")"
        Debug.Print "Idling (" & Format(Now, "HH:MM AM/PM") & ")"
        Close #1
    End If
End Sub

Public Sub PrintTempFile(CTTempText) ' Writes startup data to the TEMP file (not the LOG file)
    On Error Resume Next
    Open "C:\ComTalkTemp.txt" For Append As #2
    Print #2, CTTempText
    Close #2
    
    Debug.Print CTTempText
End Sub

Public Sub CloseTempFile() ' Kills the TEMP file if ComTalk loads sucessfully
    On Error Resume Next
    Kill "C:\ComTalkTemp.txt"
End Sub

Public Sub OpenTempFile() ' Starts the TEMP file
    On Error Resume Next
    CTMPFLE = FreeFile
    Open "C:\ComTalkTemp.txt" For Output As CTMPFLE
    Print #CTMPFLE, "                                 ComTalk Temp File:"
    Print #CTMPFLE, "            This file is automatically removed after ComTalk loads."
    Print #CTMPFLE, "If ComTalk has an error while loading, send this file to dean_camera@hotmail.com"
    Print #CTMPFLE, "--------------------------------------------------------------------------------"
    Print #CTMPFLE, " "
    Close #CTMPFLE
End Sub
