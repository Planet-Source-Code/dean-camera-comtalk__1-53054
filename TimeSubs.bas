Attribute VB_Name = "TimeSubs"
Dim timetemp
Dim ONEreminder, ONEmins, ONEhours
Dim TWOreminder, TWOmins, TWOhours
Dim THREEreminder, THREEmins, THREEhours
Dim FOURreminder, FOURmins, FOURhours
Dim FIVEreminder, FIVEmins, FIVEhours
Dim SIXreminder, SIXmins, SIXhours
Dim SEVENreminder, SEVENmins, SEVENhours
Dim EIGHTreminder, EIGHTmins, EIGHThours
Dim NINEreminder, NINEmins, NINEhours
Dim TENreminder, TENmins, TENhours

Dim CorrectDay As Boolean
Dim DayToRemind As Integer

Public Sub TCOMPARE(TimeType As String, CurrentTime) ' Checks to see if user needs to be reminded of the time
    timetemp = Format$(CurrentTime, "H:MM")
    
    Select Case TimeType
    Case "Q"
        If Mid(timetemp, 3, 2) = "15" Or Mid(timetemp, 3, 2) = 15 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "15" Or Mid(timetemp, 4, 2) = 15 Then TimeSubs.Remind
        If Mid(timetemp, 3, 2) = "30" Or Mid(timetemp, 3, 2) = 30 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "30" Or Mid(timetemp, 4, 2) = 30 Then TimeSubs.Remind
        If Mid(timetemp, 3, 2) = "45" Or Mid(timetemp, 3, 2) = 45 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "45" Or Mid(timetemp, 4, 2) = 45 Then TimeSubs.Remind
        If Mid(timetemp, 3, 2) = "00" Or Mid(timetemp, 3, 2) = 0 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "00" Or Mid(timetemp, 4, 2) = 0 Then TimeSubs.Remind
    Case "H"
        If Mid(timetemp, 3, 2) = "30" Or Mid(timetemp, 3, 2) = 30 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "30" Or Mid(timetemp, 4, 2) = 30 Then TimeSubs.Remind
        If Mid(timetemp, 3, 2) = "00" Or Mid(timetemp, 3, 2) = 0 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "00" Or Mid(timetemp, 4, 2) = 0 Then TimeSubs.Remind
    Case "E"
        If Mid(timetemp, 3, 2) = "00" Or Mid(timetemp, 3, 2) = 0 Then TimeSubs.Remind
        If Mid(timetemp, 4, 2) = "00" Or Mid(timetemp, 4, 2) = 0 Then TimeSubs.Remind
    End Select
End Sub

Public Sub REMINDCOMPARE(Num As Integer, reminder, mins, hours) ' Checks to see if the user needs to be alerted of a reminder
    Debug.Print "Remind Compare sub invoked."
    
    timetemp = Format$(Time, "HHMM") ' Store the current time
    ARR = GetSetting("ComTalk", "Options", "ReadPopup", 0) ' See if the character should read out the text and popup a "XP" sidebox
    
    If Mid(hours, 1, 1) = " " Then hours = Mid(hours, 2)  ' -|
    If Mid(mins, 1, 1) = " " Then mins = Mid(mins, 2)     '  |- Fix up input data
    If Int(mins) < 10 Then mins = "0" & mins              ' -|
    
    DayToRemind = GetSetting("ComTalk", "Reminder" & Str(Num), "Date", 1)
    
    Debug.Print "Date to remind (" & Num & ") = " & DayToRemind
    CorrectDay = False ' Make sure the default is false
    
    Static CDays As Integer
    Static CMonths As Integer
    
    Select Case DayToRemind
    Case 1 ' Every Day, always allow
        CorrectDay = True
    Case 2 ' Weekday
        If IsWeekend(Now) = False Then CorrectDay = True
    Case 3 ' Weekend
        If IsWeekend(Now) = True Then CorrectDay = True
    Case 4 ' Custom Date
        CDays = GetSetting("ComTalk", "Reminder" & Str(Num), "CusDays", 0)
        CMonths = GetSetting("ComTalk", "Reminder" & Str(Num), "CusMonth", 0)
        
        Debug.Print "Reminder Custom Date - " & CDays & "/" & CMonths & " = " & Format(Now, "dd/mm")
        
        If Format(Now, "dd") = CDays And Format(Now, "mm") = CMonths Then CorrectDay = True
    End Select
    
    Debug.Print "REMIND TEST - " & Int(timetemp) & "=" & Int(hours & mins) & " (Correct Date = " & CorrectDay & ")"
    
    If CorrectDay = True Then
        If Int(timetemp) = Int(hours & mins) Then
            Debug.Print "REMIND - " & reminder
            Upop = GetSetting("ComTalk", "Options", "RemindPopup", 0)
            Debug.Print "Popup setting for custom reminder: " & Upop
            If Upop = 0 Then ' Just read reminder
                mainfrm.Agent1.Characters("Genie").Stop
                Action "Announce"
                SpeakText reminder
                Action "RestPose"
                mainfrm.customreminderscheck.Enabled = False
                mainfrm.waitenable.Enabled = True
            Else ' Read and show sidebox
                If ARR = 1 Then
                    Action "Announce"
                    SpeakText reminder
                    Action "RestPose"
                End If
                Static StrReminder As String
                StrReminder = reminder
                onnow.ShowBox StrReminder, "Reminder"
            End If
        End If
    End If
End Sub

Public Sub checkforcustomreminders() ' Checks the Custom Reminders, to see if one has elapsed
    On Error GoTo 0
    
    Debug.Print "Check for Custom Reminders sub invoked."
    
    GetCRFromReg
    
If ONEreminder <> "" And ONEmins <> "" And ONEhours <> "" Then
    REMINDCOMPARE 1, ONEreminder, Str(ONEmins), Str(ONEhours)
End If

If TWOreminder <> "" And TWOmins <> "" And TWOhours <> "" Then
    REMINDCOMPARE 2, TWOreminder, Str(TWOmins), Str(TWOhours)
End If

If THREEreminder <> "" And THREEmins <> "" And THREEhours <> "" Then
    REMINDCOMPARE 3, THREEreminder, Str(THREEmins), Str(THREEhours)
End If

If FOURreminder <> "" And FOURmins <> "" And FOURhours <> "" Then
    REMINDCOMPARE 4, FOURreminder, Str(FOURmins), Str(FOURhours)
End If

If FIVEreminder <> "" And FIVEmins <> "" And FIVEhours <> "" Then
    REMINDCOMPARE 5, FIVEreminder, Str(FIVEmins), Str(FIVEhours)
End If

If SIXreminder <> "" And SIXmins <> "" And SIXhours <> "" Then
    REMINDCOMPARE 6, SIXreminder, Str(SIXmins), Str(SIXhours)
End If

If SEVENreminder <> "" And SEVENmins <> "" And SEVENhours <> "" Then
    REMINDCOMPARE 7, SEVENreminder, Str(SEVENmins), Str(SEVENhours)
End If

If EIGHTreminder <> "" And EIGHTmins <> "" And EIGHThours <> "" Then
    REMINDCOMPARE 8, EIGHTreminder, Str(EIGHTmins), Str(EIGHThours)
End If

If NINEreminder <> "" And NINEmins <> "" And NINEhours <> "" Then
    REMINDCOMPARE 9, NINEreminder, Str(NINEmins), Str(NINEhours)
End If

If TENreminder <> "" And TENmins <> "" And TENhours <> "" Then
    REMINDCOMPARE 10, TENreminder, Str(TENmins), Str(TENhours)
End If

    Debug.Print "Finished checking for custom reminders."
End Sub

Public Sub GetCRFromReg() ' Loads the custom reminders from the registry
    ' This sub should only be called when the user changes the custom reminders,
    ' the program is loaded, or the user changes the character.
    
    Debug.Print "Get CR From Registry sub invoked."
    
    ONEreminder = GetSetting("ComTalk", "Reminder 1", "Reminder", "")
    ONEmins = GetSetting("ComTalk", "Reminder 1", "Mins", "")
    ONEhours = GetSetting("ComTalk", "Reminder 1", "Hours", "")
    TWOreminder = GetSetting("ComTalk", "Reminder 2", "Reminder", "")
    TWOmins = GetSetting("ComTalk", "Reminder 2", "Mins", "")
    TWOhours = GetSetting("ComTalk", "Reminder 2", "Hours", "")
    THREEreminder = GetSetting("ComTalk", "Reminder 3", "Reminder", "")
    THREEmins = GetSetting("ComTalk", "Reminder 3", "Mins", "")
    THREEhours = GetSetting("ComTalk", "Reminder 3", "Hours", "")
    FOURreminder = GetSetting("ComTalk", "Reminder 4", "Reminder", "")
    FOURmins = GetSetting("ComTalk", "Reminder 4", "Mins", "")
    FOURhours = GetSetting("ComTalk", "Reminder 4", "Hours", "")
    FIVEreminder = GetSetting("ComTalk", "Reminder 5", "Reminder", "")
    FIVEmins = GetSetting("ComTalk", "Reminder 5", "Mins", "")
    FIVEhours = GetSetting("ComTalk", "Reminder 5", "Hours", "")
    SIXreminder = GetSetting("ComTalk", "Reminder 6", "Reminder", "")
    SIXmins = GetSetting("ComTalk", "Reminder 6", "Mins", "")
    SIXhours = GetSetting("ComTalk", "Reminder 6 ", "Hours", "")
    SEVENreminder = GetSetting("ComTalk", "Reminder 7", "Reminder", "")
    SEVENmins = GetSetting("ComTalk", "Reminder 7", "Mins", "")
    SEVENhours = GetSetting("ComTalk", "Reminder 7", "Hours", "")
    EIGHTreminder = GetSetting("ComTalk", "Hours8", "Reminder", "")
    EIGHTmins = GetSetting("ComTalk", "Reminder 8", "Mins", "")
    EIGHThours = GetSetting("ComTalk", "Reminder 8", "Hours", "")
    NINEreminder = GetSetting("ComTalk", "Reminder 9", "Reminder", "")
    NINEmins = GetSetting("ComTalk", "Reminder 9", "Mins", "")
    NINEhours = GetSetting("ComTalk", "Reminder 9", "Hours", "")
    TENreminder = GetSetting("ComTalk", "Reminder 10", "Reminder", "")
    TENmins = GetSetting("ComTalk", "Reminder 10", "Mins", "")
    TENhours = GetSetting("ComTalk", "Reminder 10", "Hours", "")
End Sub

Public Sub Remind() ' Reminds the user of the time
    mainfrm.Timer2.Enabled = False
    mainfrm.Timer3.Enabled = True
    mainfrm.Agent1.Characters("Genie").Stop
    Action "Announce"
    SpeakText "The Time is " & Format(Now, "H:MM AM/PM")
    Action "restpose"
End Sub

Public Function IsWeekend(ByVal vntDate As Variant) As Boolean ' Returns true if input date is a weekend. Use "now" as an input
    Dim bResult         As Boolean
    If IsDate(vntDate) Then
        If (Weekday(vntDate) Mod 6 = 1) Then bResult = True Else bResult = False
    End If
    IsWeekend = bResult
End Function
