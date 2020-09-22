VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainfrm 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk Character"
   ClientHeight    =   2175
   ClientLeft      =   5790
   ClientTop       =   4995
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3510
   Begin MSComctlLib.ImageList Icons 
      Left            =   2520
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":075E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComTalk.chameleonButton command2 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      btype           =   3
      tx              =   "Cancel"
      enab            =   -1  'True
      font            =   "main.frx":0A7A
      coltype         =   3
      focusr          =   0   'False
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "main.frx":0AA6
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin ComTalk.USEFULLctrl UC 
      Left            =   3240
      Top             =   1680
      _extentx        =   609
      _extenty        =   1032
   End
   Begin VB.Timer customreminderscheck 
      Interval        =   30000
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer waitenable 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1320
      Top             =   840
   End
   Begin VB.Timer checkgone 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   840
   End
   Begin VB.Timer endtimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.FileListBox Character 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   1455
      Hidden          =   -1  'True
      Left            =   120
      Pattern         =   "*.acs"
      System          =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   2295
   End
   Begin ComTalk.chameleonButton command3 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1140
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      btype           =   3
      tx              =   "Preview"
      enab            =   -1  'True
      font            =   "main.frx":0AC4
      coltype         =   3
      focusr          =   0   'False
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "main.frx":0AF0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin ComTalk.chameleonButton command1 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      btype           =   3
      tx              =   "Select"
      enab            =   -1  'True
      font            =   "main.frx":0B0E
      coltype         =   3
      focusr          =   0   'False
      bcol            =   12632256
      bcolo           =   12632256
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "main.frx":0B3A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "main.frx":0B58
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label menufont 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin AgentObjectsCtl.Agent Agent2 
      Left            =   360
      Top             =   1200
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   480
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'                   ---------------------------------
'                 | COMTALK BY DEAN CAMERA, 2001-2004 |
'                   ---------------------------------
'                        (C) Dean Camera, 2004


'Agent Variables/Constants
Dim Genie As IAgentCtlCharacter
Dim DataPath As String
'Program Variables/Declares
Dim FastExit
Dim FontToUse
Dim changingchar As Boolean
Dim allowdiskcheck As Boolean
Dim BadListNum
Dim Greeting As String
Dim sg As String
Dim doremind
Dim NoLog
Dim tempx, tempy
Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long


Private Sub Agent1_ActiveClientChange(ByVal CharacterID As String, ByVal Active As Boolean)
    If Active = False Then
        MsgBox "The current character is being used by another application." & vbNewLine & "It is not recommended to use ComTalk while the current character is in use." & vbNewLine & "You should either close the other application or change the character.", vbExclamation + vbSystemModal, "ComTalk - Error"
    End If
End Sub

Private Sub Agent1_Bookmark(ByVal BookmarkID As Long)
    'When the intro is done, a bookmark is called. Re-enable all timer functions.
    mainfrm.Timer2.Enabled = True
    mainfrm.customreminderscheck = True
End Sub

Sub SideBoxButtonClick(ButtonName)
    
    If ButtonName = "Set Options" Then
        Options.Show
    End If
    
    If ButtonName = "En-Tech Website" Then
        Shell "explorer.exe http://www.en-tech.i8.com", vbNormalFocus
    End If
    
End Sub

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If Button = 1 Then
        If Genie.Visible = True Then
            If FRunning = True Then
                FRunning = False
                customreminderscheck.Enabled = True
                Timer2.Enabled = True
            End If
            
            Genie.StopAll
            Action "restpose"
            If endtimer.Enabled = True Then KillApp ' User told ComTalk to quit, but clicked on the character and it stopped, so endtimer cannot work as the character is still shown
        End If
    Else
        
        Unload CustomMenu ' Force menu refresh
        Load CustomMenu   '-------/
        
        If Genie.Visible = True Then
            If UseComMenu = True Then
                Agent1.Characters("Genie").AutoPopupMenu = False
                Declairs.SetForegroundWindow Me.hwnd
                CustomMenu.ShowComputerMenu
                Declairs.SetForegroundWindow Me.hwnd
                Exit Sub
            End If
            Agent1.Characters("Genie").AutoPopupMenu = False
            Declairs.SetForegroundWindow Me.hwnd
            If PopupDLLFail = False Then
                Debug.Print "-- New Menu Popup"
                CustomMenu.CusMenu.XPPopUpMenu 1
            Else
                Debug.Print "-- Old Menu Popup"
                PopupMenu CustomMenu.MenuCmds
            End If
            Declairs.SetForegroundWindow Me.hwnd
        Else
            Agent1.Characters("Genie").AutoPopupMenu = False
            Declairs.SetForegroundWindow Me.hwnd
            CustomMenu.ShowCharMenu
            Declairs.SetForegroundWindow Me.hwnd
        End If
    End If
End Sub

Private Sub Agent1_Command(ByVal UserInput As Object)
    
    On Error Resume Next
    
    If UserInput.Name = "About" Then
        CustomMenu.MenuCommandClick "MNUCTAbout"
    End If
    
    If UserInput.Name = "CReminders" Then
        Customreminders.Show
    End If
    
    If UserInput.Name = "SSaver" Then
        Call SendMessage(Me.hwnd, &H112&, &HF140&, 0&)
    End If
    
    If UserInput.Name = "Close" Then
        Genie.StopAll
        endme
    End If
    
    If UserInput.Name = "ReadClip" Then
        CustomMenu.MenuCommandClick "MNURClip"
    End If
    
    If UserInput.Name = "SayTime" Then
        CustomMenu.MenuCommandClick "MNUSTime"
    End If
    
    If UserInput.Name = "SayDate" Then
        CustomMenu.MenuCommandClick "MNUSDate"
    End If
    
    'Log Request
    LoggingSubs.Req UserInput.Name
    
    If UserInput.Name = "HideMe" Then
        Genie.Hide
    End If
    
    If UserInput.Name = "CC" Then
        Character.Selected(0) = True
        mainfrm.Show
    End If
    
    If UserInput.Name = "Options" Then
        Options.Show
    End If
    
    If UserInput.Name = "VoiceA" Then
        vcommands.Show
    End If
    
    If UserInput.Name = "RepeatTXT" Then
        SpeakText "\lst\"
    End If
    
    GetCAFromReg ' Get voice commands from registry
    
    If UserInput.Name = ONEname Then
        Shell ONEpath, vbNormalFocus
    End If
    
    If UserInput.Name = TWOname Then
        Shell TWOpath, vbNormalFocus
    End If
    
    If UserInput.Name = THREEname Then
        Shell THREEpath, vbNormalFocus
    End If
    
    If UserInput.Name = FOURname Then
        Shell FOURpath, vbNormalFocus
    End If
    
    If UserInput.Name = FIVEname Then
        Shell FIVEpath, vbNormalFocus
    End If
    
    If UserInput.Name = SIXname Then
        Shell SIXpath, vbNormalFocus
    End If
    
    If UserInput.Name = SEVENname Then
        Shell SEVENpath, vbNormalFocus
    End If
    
    If UserInput.Name = EIGHTname Then
        Shell EIGHTpath, vbNormalFocus
    End If
    
    If UserInput.Name = NINEname Then
        Shell NINEpath, vbNormalFocus
    End If
    
    If UserInput.Name = TENname Then
        Shell TENpath, vbNormalFocus
    End If
    
    If UserInput.Name = "SS" Then
        ss.Show
    End If
    
    If UserInput.Name = "ShutDown" Then
        CustomMenu.MenuCommandClick "MNUStDwn"
    End If
    
    If UserInput.Name = "Restart" Then
        CustomMenu.MenuCommandClick "MNURStart"
    End If
    
    If UserInput.Name = "LogOff" Then
        CustomMenu.MenuCommandClick "MNULogOff"
    End If
    
    If UserInput.Name = "Actions" Then
        actions.Show
    End If
    
    If UserInput.Name = "SWRun" Then
        CustomMenu.MenuCommandClick "MNUWRuntime"
    End If
    
    If UserInput.Name = "notepad" Then
        CustomMenu.MenuCommandClick "MNUNPad"
    End If
    
    If UserInput.Name = "explorer" Then
        CustomMenu.MenuCommandClick "MNUExplorer"
    End If
    
    If UserInput.Name = "srecorder" Then
        CustomMenu.MenuCommandClick "MNUSRecorder"
    End If
    
    If UserInput.Name = "calc" Then
        CustomMenu.MenuCommandClick "MNUCalc"
    End If
    
    If UserInput.Name = "defrag" Then
        CustomMenu.MenuCommandClick "MNUDefrag"
    End If
    
    If UserInput.Name = "CloY" Then
        EndLog
        ExitWindowsEx EWX_LogOff, 0&
    End If
    
    If UserInput.Name = "CloN" Then
        Genie.Commands.RemoveAll
        populatecommands
        Genie.Commands.GlobalVoiceCommandsEnabled = True
        CustomMenu.MenuCommandClick "MNUnLO"
    End If
    
    If UserInput.Name = "CrsY" Then
        EndLog
        ExitWindowsEx EWX_REBOOT, 0&
    End If
    
    If UserInput.Name = "CrsN" Then
        Genie.Commands.RemoveAll
        populatecommands
        Genie.Commands.GlobalVoiceCommandsEnabled = True
        CustomMenu.MenuCommandClick "MNUnRS"
    End If
    
    If UserInput.Name = "CsdY" Then
        EndLog
        ExitWindowsEx EWX_SHUTDOWN, 0&
    End If
    
    If UserInput.Name = "CsdN" Then
        Genie.Commands.RemoveAll
        populatecommands
        Genie.Commands.GlobalVoiceCommandsEnabled = True
        CustomMenu.MenuCommandClick "MNUnSD"
    End If
    
    If UserInput.Name = "LockWS" Then LockWorkstation.LockStation
    
    If UserInput.Name = "OCD" Then UC.ACTN_WIN_SetCDState TrayOpen
    
    If UserInput.Name = "CCD" Then UC.ACTN_WIN_SetCDState TrayClosed
    
    If UserInput.Name = "SAC" Then GetAtomicClock
    
    If UserInput.Name = "CCV" Then iUpdate.GetNewVer
    
    If UserInput.Name = "SIBPP" Then SetProgPaths.Show
    
    If UserInput.Name = "MSAgentOptions" Then CustomMenu.MenuCommandClick ("MNUMSAgentOptions")
    
    If Mid(UserInput.Name, 1, 4) = "RES_" Then Resolution.ChangeRes Mid(UserInput.Name, 5)
    
    Exit Sub
    
errordes:
End Sub

Private Sub Agent1_IdleStart(ByVal CharacterID As String)
    IdlingChar
End Sub

Private Sub Agent1_Move(ByVal CharacterID As String, ByVal X As Integer, ByVal Y As Integer, ByVal Cause As Integer)
    If CharacterID = "Genie" Then
        SaveSetting "ComTalk", "Options", "X", X
        SaveSetting "ComTalk", "Options", "Y", Y
    End If
End Sub

Private Sub checkgone_Timer()
    If Agent1.Characters("Genie").Visible = False Then ' Wait until character hidden before loading new character
        checkgone.Enabled = False
        Agent1.Characters.Unload "Genie"
        LoggingSubs.ChangingCharacter
        changingchar = True
        
        LoadComTalk True
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    
    For I = 1 To CharactersList.ListItems.Count
        If CharactersList.ListItems.Item(I).Selected = True Then Character.Selected(I - 1) = True
    Next
    
    SaveSetting "ComTalk", "Options", "MyCharacter", Character.FileName
    Me.Hide
    tempx = Genie.Left
    tempy = Genie.Top
    Genie.Hide
    checkgone.Enabled = True
End Sub

Private Sub command2_Click()
    Me.Hide
End Sub

Private Sub Command3_Click()
    For I = 1 To CharactersList.ListItems.Count
        If CharactersList.ListItems.Item(I).Selected = True Then Character.Selected(I - 1) = True
    Next
    
    On Error Resume Next
    Temp = Len(Character.FileName)
    If Character.FileName <> Mid(Declairs.MyCharacter, 1, Temp) Then
        Agent2.Characters.Load "PrevChar", Declairs.GetWindowsDir & "\Msagent\Chars\" & Character.FileName
        Character.Enabled = False
        Agent2.Characters("PrevChar").Show
        Agent2.Characters("PrevChar").Hide
        command1.Enabled = False
        command2.Enabled = False
        command3.Enabled = False
        Timer1.Enabled = True
    End If
End Sub

Private Sub customreminderscheck_Timer()
    On Error GoTo CusRemFail
    checkforcustomreminders
    Exit Sub
CusRemFail:
    SpeakError "Location: Custom Reminder Timer Sub. ComTalk can function without this sub, but custom reminders will be disabled. Error Information: Number: " & Err.Number & ". Description: " & Err.Description & "."
    customreminderscheck.Enabled = False
End Sub

Private Sub endtimer_Timer()
    If Genie.Visible = False Then ' Wait until character is hidden before quitting
        KillApp ' Close ComTalk
    End If
End Sub

Sub KillApp()
    On Error Resume Next
    Agent1.Characters.Unload "Genie"
    LoggingSubs.EndLog
    Dim I As Integer
    While Forms.Count > 1
        I = 0
        While Forms(I).Caption = Me.Caption
            I = I + 1
        Wend
        
        Unload Forms(I) ' Unload forms
    Wend
End ' Kill application
End Sub

Private Sub Form_Load()
    OpenTempFile
    
    Load CustomMenu ' Load the XP style menu
    
    LoadComTalk 'Load main ComTalk sub to load the Character
    ReTrans Me 'Make top of form rounded
    
    
    '---- Test Statements ----
    '-------------------------
End Sub

Sub LoadComTalk(Optional JustChangedChar As Boolean)
    
    If NoLockDown = True Then ' LockDown command diabled on CommandLine
        CustomMenu.MNULDown.Visible = False
        Options.Text1.Enabled = False
        Options.Text2.Enabled = False
    End If
    
    PrintTempFile "ComTalk Load> DECODE PASSWORD"
    DecodePassword 'Decode ComTalk password from the registry
    
    FastExit = GetSetting("ComTalk", "Options", "UseFastExit", 0)
    FontToUse = menufont.Font
    
    PrintTempFile "ComTalk Load> REGS"
    TimeSubs.GetCRFromReg 'Get Custom Reminders
    Declairs.GetCAFromReg 'Get Custom Actions
    SIMenu = GetSetting("ComTalk", "Options", "ShowInternetMenu", 1) 'Check to see if the Internet menu should be shown
    If SIMenu = 1 Then CustomMenu.MNUInet.Visible = True Else CustomMenu.MNUInet.Visible = False 'HIDE/SHOW Internet menu
    
    PrintTempFile "ComTalk Load> CHANGING CHAR CHECK"
    If changingchar = False Then
        UseLog = GetSetting("ComTalk", "Options", "LogEvents", 0) 'Check to see if ComTalk should log events
        If NoLog = False Then
            unexpected = GetSetting("ComTalk", "Program", "IsOpen", 0) 'Check for an unexpected session ending in the log (ComTalk crashed)
            SaveSetting "ComTalk", "Program", "IsOpen", 1 'Tell plugins ComTalk is open
            BeginLog
        Else
            changingchar = False
            ChangingCharacter
        End If
    Else
        NoLog = False
    End If
    
    If GetSetting("ComTalk", "Options", "CheckDiskSpace", 1) = 1 Then 'See if ComTalk should check Disk Space on C drive
        allowdiskcheck = True
    End If
    
    PrintTempFile "ComTalk Load> PLUGINS"
    PlugIns.LoadPlugs 'Load plugins
    
    PrintTempFile "ComTalk Load> REMIND VARS"
    doremind = GetSetting("ComTalk", "Options", "RemindTime", 0) 'Remind user of the time
    RemindTVal = GetSetting("ComTalk", "Options", "RemindTimeVal", "Quarter") 'When ComTalk should remind user of the time
    
    PrintTempFile "ComTalk Load> PREV INST"
    Me.Hide
    DataPath = Declairs.GetWindowsDir & "\msagent\chars\"
    Agent1.RaiseRequestErrors = False
    If App.PrevInstance = True Then 'Check if ComTalk is already open
        frmSplash.Hide
        MsgBox "Error - ComTalk is already open.", vbCritical, "ComTalk - Error"
        killprogram 'Don't let the program run more than once at a time.
    End If
    
    'Set Variables
    PrintTempFile "ComTalk Load> VARS"
    GoCharacter = True
    clicklistnum = 1
    UserName = Declairs.MyNameToRead 'Get the user's name
    If Format(Now, "HH") > 11 And Format(Now, "HH") < 18 Then
        Greeting = "Good Afternoon, " & UserName
    ElseIf Format(Now, "HH") > 17 Then
        Greeting = "Good Evening, " & UserName
    Else
        Greeting = "Good Morning, " & UserName
    End If
    
    PrintTempFile "ComTalk Load> BADLIST"
    BadListNum = GetSetting("ComTalk", "BadList", "BadTotal", 0)
    If BadListNum <> 0 Then 'Check for disabled characters
        For I = 1 To BadListNum
            Temp = GetSetting("ComTalk", "BadList", I)
            Temp = UCase(Temp)
            If UCase(Declairs.MyCharacter) = Temp Then
                frmSplash.Hide
                MsgBox "Current Character has been added to BadList. Please select a new character.", vbCritical, "ComTalk - Error"
                Agent1.Characters.Load "Genie"
                Set Genie = Agent1.Characters("Genie")
                Character.Path = Declairs.GetWindowsDir & "msagent\chars\"
                mainfrm.Show
                Exit Sub
            End If
        Next I
    End If
    
    PrintTempFile "ComTalk Load> LOAD CHAR"
    Agent1.Characters.Load "Genie", Declairs.MyCharacter 'Load character
    PrintTempFile "ComTalk Load> SET CHAR"
    Set Genie = Agent1.Characters("Genie") 'Put character into variable for easy access
    PrintTempFile "ComTalk Load> LOAD SR"
    Label1.Caption = Agent1.Characters("Genie").SRStatus 'Initialize Speech Recognition
    PrintTempFile "ComTalk Load> HIDE POP"
    Agent1.Characters("Genie").AutoPopupMenu = False 'Disable in-built popup menu
    
    PrintTempFile "ComTalk Load> POPULATE COMMANDS"
    populatecommands 'Add Commands to popup menu
    
    Agent1.Characters("Genie").Commands.FontName = FontToUse
    Agent1.Characters("Genie").Commands.Voice = "ComTalk"
    Agent1.Characters("Genie").Listen False 'Don't use Speech Recognition until fully loaded
    
    PrintTempFile "ComTalk Load> OTHERCLIENT CHECK"
    If Agent1.Characters("Genie").HasOtherClients = True Then 'Check to see if other programs are using the same character
        frmSplash.Hide
        MsgBox "The current character is being used by another application." & vbNewLine & "It is not recommended to use ComTalk while the current character is in use." & vbNewLine & "You should either close the other application or change the character.", vbExclamation, "ComTalk - Error"
    End If
    
    PrintTempFile "ComTalk Load> SOUND EFFECTS"
    Genie.Commands.Caption = "ComTalk"
    Genie.Commands.VoiceCaption = "ComTalk"
    Genie.Commands.Voice = True
    Character.Path = Declairs.GetWindowsDir & "msagent\chars\"
    usesnd = GetSetting("ComTalk", "Options", "UseCharacterSNDs", 1)
    If usesnd = 1 Then 'Check to see if Character's sound effects should be enabled
        Genie.SoundEffectsOn = True
    Else
        Genie.SoundEffectsOn = False
    End If
    
    PrintTempFile "ComTalk Load> SIZE BUBBLE"
    If GetSetting("ComTalk", "Options", "SizeBubbleToText", 1) <> 1 Then 'Get character speech bubble size
        Genie.Balloon.Style = Genie.Balloon.Style And (Not SizeToText)
    Else
        Genie.Balloon.Style = Genie.Balloon.Style Or SizeToText
    End If
    
    PrintTempFile "ComTalk Load> BUBBLE STYLE"
    If GetSetting("ComTalk", "Options", "ShowSpeechBubble", 1) <> 1 Then 'Get character speech bubble style
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not BalloonOn)
    Else
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or BalloonOn
    End If
    
    If GetSetting("ComTalk", "Options", "AutoWordPace", 1) <> 1 Then
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not AutoPace)
    Else
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or AutoPace
    End If
    
    PrintTempFile "ComTalk Load> POSITIONS"
    'Load Positions
    Genie.Left = GetSetting("ComTalk", "Options", "X", 10)
    Genie.Top = GetSetting("ComTalk", "Options", "Y", 10)
    'Set Height
    Genie.Height = GetSetting("ComTalk", "Options", "CustomCharHeight", mainfrm.Agent1.Characters("Genie").OriginalHeight)
    Genie.Width = GetSetting("ComTalk", "Options", "CustomCharWidth", mainfrm.Agent1.Characters("Genie").OriginalWidth)
    
    sg = GetSetting("ComTalk", "Options", "SayNameOnStartup", 1)
    pitchnum = GetSetting("ComTalk", "Options", "CustomPitch", Agent1.Characters("Genie").Pitch)
    speednum = GetSetting("ComTalk", "Options", "CustomSpeed", Agent1.Characters("Genie").Speed)
    
    useidle = GetSetting("ComTalk", "Options", "UseIdleAnimations", 1)
    If useidle = 1 Then mainfrm.Agent1.Characters("Genie").IdleOn = True
    If useidle = 0 Then mainfrm.Agent1.Characters("Genie").IdleOn = False
    
    On Error Resume Next
    
    PrintTempFile "ComTalk Load> POPUP MENU"
    Agent1.Characters("Genie").AutoPopupMenu = False
    
    PrintTempFile "ComTalk Load> E/D COMMANDS"
    DisableCommands.EnableDisableCommands
    
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3 'Make form stay on top
    
    PrintTempFile "ComTalk Load> SHOW CHAR"
    Genie.Show    'Show Character
    
    If JustChangedChar = True Then
        Debug.Print "FIRST RUN - Fixing Speed and Pitch" 'Fix speed and pitch settings on the first run (stop's null values, making character talk like a Chipmunk!)
        SaveSetting "ComTalk", "Options", "CustomSpeed", mainfrm.Agent1.Characters("Genie").Speed
        SaveSetting "ComTalk", "Options", "CustomPitch", mainfrm.Agent1.Characters("Genie").Pitch
    End If
    
    mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not BalloonOn)
    SpeakText "\pau=1\" 'IMPORTANT - CAUSES MSAGENT TO WAIT UNTIL CHARACTER IS SHOWED BEFORE SPEAKING, AND FIXES SPEED AND PITCH
    Unload Options
    Load Options
    mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or BalloonOn 'Set the speech balloon style
    
    If FirstRun.IsFirstRun = True Then
        FirstRun.FirstRunText ' Speaks the intro test if first run of ComTalk
    Else
        CheckForOldEncMethod ' Checks for the old (ASCII Shift) encription method
    End If
    
    If sg = 1 Then 'Speak Greeting
        Action "Greet"
        SpeakText Greeting
        Action "restpose"
    End If
    
    startupMval = GetSetting("ComTalk", "Options", "StartupMessage", "N")
    startupmcheck 'Check to see if Time and/or Date should be announced on startup
    
    PrintTempFile "ComTalk Load> DISK SPACE"
    checkdiskspace 'Check disk space
    
    PrintTempFile "ComTalk Load> POPULATE ACTIONS"
    actions.PopulateActions ' Add the characters actions tot he action list
    
    AllowSpashHide = True
    
    LoggingSubs.CloseTempFile ' Kill temp file
    
    DoEvents
    Exit Sub
    
logerror:
    AddLogText "ERROR:" & Err.Description
    DoEvents
End Sub

Sub populatecommands() 'Add the inbuilt commands
    Genie.Commands.RemoveAll
    Agent1.Characters("Genie").Commands.Caption = "ComTalk"      '\ Changes the
    Agent1.Characters("Genie").Commands.VoiceCaption = "ComTalk" '/ Global Caption
    Genie.Commands.Add "SayTime", "Say Time", "What is the Time?", True, True
    Genie.Commands.Add "SayDate", "Say Date", "What is the Date?", True, True
    Genie.Commands.Add "ReadClip", "Read Clipboard", "Read Clipboard", True, True
    Genie.Commands.Add "SS", "Say Somthing", "Say Somthing", True, True
    Genie.Commands.Add "CC", "Change Character", "Change Character", True, True
    Genie.Commands.Add "Options", "Options", "Options", True, True
    Genie.Commands.Add "Actions", "Actions", "Actions", True, True
    Genie.Commands.Add "VoiceA", "Voice Actions", "Voice Actions", True, True
    Genie.Commands.Add "CReminders", "Custom Reminders", "Custom Reminders", True, True
    Genie.Commands.Add "notepad", "Notepad", "Notepad", True, True
    Genie.Commands.Add "srecorder", "Sound Recorder", "Sound Recorder", True, True
    Genie.Commands.Add "calc", "Calculator", "Calculator", True, True
    Genie.Commands.Add "defrag", "Defragmentor", "Defrag", True, True
    Genie.Commands.Add "OCD", "Open CD Drive", "Open CD Drive", True, True
    Genie.Commands.Add "CCD", "Close CD Drive", "Close CD Drive", True, True
    Genie.Commands.Add "explorer", "Explorer", "Explorer", True, True
    Genie.Commands.Add "Close", "Exit ComTalk", "Exit ComTalk", True, True
    Genie.Commands.Add "ShutDown", "Shut Down", "Shut Down Computer", True, True
    Genie.Commands.Add "Restart", "Restart", "Restart Computer", True, True
    Genie.Commands.Add "LogOff", "Log Off", "Log Off Computer", True, True
    If NoLockDown = False Then ' Don't load command if command line parameter to disable it is used
        Genie.Commands.Add "LockWS", "Lock Workstation (Win 9x)", "Lock Workstation", True, True
    End If
    Genie.Commands.Add "SWRun", "Windown Runtime (Mins)", "How long has Windows been running?", True, True
    Genie.Commands.Add "About", "About", "About ComTalk", True, True
    Genie.Commands.Add "SAC", "Sync Clock To Atomic Clock", "Sync Clock", True, True
    Genie.Commands.Add "SIBPP", "Set Built-In Program List", "Set Built In Program Paths", True, True
    Genie.Commands.Add "MSAgentOptions", "MSAgent Options", "Set Agent Options", True, True
    
    Genie.Commands.Add "RES_1600x1200", "1600x1200", "Res One-Six-Hundred ... One-Two-Hundred"
    Genie.Commands.Add "RES_1280x1024", "1280x1024", "Res One-Two-Eighty ... One-Oh-Two-Four"
    Genie.Commands.Add "RES_1152x864", "1152x864", "Res One-One-Five-Two ... Eight-Six-Four"
    Genie.Commands.Add "RES_1024x768", "1024x768", "Res One-Oh-Two-Four ... Seven-Six-Eight"
    Genie.Commands.Add "RES_800x600", "800x600", "Res Eight-Hundred ... Six-Hundred"
    Genie.Commands.Add "RES_640x480", "640x480", "Res Six-Fourty ... Four-Eighty"
    
    ExTemp = GetSetting("ComTalk", "Options", "ExtendedVoice", 0)
    If ExTemp = 1 Then CmdExt.ExtendCommands
    
    populateVcommands 'Add custom commands
End Sub

Sub startupmcheck() 'Check for startup Time/Date greeting
    startupMval = GetSetting("ComTalk", "Options", "StartupMessage", "N")
    
    If startupMval = "N" Then ' Nothing
        
    ElseIf startupMval = "T" Then  ' Say Time
        CustomMenu.MenuCommandClick "MNUSTime"
    ElseIf startupMval = "D" Then  ' Say Date
        CustomMenu.MenuCommandClick "MNUSDate"
    ElseIf startupMval = "TD" Then ' Say Time and Date
        CustomMenu.MenuCommandClick "MNUSTime"
        CustomMenu.MenuCommandClick "MNUSDate"
    End If
End Sub

Sub populateVcommands() 'Add custom commands 1-10
    GetCAFromReg ' Load
    CmdExt.ExtendedMenu False
    
    On Error Resume Next
    
    CustomMenu.MNUONE.Visible = False
    CustomMenu.MNUTWO.Visible = False
    CustomMenu.MNUTHREE.Visible = False
    CustomMenu.MNUFOUR.Visible = False
    CustomMenu.MNUFIVE.Visible = False
    CustomMenu.MNUSIX.Visible = False
    CustomMenu.MNUSEVEN.Visible = False
    CustomMenu.MNUEIGHT.Visible = False
    CustomMenu.MNUNINE.Visible = False
    CustomMenu.MNUTEN.Visible = False
    
    
    If ONEname <> "" Then
        If ONEpath <> "" Then
            If ONEcommand <> "" Then
                Genie.Commands.Add ONEname, ONEname, ONEcommand
                CustomMenu.MNUONE.Caption = ONEcommand
                CustomMenu.MNUONE.Visible = True
            End If
        End If
    End If
    
    If TWOname <> "" Then
        If TWOpath <> "" Then
            If TWOcommand <> "" Then
                Genie.Commands.Add TWOname, TWOname, TWOcommand
                CustomMenu.MNUTWO.Caption = TWOcommand
                CustomMenu.MNUTWO.Visible = True
            End If
        End If
    End If
    
    If THREEname <> "" Then
        If THREEpath <> "" Then
            If THREEcommand <> "" Then
                Genie.Commands.Add THREEname, THREEname, THREEcommand
                CustomMenu.MNUTHREE.Caption = THREEcommand
                CustomMenu.MNUTHREE.Visible = True
            End If
        End If
    End If
    
    If FOURname <> "" Then
        If FOURpath <> "" Then
            If FOURcommand <> "" Then
                Genie.Commands.Add FOURname, FOURname, FOURcommand
                CustomMenu.MNUFOUR.Caption = FOURcommand
                CustomMenu.MNUFOUR.Visible = True
            End If
        End If
    End If
    
    If FIVEname <> "" Then
        If FIVEpath <> "" Then
            If FIVEcommand <> "" Then
                Genie.Commands.Add FIVEname, FIVEname, FIVEcommand
                CustomMenu.MNUFIVE.Caption = FIVEcommand
                CustomMenu.MNUFIVE.Visible = True
            End If
        End If
    End If
    
    If SIXname <> "" Then
        If SIXpath <> "" Then
            If SIXcommand <> "" Then
                Genie.Commands.Add SIXname, SIXname, SIXcommand
                CustomMenu.MNUSIX.Caption = SIXcommand
                CustomMenu.MNUSIX.Visible = True
            End If
        End If
    End If
    
    If SEVENname <> "" Then
        If SEVENpath <> "" Then
            If SEVENcommand <> "" Then
                Genie.Commands.Add SEVENname, SEVENname, SEVENcommand
                CustomMenu.MNUSEVEN.Caption = SEVENcommand
                CustomMenu.MNUSEVEN.Visible = True
            End If
        End If
    End If
    
    If EIGHTname <> "" Then
        If EIGHTpath <> "" Then
            If EIGHTcommand <> "" Then
                Genie.Commands.Add EIGHTname, EIGHTname, EIGHTcommand
                CustomMenu.MNUEIGHT.Caption = EIGHTcommand
                CustomMenu.MNUEIGHT.Visible = True
            End If
        End If
    End If
    
    If NINEname <> "" Then
        If NINEpath <> "" Then
            If NINEcommand <> "" Then
                Genie.Commands.Add NINEname, NINEname, NINEcommand
                CustomMenu.MNUNINE.Caption = NINEcommand
                CustomMenu.MNUNINE.Visible = True
            End If
        End If
    End If
    
    If TENname <> "" Then
        If TENpath <> "" Then
            If TENcommand <> "" Then
                Genie.Commands.Add TENname, TENname, TENcommand
                CustomMenu.MNUTEN.Caption = TENcommand
                CustomMenu.MNUTEN.Visible = True
            End If
        End If
    End If
End Sub

Private Sub Form_Terminate()
    Temp = GetSetting("ComTalk", "Program", "IsOpen", 0) 'Tell plugins ComTalk is closed
    If Temp = 1 Then
        SaveSetting "ComTalk", "Program", "IsOpen", 0
        EndLog
        SaveSetting "ComTalk", "Program", "IsOpen", 0
    End If
    
    PlugIns.KillPlugs 'Unload plugins
End Sub




Private Sub Timer1_Timer() 'Wait until the preview character is hidden before unloading it
    If Agent2.Characters("PrevChar").Visible = False Then
        Character.Enabled = True
        command1.Enabled = True
        command2.Enabled = True
        command3.Enabled = True
        Agent2.Characters.Unload "PrevChar"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
    On Error GoTo AlertError
    'This sub is called every 30 seconds. It checks the current time
    'and disk space.
    
    Place = "Check Disk Space Command"
    checkdiskspace
    
    Place = "Time Reminder Check Commands"
    If doremind = 1 Then
        If RemindTVal = "Quarter" Then
            TCOMPARE "Q", Time
        ElseIf RemindTVal = "Half" Then
            TCOMPARE "H", Time
        Else
            TCOMPARE "E", Time
        End If
    End If
    Exit Sub
    
AlertError:
    SpeakError "Location: Main Timer Sub. Command: " & Place & ". ComTalk can function without this sub, but Disk Space Check and Time Alert will be disabled. Error Information: Number: " & Err.Number & ". Description: " & Err.Description & "."
    Timer2.Enabled = False
End Sub

Sub reload()
    Form_Load
End Sub

Sub checkdiskspace()
    If CMDDDC = False Then
        If allowdiskcheck = True Then 'Disk Space Check has not been disabled in the options menu
            If DiskSubs.GetDiskInfo("C:\") < 50 Then 'See if drive C has less than 50 megs left
                If IveWarned = False Then
                    IveWarned = True
                    SpeakText "Warning! Drive C has less than 50 megabytes of disk space left!"
                    allowdiskcheck = False
                    Timer2.Enabled = False
                    Timer3.Enabled = True
                End If
            End If
        Else
            Debug.Print "DiskSpaceCheck Canceled."
        End If
    End If
End Sub

Sub endme() 'Exit ComTalk
    If FastExit = 1 Then 'Check for fast exit
        SaveSetting "ComTalk", "Program", "IsOpen", 0
        killprogram
    Else
        SaveSetting "ComTalk", "Program", "IsOpen", 0
        Genie.StopAll
        Genie.Hide
        endtimer.Enabled = True
    End If
End Sub

Private Sub Timer3_Timer()
    'Stops the character from reminding the user every 30 secs.
    Timer2.Enabled = True
    Timer3.Enabled = False
End Sub


Private Sub waitenable_Timer()
    customreminderscheck.Enabled = True
    waitenable.Enabled = False
End Sub
