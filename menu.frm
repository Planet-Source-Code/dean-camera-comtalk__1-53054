VERSION 5.00
Begin VB.Form CustomMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu   - DO NOT SHOW THIS FORM"
   ClientHeight    =   795
   ClientLeft      =   5790
   ClientTop       =   5235
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add ComTalk commands to this form's menu to show it in the new XP style popup menu."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu MenuCmds 
      Caption         =   "MenuCommands"
      Begin VB.Menu MNUTaDsub 
         Caption         =   "Speech"
         Begin VB.Menu MNUSSomthing 
            Caption         =   "Say Somthing"
         End
         Begin VB.Menu MNURClip 
            Caption         =   "Read Clipboard"
         End
         Begin VB.Menu MNUSTime 
            Caption         =   "Say Time"
         End
         Begin VB.Menu MNUSDate 
            Caption         =   "Say Date"
         End
      End
      Begin VB.Menu MNUCptr 
         Caption         =   "Computer"
         Begin VB.Menu MNUStDwn 
            Caption         =   "Shut Down"
         End
         Begin VB.Menu MNURStart 
            Caption         =   "Restart"
         End
         Begin VB.Menu MNULogOff 
            Caption         =   "Log Off"
         End
      End
      Begin VB.Menu MNUWdows 
         Caption         =   "Windows"
         Begin VB.Menu MNUWRuntime 
            Caption         =   "Runtime"
         End
         Begin VB.Menu MNULDown 
            Caption         =   "LockDown"
         End
      End
      Begin VB.Menu MNUCTalk 
         Caption         =   "ComTalk"
         Begin VB.Menu MNUOptns 
            Caption         =   "Options"
         End
         Begin VB.Menu MNUMSAgentOptions 
            Caption         =   "MS Agent Options"
         End
         Begin VB.Menu MNUCTAbout 
            Caption         =   "About"
         End
         Begin VB.Menu MNUExit 
            Caption         =   "Exit"
         End
      End
      Begin VB.Menu MNUChtr 
         Caption         =   "Character"
         Begin VB.Menu MNUHideChar 
            Caption         =   "Hide"
         End
         Begin VB.Menu MNUActns 
            Caption         =   "Actions"
         End
         Begin VB.Menu MNUCReminders 
            Caption         =   "Custom Reminders"
         End
         Begin VB.Menu MNUCChange 
            Caption         =   "Change"
         End
      End
      Begin VB.Menu MNUCD 
         Caption         =   "CD Drive"
         Begin VB.Menu MNUCDOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu MNUCDClose 
            Caption         =   "Close"
         End
      End
      Begin VB.Menu MNUPGrams 
         Caption         =   "Programs"
         Begin VB.Menu MNUCalc 
            Caption         =   "Calculator"
         End
         Begin VB.Menu MNUDefrag 
            Caption         =   "Defragmentor"
         End
         Begin VB.Menu MNUNPad 
            Caption         =   "Notepad"
         End
         Begin VB.Menu MNUExplorer 
            Caption         =   "Explorer"
         End
         Begin VB.Menu MNUSRecorder 
            Caption         =   "Sound Recorder"
         End
         Begin VB.Menu MNUseperator3 
            Caption         =   "-"
         End
         Begin VB.Menu MNUConfigIBP 
            Caption         =   "Configure Paths..."
         End
      End
      Begin VB.Menu MNUCProgs 
         Caption         =   "Custom Programs"
         Begin VB.Menu MNUVActions 
            Caption         =   "Edit List..."
         End
         Begin VB.Menu MNUseperator 
            Caption         =   "-"
         End
         Begin VB.Menu MNUONE 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUTWO 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUTHREE 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUFOUR 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUFIVE 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUSIX 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUSEVEN 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUEIGHT 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUNINE 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUTEN 
            Caption         =   ""
            Visible         =   0   'False
         End
         Begin VB.Menu MNUCusEmpty 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MNUPlugIns 
         Caption         =   "Plugins"
         Begin VB.Menu MNUPlugEmpty 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
         End
         Begin VB.Menu MNUPlugs 
            Caption         =   "PLUG"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu MNUPSep 
            Caption         =   "-"
         End
         Begin VB.Menu MNUAPlugs 
            Caption         =   "About Plugins..."
         End
      End
      Begin VB.Menu MNUInet 
         Caption         =   "Internet"
         Begin VB.Menu MNUSyncAClock 
            Caption         =   "Sync System Time to Atomic Clock"
         End
      End
      Begin VB.Menu MNUseperator2 
         Caption         =   "-"
      End
      Begin VB.Menu MNUClose 
         Caption         =   "(Close)"
      End
   End
   Begin VB.Menu ConfirmMNUsd 
      Caption         =   "Confirm"
      Begin VB.Menu MNUySD 
         Caption         =   "Yes, Shutdown"
      End
      Begin VB.Menu MNUnSD 
         Caption         =   "No, Don't Shutdown"
      End
   End
   Begin VB.Menu ConfirmMNUrs 
      Caption         =   "Confirm"
      Begin VB.Menu MNUyRS 
         Caption         =   "Yes, Restart"
      End
      Begin VB.Menu MNUnRS 
         Caption         =   "No, Don't Restart"
      End
   End
   Begin VB.Menu ConfirmMNUlo 
      Caption         =   "Confirm"
      Begin VB.Menu MNUyLO 
         Caption         =   "Yes, LogOff"
      End
      Begin VB.Menu MNUnLO 
         Caption         =   "No, Don't LogOff"
      End
   End
   Begin VB.Menu NoCmdsMNU 
      Caption         =   "Hidden Commands"
      Begin VB.Menu MNUCShow 
         Caption         =   "Show"
      End
      Begin VB.Menu MNUExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "CustomMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

Public WithEvents CusMenu As cMenus
Attribute CusMenu.VB_VarHelpID = -1

Private Sub CusMenu_Click(ByVal Index As Long)
For Each Control In Me.Controls
If TypeOf Control Is Menu Then ' Find the menu name of the clicked menu index ("XP" style menu uses an index number, and dosen't acount for hidden menu items)
I = I + 1
If I = Index Then
NumOfVcmds = 10 - AddVCommandInt ' Find the number of Voice Commands used
If I > 38 + NumOfVcmds Then I = I + AddVCommandInt: I = I + 1 ' Algorithm to fix up index to account for hidden voice commands
If I < 50 + CustomMenu.MNUPlugs.Count + 1 And I > 50 Then
PlugIns.GoPlug MNUPlugs(I - 51).Caption, "ComTalkMenuClick"
Else
If I > (51 + MNUPlugs.Count) - 1 Then I = (I - (MNUPlugs.Count - 1)) ' This (and the other algorithms in this sub) took me HOURS to work out - they have to compensate for an unknown number of plugins,
MenuCommandClick CustomMenu.Controls(I).Name                         ' and voice commands. The new menu uses a "index" system rather than the menu name - and it dosn't add a index number for hidden items
End If                                                               ' which is a REAL pain - so this sub will compensate for it.
End If
End If
Next
End Sub

Private Sub Form_Load()
If NoLockDown = True Then Me.MNULDown.Visible = False

CmdExt.ExtendedMenu True
PlugIns.PutPlugsInMenu

On Error GoTo SayError
Set CusMenu = New cMenus

Me.Visible = False
Me.Height = 0
Me.Width = 0

DisableCommands.EnableDisableCommands

If PopupDLLFail = False Then
CusMenu.CreateFromForm Me
CusMenu.DrawStyle = mds_XP
End If
Exit Sub

SayError:
If PopupDLLFail = False Then
PopupDLLFail = True
MsgBox "There was an error initalising the CusMenus subclass. ComTalk will revert to the standard popup menu.", vbCritical, "ComTalk Popup Menu Fail"
End If
End Sub

Sub MenuCommandClick(MenuName As String)
Debug.Print "MENU CLICK: " & MenuName

Select Case MenuName

    Case "MNUActns"
            actions.Show
            LoggingSubs.Action "Actions"
    Case "MNUAPlugs"
            aboutplugs.Show
    Case "MNUCalc"
            On Error Resume Next
            PrgPath = GetSetting("ComTalk", "InBuiltPath", "Calculator", "C:\Windows\calc.exe")
            Shell PrgPath, vbNormalFocus
            LoggingSubs.Action "calc"
    Case "MNUCChange"
            mainfrm.Character.Selected(0) = True
            mainfrm.Show
            LoggingSubs.Action "CC"
    Case "MNUCDClose"
            mainfrm.UC.ACTN_WIN_SetCDState TrayClosed
            LoggingSubs.Action "CCD"
    Case "MNUCDOpen"
            mainfrm.UC.ACTN_WIN_SetCDState TrayOpen
            LoggingSubs.Action "OCD"
    Case "MNUConfigIBP"
            SetProgPaths.Show
            LoggingSubs.Action "CIBP"
    Case "MNUCReminders"
            Customreminders.Show
            LoggingSubs.Action "CReminders"
    Case "MNUCShow"
            If mainfrm.Agent1.Characters("Genie").Visible = False Then
            mainfrm.Agent1.Characters("Genie").Show
            End If
            LoggingSubs.Action "Show"
    Case "MNUCTAbout"
            onnow.DeanPic.Top = 240
            onnow.ShowBox "                     Name: Dean Camera" & vbmewline & _
                          "                               Nationality: Australian" & vbNewLine & _
                          "                     Age: 14", "ComTalk", "En-Tech Website"
            SpeakText "ComTalk is " & Chr(169) & " Dean Camera, En-Tech, \map=""Two thousand and four""=""2004""\."
            Action "Blink"
            SpeakText "This is version " & App.Major & "." & App.Minor & "." & App.Revision
            Action "Blink"
            LoggingSubs.Action "About"
    Case "MNUDefrag"
            On Error Resume Next
            PrgPath = GetSetting("ComTalk", "InBuiltPath", "Defrag", "C:\Windows\Defrag.exe")
            Shell PrgPath, vbNormalFocus
            LoggingSubs.Action "defrag"
    Case "MNUTEN"
            On Error Resume Next
            Shell TENpath, vbNormalFocus
    Case "MNUNINE"
            On Error Resume Next
            Shell NINEpath, vbNormalFocus
    Case "MNUEIGHT"
            On Error Resume Next
            Shell EIGHTpath, vbNormalFocus
    Case "MNUSEVEN"
            On Error Resume Next
            Shell SEVENpath, vbNormalFocus
    Case "MNUSIX"
            On Error Resume Next
            Shell SIXpath, vbNormalFocus
    Case "MNUFIVE"
            On Error Resume Next
            Shell FIVEpath, vbNormalFocus
    Case "MNUFOUR"
            On Error Resume Next
            Shell FOURpath, vbNormalFocus
    Case "MNUTHREE"
            On Error Resume Next
            Shell THREEpath, vbNormalFocus
    Case "MNUTWO"
            On Error Resume Next
            Shell TWOpath, vbNormalFocus
    Case "MNUONE"
            On Error Resume Next
            Shell ONEpath, vbNormalFocus
    Case "MNUExit"
            mainfrm.Agent1.Characters("Genie").StopAll
            mainfrm.endme ' Quit ComTalk
            LoggingSubs.Action "Close"
    Case "MNUExplorer"
            On Error Resume Next
            PrgPath = GetSetting("ComTalk", "InBuiltPath", "Explorer", "C:\Windows\Explorer.exe")
            Shell PrgPath, vbNormalFocus
            LoggingSubs.Action "explorer"
    Case "MNUHideChar"
            mainfrm.Agent1.Characters("Genie").Hide
            LoggingSubs.Action "HideMe"
    Case "MNULDown"
            LockWorkstation.LockStation
            LoggingSubs.Action "LockWS"
    Case "MNULogOff"
            UseComMenu = True
            ComMenu = 2
            mainfrm.Agent1.Characters("Genie").Commands.RemoveAll 'Show logoff commands
            mainfrm.Agent1.Characters("Genie").Commands.Add "CloY", "Yes, Log Off!", "Yes"
            mainfrm.Agent1.Characters("Genie").Commands.Add "CloN", "No, Don't Log Off!", "No"
            mainfrm.Agent1.Characters("Genie").Commands.GlobalVoiceCommandsEnabled = False
            SpeakText "Really Log Off?"
    Case "MNUMSAgentOptions"
            mainfrm.Agent1.PropertySheet.Visible = True
    Case "MNUnLO"
            LoggingSubs.Action "CloN"
            UseComMenu = False
            mainfrm.Agent1.Characters("Genie").Commands.RemoveAll
            mainfrm.populatecommands
    Case "MNUNPad"
            On Error Resume Next
            PrgPath = GetSetting("ComTalk", "InBuiltPath", "NotePad", "C:\Windows\Notepad.exe")
            Shell PrgPath, vbNormalFocus
            LoggingSubs.Action "notepad"
    Case "MNUnRS"
            LoggingSubs.Action "CrsN"
            UseComMenu = False
            mainfrm.Agent1.Characters("Genie").Commands.RemoveAll
            mainfrm.populatecommands
    Case "MNUnSD"
            UseComMenu = False
            LoggingSubs.Action "CsdN"
            mainfrm.Agent1.Characters("Genie").Commands.RemoveAll
            mainfrm.populatecommands
    Case "MNUOptns"
            Options.Show
            LoggingSubs.Action "Options"
    Case "MNURClip"
            LoggingSubs.Action "ReadClip"
            If Clipboard.GetText <> "" Then
            Action "read"
            SpeakText Clipboard.GetText
            Action "readreturn"
            Else
            Action "Confused"
            SpeakText "There is no text in the clipboard."
            Action "restpose"
            End If
    Case "MNURStart"
            ComMenu = 3
            UseComMenu = True
            mainfrm.Agent1.Characters("Genie").Commands.RemoveAll 'Show restart commands
            mainfrm.Agent1.Characters("Genie").Commands.Add "CrsY", "Yes, Restart!", "Yes"
            mainfrm.Agent1.Characters("Genie").Commands.Add "CrsN", "No, Don't Restart!", "No"
            mainfrm.Agent1.Characters("Genie").Commands.GlobalVoiceCommandsEnabled = False
            SpeakText "Really Restart?"
    Case "MNUSDate"
        LoggingSubs.Action "SayDate"

    Temp = Format(Now, "dddd, mmmm")
    Temp2 = Format(Now, "dd")
    Temp4 = Format(Now, "yyyy")

        Select Case Right(Val(Temp2), 1) 'Check ending
            Case "1"
                Temp3 = "st"
            Case "2"
                Temp3 = "nd"
            Case "3"
                Temp3 = "rd"
            Case 1
                Temp3 = "st"
            Case 2
                Temp3 = "nd"
            Case 3
                Temp3 = "rd"
            Case Else
                Temp3 = "th"
        End Select

        Select Case Right(Val(Temp2), 2) 'Check ending
            Case 11
                Temp3 = "th"
            Case 12
                Temp3 = "th"
            Case 13
                Temp3 = "th"
            Case "11"
                Temp3 = "th"
            Case "12"
                Temp3 = "th"
            Case "13"
                Temp3 = "th"
        End Select

    If Mid(Temp2, 1, 1) = "0" Then 'If Date is less than 10, strip the 0
        Temp2 = Mid(Temp2, 2, 1)
    End If

    Action "Announce"
        SpeakText "The Date is " & Temp & " the " & Temp2 & Temp3 & ", " & Temp4
    Action "restpose"
    
    Case "MNUSRecorder"
        On Error Resume Next
        PrgPath = GetSetting("ComTalk", "InBuiltPath", "SoundRecorder", "C:\Windows\Sndrec32.exe")
        Shell PrgPath, vbNormalFocus
        LoggingSubs.Action "srecorder"
    Case "MNUSSomthing"
        ss.Show
        LoggingSubs.Action "SS"
    Case "MNUStDwn"
        UseComMenu = True
        ComMenu = 1
        mainfrm.Agent1.Characters("Genie").Commands.RemoveAll 'Show ShutDown commands
        mainfrm.Agent1.Characters("Genie").Commands.Add "CsdY", "Yes, Shutdown!", "Yes"
        mainfrm.Agent1.Characters("Genie").Commands.Add "CsdN", "No, Don't Shutdown!", "No"
        mainfrm.Agent1.Characters("Genie").Commands.GlobalVoiceCommandsEnabled = False
        SpeakText "Really Shut Down?"
    Case "MNUSyncAClock"
        GetAtomicClock
    Case "MNUSTime"
        LoggingSubs.Action "SayTime"
        Action "Announce"
        SpeakText "The Time is " & Format(Now, "H:MM AM/PM")
        Action "restpose"
    Case "MNUVActions"
        vcommands.Show
        LoggingSubs.Action "VoiceA"
    Case "MNUWRuntime"
        LoggingSubs.Action "SWRun"
        SpeakText "Windows has been running for " & mainfrm.UC.PROP_WIN_WindowsRunTimeMinutes & " minutes."
    Case "MNUyLO"
        LoggingSubs.Action "CloY"
        EndLog
        ExitWindowsEx EWX_LogOff, 0&
    Case "MNUyRS"
        LoggingSubs.Action "CrsY"
        EndLog
        ExitWindowsEx EWX_REBOOT, 0&
    Case "MNUySD"
        LoggingSubs.Action "CsdY"
        EndLog
        ExitWindowsEx EWX_SHUTDOWN, 0&
    Case Else
        GoPlug MenuName, "MenuClick"
End Select
End Sub

Function AddVCommandInt() As Integer ' Add a value to the menu index depending on how many voice commands are shown
                                     '  ( the new menu uses an index property intead of menu names and does not include hidden menu items)
Static AddI As Integer
AddI = 0

If Me.MNUONE.Visible = False Then AddI = AddI + 1
If Me.MNUTWO.Visible = False Then AddI = AddI + 1
If Me.MNUTHREE.Visible = False Then AddI = AddI + 1
If Me.MNUFOUR.Visible = False Then AddI = AddI + 1
If Me.MNUFIVE.Visible = False Then AddI = AddI + 1
If Me.MNUSIX.Visible = False Then AddI = AddI + 1
If Me.MNUSEVEN.Visible = False Then AddI = AddI + 1
If Me.MNUEIGHT.Visible = False Then AddI = AddI + 1
If Me.MNUNINE.Visible = False Then AddI = AddI + 1
If Me.MNUTEN.Visible = False Then AddI = AddI + 1

AddVCommandInt = AddI
End Function

Sub ShowComputerMenu()
If PopupDLLFail = True Then
Select Case ComMenu
    Case 1
PopupMenu Me.ConfirmMNUsd
    Case 2
PopupMenu Me.ConfirmMNUlo
    Case 3
PopupMenu Me.ConfirmMNUrs
End Select
Else
Select Case ComMenu
    Case 1
CusMenu.XPPopUpMenu Me.ConfirmMNUsd
    Case 2
CusMenu.XPPopUpMenu Me.ConfirmMNUlo
    Case 3
CusMenu.XPPopUpMenu Me.ConfirmMNUrs
End Select
End If
End Sub

Sub ShowCharMenu()
If PopupDLLFail = False Then
CusMenu.XPPopUpMenu Me.NoCmdsMNU
Else
PopupMenu Me.NoCmdsMNU
End If
End Sub

Private Sub MNUActns_Click()
MenuCommandClick "MNUActns"
End Sub

Private Sub MNUAPlugs_Click()
MenuCommandClick "MNUAPlugs"
End Sub

Private Sub MNUCalc_Click()
MenuCommandClick "MNUCalc"
End Sub

Private Sub MNUCChange_Click()
MenuCommandClick "MNUCChange"
End Sub

Private Sub MNUCDClose_Click()
MenuCommandClick "MNUCDClose"
End Sub

Private Sub MNUCDOpen_Click()
MenuCommandClick "MNUCDOpen"
End Sub

Private Sub MNUConfigIBP_Click()
MenuCommandClick "MNUConfigIBP"
End Sub

Private Sub MNUCReminders_Click()
MenuCommandClick "MNUCReminders"
End Sub

Private Sub MNUCShow_Click()
MenuCommandClick "MNUCShow"
End Sub

Private Sub MNUCTAbout_Click()
MenuCommandClick "MNUCTAbout"
End Sub

Private Sub MNUDefrag_Click()
MenuCommandClick "MNUDefrag"
End Sub

Private Sub MNUEIGHT_Click()
MenuCommandClick "MNUEIGHT"
End Sub

Private Sub MNUExit_Click()
MenuCommandClick "MNUExit"
End Sub

Private Sub MNUExit2_Click()
MenuCommandClick "MNUExit"
End Sub

Private Sub MNUExplorer_Click()
MenuCommandClick "MNUExplorer"
End Sub

Private Sub MNUFIVE_Click()
MenuCommandClick "MNUFIVE"
End Sub

Private Sub MNUFOUR_Click()
MenuCommandClick "MNUFOUR"
End Sub

Private Sub MNUHideChar_Click()
MenuCommandClick "MNUHideChar"
End Sub

Private Sub MNULDown_Click()
MenuCommandClick "MNULDown"
End Sub

Private Sub MNULogOff_Click()
MenuCommandClick "MNULogOff"
End Sub

Private Sub MNUMSAgentOptions_Click()
MenuCommandClick "MNUMSAgentOptions"
End Sub

Private Sub MNUNINE_Click()
MenuCommandClick "MNUNINE"
End Sub

Private Sub MNUnLO_Click()
MenuCommandClick "MNUnLO"
End Sub

Private Sub MNUNPad_Click()
MenuCommandClick "MNUNPad"
End Sub

Private Sub MNUnRS_Click()
MenuCommandClick "MNUnRS"
End Sub

Private Sub MNUnSD_Click()
MenuCommandClick "MNUnSD"
End Sub

Private Sub MNUONE_Click()
MenuCommandClick "MNUONE"
End Sub

Private Sub MNUOptns_Click()
MenuCommandClick "MNUOptns"
End Sub

Private Sub MNUPlugs_Click(Index As Integer)
    PlugIns.GoPlug MNUPlugs(Index).Caption, "ComTalkMenuClick"
End Sub

Private Sub MNURClip_Click()
MenuCommandClick "MNURClip"
End Sub

Private Sub MNURStart_Click()
MenuCommandClick "MNURStart"
End Sub

Private Sub MNUSDate_Click()
MenuCommandClick "MNUSDate"
End Sub

Private Sub MNUSEVEN_Click()
MenuCommandClick "MNUSEVEN"
End Sub

Private Sub MNUSIX_Click()
MenuCommandClick "MNUSIX"
End Sub

Private Sub MNUSRecorder_Click()
MenuCommandClick "MNUSRecorder"
End Sub

Private Sub MNUSSomthing_Click()
MenuCommandClick "MNUSSomthing"
End Sub

Private Sub MNUStDwn_Click()
MenuCommandClick "MNUStDwn"
End Sub

Private Sub MNUSTime_Click()
MenuCommandClick "MNUSTime"
End Sub

Private Sub MNUSyncAClock_Click()
MenuCommandClick "MNUSyncAClock"
End Sub

Private Sub MNUTEN_Click()
MenuCommandClick "MNUTEN"
End Sub

Private Sub MNUTHREE_Click()
MenuCommandClick "MNUTHREE"
End Sub

Private Sub MNUTWO_Click()
MenuCommandClick "MNUTWO"
End Sub

Private Sub MNUVActions_Click()
MenuCommandClick "MNUVActions"
End Sub

Private Sub MNUWRuntime_Click()
MenuCommandClick "MNUWRuntime"
End Sub

Private Sub MNUyLO_Click()
MenuCommandClick "MNUyLO"
End Sub

Private Sub MNUyRS_Click()
MenuCommandClick "MNUyRS"
End Sub

Private Sub MNUySD_Click()
MenuCommandClick "MNUySD"
End Sub
