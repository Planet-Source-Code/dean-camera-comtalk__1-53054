VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Check 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Checker"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1710
   ControlBox      =   0   'False
   Icon            =   "Check.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   1710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Height          =   255
      Left            =   1200
      Picture         =   "Check.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add to Bad List"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Height          =   300
      Left            =   840
      Picture         =   "Check.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Stop Animation"
      Top             =   1770
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   480
      Picture         =   "Check.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Play Animation"
      Top             =   1770
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   120
      Picture         =   "Check.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Delete from Bad List"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   960
      Picture         =   "Check.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Bad List"
      Top             =   1080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog common 
      Left            =   2400
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.acs"
      DialogTitle     =   "Open Character File"
      Filter          =   "Character Files | *.acs"
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "Check.frx":0C74
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open Character File"
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Actions For The Current Character"
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "NONE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "Type of character file"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin AgentObjectsCtl.Agent Agent 
      Left            =   3480
      Top             =   2280
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim req As AgentObjectsCtl.IAgentCtlRequest
Dim playreq As AgentObjectsCtl.IAgentCtlRequest
Dim Character As AgentObjectsCtl.IAgentCtlCharacter

Private Sub Command1_Click()
common.ShowOpen
If common.FileName <> "" Then
Agent.Characters.Unload "Character"
Set req = Agent.Characters.Load("Character", common.FileName)
If req < 1 Then
tmp = MsgBox("Error loading character. Do you want to add it to Bad List?", vbExclamation + vbYesNo, "Loading Error")
If tmp = 6 Then
SaveSetting "ComTalk", "BadList", List1.ListCount, common.FileTitle
List1.AddItem common.FileName, List1.ListCount + 1
SaveSetting "ComTalk", "BadList", "BadTotal", List1.ListCount
End If
End If
Agent.Characters("Character").Show
Set playreq = Agent.Characters("Character").Play("SendMail")
Agent.Characters("Character").Stop
If playreq < 1000 Then
Label2.Caption = "Office Assist."
delcommands
addofficecommands
tmp = MsgBox("Character is not ComTalk compatible. Do you want to add it to Bad List?", vbExclamation + vbYesNo, "Compatibility Error")
If tmp = 6 Then
List1.AddItem common.FileTitle
SaveSetting "ComTalk", "BadList", "BadTotal", List1.ListCount

End If
Else
Label2.Caption = "Normal"
delcommands
For Each Item In Agent.Characters("Character").AnimationNames
Combo1.AddItem Item
Next
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List1.ListCount > 0 Then
For i = 0 To List1.ListCount
If List1.Selected(i) = True Then
List1.RemoveItem i
SaveSetting "ComTalk", "BadList", "BadTotal", List1.ListCount
Exit For
End If
Next i
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
For i = 0 To List1.ListCount
If List1.List(i) = common.FileTitle Then
Exit Sub
End If
Next i
List1.AddItem common.FileTitle, List1.ListCount
SaveSetting "ComTalk", "BadList", List1.ListCount, common.FileTitle
SaveSetting "ComTalk", "BadList", "BadTotal", List1.ListCount
End Sub

Private Sub Command5_Click()
If Combo1.Text <> "" Then
Set playreq = Agent.Characters("Character").Play(Combo1.Text)
End If
End Sub

Private Sub Command6_Click()
Agent.Characters("Character").Stop
End Sub

Private Sub Form_Load()
Agent.RaiseRequestErrors = False
Agent.Characters.Load "Character"

On Error Resume Next
temp = GetSetting("ComTalk", "BadList", "BadTotal", 0)
For i = 1 To temp
tmp = GetSetting("ComTalk", "BadList", i, "")
List1.AddItem tmp
Next
End Sub

Sub delcommands()
On Error Resume Next
For i = 0 To Combo1.ListCount
Combo1.RemoveItem i
Next
End Sub

Sub addofficecommands()
Combo1.AddItem "Alert"
Combo1.AddItem "CheckingSomething"
Combo1.AddItem "Congratulate"
Combo1.AddItem "EmptyTrash"
Combo1.AddItem "Explain"
Combo1.AddItem "GestureUp"
Combo1.AddItem "GestureDown"
Combo1.AddItem "GestureLeft"
Combo1.AddItem "GestureRight"
Combo1.AddItem "GetArtsy"
Combo1.AddItem "GetAttention"
Combo1.AddItem "GetTechy"
Combo1.AddItem "GetWizardy"
Combo1.AddItem "Goodbye"
Combo1.AddItem "Greeting"
Combo1.AddItem "Hearing_1"
Combo1.AddItem "Hide"
Combo1.AddItem "Idle1_1"
Combo1.AddItem "Idle1_2"
Combo1.AddItem "Idle1_3"
Combo1.AddItem "Idle1_4"
Combo1.AddItem "Idle2_1"
Combo1.AddItem "Idle2_2"
Combo1.AddItem "Idle3_1"
Combo1.AddItem "Idle3_2"
Combo1.AddItem "LookDown"
Combo1.AddItem "LookDownLeft"
Combo1.AddItem "LookDownRight"
Combo1.AddItem "LookLeft"
Combo1.AddItem "LookRight"
Combo1.AddItem "LookUp"
Combo1.AddItem "LookUpRight"
Combo1.AddItem "Print"
Combo1.AddItem "Processing"
Combo1.AddItem "RestPose"
Combo1.AddItem "Save"
Combo1.AddItem "Searching"
Combo1.AddItem "SendMail"
Combo1.AddItem "Show"
Combo1.AddItem "Thinking"
Combo1.AddItem "Wave"
Combo1.AddItem "Writing"
End Sub
