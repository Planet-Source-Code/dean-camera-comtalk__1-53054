VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form vcommands 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk Voice Commands"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "command.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   3960
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComTalk.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   14
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "command.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ComTalk.chameleonButton command1 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "command.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   2010
      ItemData        =   "command.frx":0902
      Left            =   120
      List            =   "command.frx":0924
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin ComTalk.chameleonButton command2 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "command.frx":0997
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Voice Command:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Path:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "vcommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chameleonButton1_Click()
    CD.DefaultExt = "*.exe"
    CD.Filter = "Executables (*.exe)|*.exe"
    CD.ShowOpen
    Text2.Text = CD.filename
End Sub

Private Sub Command1_Click()
    If List1.Selected(0) = True Then
        SaveSetting "ComTalk", "Vcommand1", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand1", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand1", "Command", Text3.Text
    ElseIf List1.Selected(1) = True Then
        SaveSetting "ComTalk", "Vcommand2", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand2", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand2", "Command", Text3.Text
    ElseIf List1.Selected(2) = True Then
        SaveSetting "ComTalk", "Vcommand3", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand3", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand3", "Command", Text3.Text
    ElseIf List1.Selected(3) = True Then
        SaveSetting "ComTalk", "Vcommand4", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand4", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand4", "Command", Text3.Text
    ElseIf List1.Selected(4) = True Then
        SaveSetting "ComTalk", "Vcommand5", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand5", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand5", "Command", Text3.Text
    ElseIf List1.Selected(5) = True Then
        SaveSetting "ComTalk", "Vcommand6", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand6", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand6", "Command", Text3.Text
    ElseIf List1.Selected(6) = True Then
        SaveSetting "ComTalk", "Vcommand7", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand7", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand7", "Command", Text3.Text
    ElseIf List1.Selected(7) = True Then
        SaveSetting "ComTalk", "Vcommand8", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand8", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand8", "Command", Text3.Text
    ElseIf List1.Selected(8) = True Then
        SaveSetting "ComTalk", "Vcommand9", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand9", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand9", "Command", Text3.Text
    ElseIf List1.Selected(9) = True Then
        SaveSetting "ComTalk", "Vcommand10", "Name", Text1.Text
        SaveSetting "ComTalk", "Vcommand10", "Path", Text2.Text
        SaveSetting "ComTalk", "Vcommand10", "Command", Text3.Text
    End If
End Sub

Private Sub command2_Click()
    mainfrm.Agent1.Characters("Genie").Commands.RemoveAll
    mainfrm.populatecommands
    Unload Me
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 ' See if form should be on top
    List1.Selected(1) = True ' Select first item by default
    ReTrans Me ' Make top of form rounded
End Sub

Private Sub List1_Click() ' Get custom command's data from the registry
    If List1.Selected(0) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand1", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand1", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand1", "Command", "")
    ElseIf List1.Selected(1) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand2", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand2", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand2", "Command", "")
    ElseIf List1.Selected(2) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand3", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand3", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand3", "Command", "")
    ElseIf List1.Selected(3) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand4", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand4", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand4", "Command", "")
    ElseIf List1.Selected(4) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand5", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand5", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand5", "Command", "")
    ElseIf List1.Selected(5) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand6", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand6", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand6", "Command", "")
    ElseIf List1.Selected(6) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand7", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand7", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand7", "Command", "")
    ElseIf List1.Selected(7) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand8", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand8", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand8", "Command", "")
    ElseIf List1.Selected(8) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand9", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand9", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand9", "Command", "")
    ElseIf List1.Selected(9) = True Then
        Text1.Text = GetSetting("ComTalk", "Vcommand10", "Name", "")
        Text2.Text = GetSetting("ComTalk", "Vcommand10", "Path", "")
        Text3.Text = GetSetting("ComTalk", "Vcommand10", "Command", "")
    End If
End Sub
