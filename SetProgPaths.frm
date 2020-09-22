VERSION 5.00
Begin VB.Form SetProgPaths 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Inbuilt Program Paths"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "SetProgPaths.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      ForeColor       =   &H00C00000&
      Height          =   1035
      ItemData        =   "SetProgPaths.frx":23D2
      Left            =   120
      List            =   "SetProgPaths.frx":23E5
      TabIndex        =   3
      Top             =   525
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   765
      Width           =   2655
   End
   Begin ComTalk.chameleonButton command2 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1125
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
      MICON           =   "SetProgPaths.frx":2420
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
      TabIndex        =   1
      Top             =   1125
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
      MICON           =   "SetProgPaths.frx":243C
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program Path:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   525
      Width           =   2055
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
Attribute VB_Name = "SetProgPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Command1_Click()
    If List1.Selected(0) = True Then
        SaveSetting "ComTalk", "InBuiltPath", "Calculator", Text1.Text
    ElseIf List1.Selected(1) = True Then
        SaveSetting "ComTalk", "InBuiltPath", "NotePad", Text1.Text
    ElseIf List1.Selected(2) = True Then
        SaveSetting "ComTalk", "InBuiltPath", "Defrag", Text1.Text
    ElseIf List1.Selected(3) = True Then
        SaveSetting "ComTalk", "InBuiltPath", "Explorer", Text1.Text
    ElseIf List1.Selected(4) = True Then
        SaveSetting "ComTalk", "InBuiltPath", "SoundRecorder", Text1.Text
    End If
End Sub

Private Sub command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 'Check to see if form should be on top
    List1.Selected(0) = True 'Select first item by default

    ReTrans Me 'Make top of form rounded
End Sub

Private Sub List1_Click() 'Get inbuilt program's selected path
    If List1.Selected(0) = True Then
        Text1.Text = GetSetting("ComTalk", "InBuiltPath", "Calculator", "C:\Windows\calc.exe")
    ElseIf List1.Selected(1) = True Then
        Text1.Text = GetSetting("ComTalk", "InBuiltPath", "NotePad", "C:\Windows\Notepad.exe")
    ElseIf List1.Selected(2) = True Then
        Text1.Text = GetSetting("ComTalk", "InBuiltPath", "Defrag", "C:\Windows\Defrag.exe")
    ElseIf List1.Selected(3) = True Then
        Text1.Text = GetSetting("ComTalk", "InBuiltPath", "Explorer", "C:\Windows\Explorer.exe")
    ElseIf List1.Selected(4) = True Then
        Text1.Text = GetSetting("ComTalk", "InBuiltPath", "SoundRecorder", "C:\Windows\Sndrec32.exe")
    End If
End Sub

