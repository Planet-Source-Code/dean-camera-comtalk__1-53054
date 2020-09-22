VERSION 5.00
Begin VB.Form onnow 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "Reminder Form"
   ClientHeight    =   2340
   ClientLeft      =   6150
   ClientTop       =   5085
   ClientWidth     =   3150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   Begin ComTalk.chameleonButton CommandButton 
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Custom"
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
      COLTYPE         =   4
      FOCUSR          =   0   'False
      BCOL            =   16053492
      BCOLO           =   16053492
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "onnow.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox DeanPic 
      BackColor       =   &H00F4F4F4&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      Picture         =   "onnow.frx":001C
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin ComTalk.chameleonButton chameleonButton1 
      Height          =   135
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "onnow.frx":1D96
      PICN            =   "onnow.frx":1DB2
      PICH            =   "onnow.frx":1FE4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox RText 
      BackColor       =   &H00F4F4F4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "onnow.frx":2216
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image BO 
      Height          =   195
      Left            =   1440
      Picture         =   "onnow.frx":2230
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image BU 
      Height          =   180
      Left            =   1200
      MouseIcon       =   "onnow.frx":2452
      MousePointer    =   99  'Custom
      Picture         =   "onnow.frx":275C
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image BD 
      Height          =   180
      Left            =   1680
      MouseIcon       =   "onnow.frx":297E
      MousePointer    =   99  'Custom
      Picture         =   "onnow.frx":2C88
      Top             =   2040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Right 
      Height          =   1260
      Left            =   3090
      Picture         =   "onnow.frx":307C
      Stretch         =   -1  'True
      Top             =   195
      Width           =   60
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   20
      Width           =   2970
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   0
      Picture         =   "onnow.frx":312A
      Top             =   1440
      Width           =   60
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   3120
      Picture         =   "onnow.frx":319C
      Top             =   1440
      Width           =   60
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   0
      Picture         =   "onnow.frx":320E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3165
   End
   Begin VB.Image Left 
      Height          =   1260
      Left            =   0
      Picture         =   "onnow.frx":3280
      Stretch         =   -1  'True
      Top             =   195
      Width           =   60
   End
   Begin VB.Image TitleRight 
      Height          =   210
      Left            =   3000
      Picture         =   "onnow.frx":332E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   150
   End
   Begin VB.Image TitleLeft 
      Height          =   210
      Left            =   0
      Picture         =   "onnow.frx":3910
      Stretch         =   -1  'True
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Title 
      Height          =   210
      Left            =   120
      Picture         =   "onnow.frx":3EF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2879
   End
End
Attribute VB_Name = "onnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Sub chameleonButton1_Click()
HideSB
DeanPic.Top = 1000
End Sub

Private Sub CommandButton_Click()
mainfrm.SideBoxButtonClick CommandButton.Caption
End Sub

Private Sub Form_Load()
    Me.Height = 1515
    ReTrans Me 'Make top of form rounded
    Me.Left = Screen.Width
    Me.Top = Screen.Height - (Me.Height * 4)
End Sub

Private Sub Image2_Click()
Me.Hide
Me.Left = Screen.Width - Me.Width
End Sub

Sub ShowSB()
Do
    Me.Left = Me.Left - 1 'Show form
    DoEvents
Loop While Me.Left > Screen.Width - Me.Width
End Sub

Sub HideSB()
Do
    Me.Left = Me.Left + 1 'Hide form
    DoEvents
Loop While Me.Left <> Screen.Width
End Sub

Sub ShowBox(BoxText As String, BoxTitle As String, Optional ButtonCaption As String) 'Show reminder box from the right of the screen
    Debug.Print "Sidebox shown - Title: " & BoxTitle & " Message: " & BoxText
        
    If ButtonCaption <> "" Then
    CommandButton.Caption = ButtonCaption
    CommandButton.Visible = True
    CommandButton.Width = Len(ButtonCaption) * 100 + 20
    CommandButton.Left = (Me.Width / 2) - (CommandButton.Width / 2)
    Else
    CommandButton.Visible = False
    End If
    
    Me.Show
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'Make form stay on top
    RText.Text = BoxText
    Label2.Caption = BoxTitle
    ShowSB
End Sub

Private Sub Image1_Click()

End Sub
