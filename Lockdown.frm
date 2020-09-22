VERSION 5.00
Begin VB.Form Lockdown 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Workststion Lock"
   ClientHeight    =   6885
   ClientLeft      =   5040
   ClientTop       =   4215
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin ComTalk.USEFULLctrl USEFULLctrl1 
      Left            =   960
      Top             =   1680
      _ExtentX        =   609
      _ExtentY        =   1032
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Workstation Locked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1400
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Lockdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    Me.Hide ' Hide lockdown until invoked
End Sub

Private Sub Text1_Change()
    If UCase(Text1.Text) = UCase(PWORD) Then 'Check for correct password
        USEFULLctrl1.ACTN_WIN_DisableCTRLALTDEL False
        USEFULLctrl1.ACTN_WIN_HideTaskBar False
        LockWorkstation.UnlockStation
        Text1.Text = ""
        Me.Hide
    End If
End Sub

Sub FixMyForm()
    'LOCKDOWN Functions (using Usefull Control)
If NoLockDown = True Then Exit Sub ' Don't load if command line parameter to disable it is used
    
    Text1.Text = ""
    Text1.Left = (Screen.Width / 2) - (Text1.Width / 2)
    Text1.Top = (Screen.Height / 2) - (Text1.Height / 2)
    Label1.Left = (Screen.Width / 2) - (Label1.Width / 2)
    Label1.Top = (Screen.Height / 2) - (Label1.Height / 2) - Text1.Height - 100
    USEFULLctrl1.ACTN_PRG_BringWindowToTop
    USEFULLctrl1.ACTN_PRG_CoverEntireScreen
    USEFULLctrl1.ACTN_PRG_StayOnTop True
    USEFULLctrl1.ACTN_WIN_DisableCTRLALTDEL True
    USEFULLctrl1.ACTN_WIN_MinimizeAllWindows
    USEFULLctrl1.ACTN_WIN_HideTaskBar True
    USEFULLctrl1.ACTN_WIN_ShowInTaskList False
    Shape1.Width = Me.Width
    If Label1.ForeColor = RGB(0, 0, 0) Then Label1.ForeColor = RGB(255, 255, 255)
End Sub
