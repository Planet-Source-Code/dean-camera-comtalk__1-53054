VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1695
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   480
      Top             =   1440
   End
   Begin VB.Label BetaVer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BETA - DO NOT DISTRIBUTE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visit www.en-tech.i8.com for more free programs!"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   5280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   4890
      X2              =   4890
      Y1              =   240
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   4920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Dean Camera, 2001 - 2003"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4905
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   120
      Picture         =   "frmSplash.frx":320F
      Top             =   480
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   1560
      Picture         =   "frmSplash.frx":4D1D
      Top             =   430
      Width           =   2055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
    
    If ProgRelease = "BETA" Then 'Show BETA message (if applicable)
        Me.BetaVer.Visible = True
    End If
    
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3  'Make form stay on top
    Label4.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Temp = GetSetting("ComTalk", "Options", "UseSplash", 1)
    
    ReTrans Me 'Make top of form rounded
    
    BTEMP = GetSetting("ComTalk", "Options", "ButtonType", 2) 'Fix ComTalk's forms buttons to the correct style
    If BTEMP = 0 Then CmdExt.GoButtonAppearance [STYLE Windows 16-bit]
    If BTEMP = 1 Then CmdExt.GoButtonAppearance [STYLE Windows 32-bit]
    If BTEMP = 2 Then CmdExt.GoButtonAppearance [STYLE Windows XP]
    If BTEMP = 3 Then CmdExt.GoButtonAppearance [STYLE MacOS 9]
    If BTEMP = 4 Then CmdExt.GoButtonAppearance [STYLE Java metal]
    If BTEMP = 5 Then CmdExt.GoButtonAppearance [STYLE Netscape 6]
    If BTEMP = 6 Then CmdExt.GoButtonAppearance [STYLE Simple Flat]
    If BTEMP = 7 Then CmdExt.GoButtonAppearance [STYLE Flat Highlight]
    If BTEMP = 8 Then CmdExt.GoButtonAppearance [STYLE Office XP]
    If BTEMP = 9 Then CmdExt.GoButtonAppearance [STYLE 3D Hover]
    If BTEMP = 10 Then CmdExt.GoButtonAppearance [STYLE Oval Flat]
    If BTEMP = 11 Then CmdExt.GoButtonAppearance [STYLE KDE 2]
    
    If Temp = 1 Then 'If "Use Splash Screen" option enabled, show the splash screen
        Me.Show
        Timer1.Enabled = True
    Else 'or just load the program
        Timer1_Timer
    End If

    DoEvents
End Sub


Private Sub Title_Click()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "FORM SPLASH SCREEN UNLOADED TO FREE UP MEMORY."
End Sub

Private Sub Timer1_Timer()
    Me.Hide
    Load mainfrm
    Unload Me ' Save memory by unloading the splash screen from memory
End Sub
