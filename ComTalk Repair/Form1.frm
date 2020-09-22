VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F2A762&
   BorderStyle     =   0  'None
   Caption         =   "ComTalk Repair"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ForeColor       =   &H00F2A762&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":57E2
   ScaleHeight     =   4245
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComTalkRepair.AquaButton Command2 
      Height          =   450
      Left            =   1920
      TabIndex        =   11
      Top             =   3750
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComTalkRepair.AquaButton Command6 
      Height          =   450
      Left            =   3480
      TabIndex        =   10
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   794
      Caption         =   "Wipe Plugin Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComTalkRepair.AquaButton Command3 
      Height          =   450
      Left            =   3480
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   794
      Caption         =   "Delete Password"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComTalkRepair.AquaButton Command1 
      Height          =   450
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
      Caption         =   "Select New Character"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComTalkRepair.AquaButton Command4 
      Height          =   450
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
      Caption         =   "Fix ""IsOpen"" Value"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComTalkRepair.AquaButton Command5 
      Height          =   450
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   794
      Caption         =   "Wipe Registry Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   480
      Pattern         =   "*.acs"
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   75
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "Form1.frx":D7AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   90
      Picture         =   "Form1.frx":D800
      Top             =   45
      Width           =   1950
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   0
      Picture         =   "Form1.frx":10322
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   6645
   End
   Begin VB.Image Right 
      Height          =   4380
      Left            =   6555
      Picture         =   "Form1.frx":10374
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   60
   End
   Begin VB.Image Left 
      Height          =   4260
      Left            =   0
      Picture         =   "Form1.frx":103C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   6600
      Picture         =   "Form1.frx":10410
      Top             =   4200
      Width           =   60
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   0
      Top             =   4200
      Width           =   60
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use this to delete just the plugins settings, if a plugin is malfunctioning."
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use this to clear the current settings, if the current settings are causing ComTalk to malfunction."
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use this if ComTalk crashes, and plugins wont allow you to access their menu."
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use this if you have forgotten your LockDown password."
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use this if the current character is not functioning correctly."
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   1320
      Top             =   45
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Command1_Click()
SaveSetting "ComTalk", "Options", "MyCharacter", File1.FileName
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
DeleteSetting "ComTalk", "Options", "Lock Station PWord"
End Sub

Private Sub Command4_Click()
SaveSetting "ComTalk", "Program", "IsOpen", 0
End Sub

Private Sub Command5_Click()
On Error Resume Next
DeleteSetting "ComTalk"
End Sub

Private Sub Command6_Click()
On Error Resume Next
DeleteSetting "ComTalk", "Plugins"
End Sub

Private Sub Form_Load()
On Error Resume Next
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
ReTrans
File1.Path = GetWindowsDir & "MSAgent\Chars\"
File1.Selected(0) = True
End Sub

Function GetWindowsDir() As String
        GetWindowsDir = Environ("windir") & "\"
End Function

Sub ReTrans()
' Rounded edges code "Borrowed" from Olsen XP Components
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    X = Me.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = Me.Height / Screen.TwipsPerPixelY  'form in pixels
    
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn Me.hWnd, Sum, True   'Sets corners transparent
End Sub

