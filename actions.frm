VERSION 5.00
Begin VB.Form actions 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk Actions"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   Icon            =   "actions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin ComTalk.chameleonButton Command2 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "actions.frx":0E42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ComTalk.chameleonButton Command3 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Stop"
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "actions.frx":0E5E
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Play"
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "actions.frx":0E7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "actions.frx":0E96
      Left            =   120
      List            =   "actions.frx":0E98
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub command2_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Action Combo1.Text
End Sub

Private Sub Command3_Click()
    mainfrm.Agent1.Characters("Genie").Stop
    Action "restpose"
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3  'Check to see if form stays on top
    
    ReTrans Me 'Make the top of the form rounded
End Sub

Sub PopulateActions()
On Error Resume Next
Combo1.Clear
    For Each Item In mainfrm.Agent1.Characters("Genie").AnimationNames 'Get the available animation names
        Combo1.AddItem Item
    Next
End Sub
