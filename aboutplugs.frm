VERSION 5.00
Begin VB.Form aboutplugs 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk About Plugins"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "aboutplugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
      Begin VB.Label HCTM 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label PReq 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label MBy 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label CName 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label PBS 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden From ComTalk Plugin Menu:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Requirements:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Made By:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pass Spoken Text to Plugin:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   1035
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin ComTalk.chameleonButton CloseMe 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   4320
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
      MICON           =   "aboutplugs.frx":0E42
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
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Currently Active Plugins (used in ComTalk)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   4335
   End
End
Attribute VB_Name = "aboutplugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub CloseMe_Click()
    Me.Hide
    List1.Clear
    Unload Me
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 'Check to see if form stays on top
    LoadPlugInList 'Load the list of available plugins
    
    ReTrans Me 'Make top of form rounded
End Sub

Private Sub Label7_Click()
    
End Sub

Private Sub List1_Click()
    On Error Resume Next
    For i = 1 To PlugInsLST.Count  'Get selected plugin's stats
        PBS.Caption = PlugInsLST(i).PassBeforeSay
        CName.Caption = PlugInsLST(i).PClassName
        MBy.Caption = PlugInsLST(i).PMadeBy
        PReq.Caption = PlugInsLST(i).PRequirements
        HCTM.Caption = Not PlugInsLST(i).ShowInMenu
    Next
End Sub

Public Function LoadPlugInList()
    List1.AddItem "TEMP"
    List1.Clear
    
    For i = 1 To PlugInsLST.Count
        List1.AddItem PlugInsLST.Item(i).FriendlyName 'Get the name of the plugin
    Next
End Function

