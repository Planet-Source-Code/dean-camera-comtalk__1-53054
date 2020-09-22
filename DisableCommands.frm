VERSION 5.00
Begin VB.Form DisableCommands 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk Hide Commands"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "DisableCommands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin ComTalk.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "DisableCommands.frx":0E42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox EDCMDS 
      Height          =   1410
      ItemData        =   "DisableCommands.frx":0E5E
      Left            =   360
      List            =   "DisableCommands.frx":0E74
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Disabling a command will also disable the new menu, due to indexing problems."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Uncheck the commands below to hide them from the ComTalk menu. If a command is hidden, it will stll be voice-active."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "DisableCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
On Error Resume Next
For I = 0 To Me.EDCMDS.ListCount - 1
SaveSetting "ComTalk", "DisableCommands", Me.EDCMDS.List(I), Me.EDCMDS.Selected(I)
Next

EnableDisableCommands
Unload CustomMenu ' Force menu refresh
Load CustomMenu   '-------/

Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3 'Check to see if form should be on top

On Error Resume Next
For I = 0 To Me.EDCMDS.ListCount - 1
Me.EDCMDS.Selected(I) = GetSetting("ComTalk", "DisableCommands", Me.EDCMDS.List(I), True)
Next
End Sub

Sub EnableDisableCommands()
For I = 0 To Me.EDCMDS.ListCount - 1
Temp = GetSetting("ComTalk", "DisableCommands", Me.EDCMDS.List(I), True)

If Temp = False Then PopupDLLFail = True

Select Case I
    Case 0
        CustomMenu.MNUActns.Visible = Temp
    Case 1
        CustomMenu.MNUCD.Visible = Temp
    Case 2
        CustomMenu.MNURClip.Visible = Temp
    Case 3
        CustomMenu.MNUSSomthing.Visible = Temp
    Case 4
        CustomMenu.MNUCptr.Visible = Temp
    Case 5
        CustomMenu.MNUWRuntime.Visible = Temp
End Select
Next
End Sub
