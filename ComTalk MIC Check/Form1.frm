VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mic Checker"
   ClientHeight    =   600
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":030A
      Top             =   1200
      Width           =   2055
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR dsr 
      Height          =   735
      Left            =   1320
      OleObjectBlob   =   "Form1.frx":044F
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label pbar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   960
      Shape           =   3  'Circle
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu E 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dsr_VUMeter(ByVal beginhi As Long, ByVal beginlo As Long, ByVal level As Long)
On Error Resume Next
pgbar level
If level > 30000 Then
Shape1.FillColor = &HFF&
Else
Shape1.FillColor = &HFF00&
End If
End Sub

Private Sub E_Click()
End
End Sub

Private Sub Form_Load()
dsr.GrammarFromString Text1.Text
dsr.Activate
End Sub

Private Sub Form_Unload(Cancel As Integer)
dsr.Deactivate
End Sub

Sub pgbar(VAL)
tmp = VAL / 1000
pbar.Caption = ""
For i = 1 To tmp
pbar.Caption = pbar.Caption & "|"
Next
End Sub

Private Sub Label1_Click()

End Sub
