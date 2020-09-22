VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Customreminders 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComTalk Custom Reminders"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "Creminders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame SetCusDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custom Date"
      Height          =   975
      Left            =   3960
      TabIndex        =   16
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox CMonth 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox CDay 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   "31"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "  Day          Month"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reminder Text"
      Height          =   1215
      Left            =   3960
      TabIndex        =   14
      Top             =   600
      Width           =   3135
      Begin VB.TextBox Text1 
         ForeColor       =   &H00C00000&
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time (24 Hour)"
      Height          =   855
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   2415
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   856
         TabIndex        =   10
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "Text2"
         BuddyDispid     =   196613
         OrigLeft        =   2160
         OrigTop         =   360
         OrigRight       =   2355
         OrigBottom      =   615
         Max             =   24
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "1"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "1"
         Top             =   315
         Width           =   390
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1696
         TabIndex        =   9
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text3"
         BuddyDispid     =   196614
         OrigLeft        =   3240
         OrigTop         =   360
         OrigRight       =   3480
         OrigBottom      =   615
         Max             =   59
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1150
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Day"
      Height          =   1335
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
      Begin VB.OptionButton RDay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Every Day"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton RDay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weekend"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton RDay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Weekday"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton RDay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custom Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
   End
   Begin ComTalk.chameleonButton command2 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3000
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Creminders.frx":0E42
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
      Left            =   2400
      TabIndex        =   1
      Top             =   3000
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Creminders.frx":0E5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00C00000&
      Height          =   2010
      ItemData        =   "Creminders.frx":0E7A
      Left            =   120
      List            =   "Creminders.frx":0E9C
      TabIndex        =   0
      Top             =   840
      Width           =   1095
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
      Width           =   7575
   End
End
Attribute VB_Name = "Customreminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Mins1_Click()
    
End Sub

Private Sub Command1_Click()
    'If user puts a '0' in front of the time, strip it
    On Error Resume Next
    If Int(Mid(Text2.Text, 1, 1)) = 0 Then Text2.Text = Mid(Text2.Text, 2)
    
    If Val(Text2.Text) > 0 And Val(Text2.Text) < 25 Then
    Else
        If Text2.Text <> "" Then
            SpeakError "Invalid hour!", vbCritical
        End If
        Exit Sub
    End If
    
    If Val(Text3.Text) > (0 - 1) And Val(Text3.Text) < 61 Then
    Else
        If Text3.Text <> "" Then
            SpeakError "Invalid minute!", vbCritical
        End If
        Exit Sub
    End If
    
    If RDay(4).Value = True Then
        If IsDate(CDay.Text & "/" & CMonth.Text) = False Then
            SpeakError "Invalid custom date!"
        Else
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "CusDays", CDay.Text
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "CusMonth", CMonth.Text
        End If
    End If
    
    For i = 1 To RDay.Count ' Find the index number of the selected reminder day
        dayselected = i
        If RDay(i).Value = True Then Exit For ' If the day selected is found, stop the loop and save the index number to a variable
    Next
    
    For i = 0 To List1.ListCount
        If List1.Selected(i) = True Then 'Save the selected reminder's data to the registry
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "Reminder", Text1.Text
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "Hours", Text2.Text
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "Mins", Text3.Text
            SaveSetting "ComTalk", "Reminder" & Str(i + 1), "Date", dayselected
        End If
    Next
    
    TimeSubs.GetCRFromReg
End Sub

Private Sub command2_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 'Check to see if form stays on top
    List1.Selected(0) = True
    
    ReTrans Me 'Make the top of the form rounded
End Sub

Private Sub List1_Click()
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then 'Load the selected reminder's data from the registry
            Text1.Text = GetSetting("ComTalk", "Reminder" & Str(i + 1), "Reminder", "")
            Text2.Text = GetSetting("ComTalk", "Reminder" & Str(i + 1), "Hours", "12")
            Text3.Text = GetSetting("ComTalk", "Reminder" & Str(i + 1), "Mins", "0")
            RDay(GetSetting("ComTalk", "Reminder" & Str(i + 1), "Date", 1)).Value = True
        End If
    Next
    
    SetCusDate.Enabled = RDay(4).Value
    
    If RDay(4).Enabled = True Then
        CDay.Text = GetSetting("ComTalk", "Reminder" & Str(i + 1), "CusDays", "31")
        CMonth.Text = GetSetting("ComTalk", "Reminder" & Str(i + 1), "CusMonth", "1")
    End If
End Sub

Private Sub RDay_Click(Index As Integer)
    SetCusDate.Enabled = RDay(4).Value
End Sub

