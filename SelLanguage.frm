VERSION 5.00
Begin VB.Form SelLanguage 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Agent Language"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   ControlBox      =   0   'False
   Icon            =   "SelLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin ComTalk.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Select"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SelLanguage.frx":628A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Spanish"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   6
      Tag             =   "&H0C0A"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Greek"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   5
      Tag             =   "&H0408"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "German"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   4
      Tag             =   "&H0407"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "French"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   3
      Tag             =   "&H040C"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dutch"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Tag             =   "&H0413"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Italian"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Tag             =   "&H0410"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton OptLanguage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "English (US)"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Tag             =   "&H0409"
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox UAIOL 
      Caption         =   "Use speech output and input engine of the same language"
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "LanID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Please select your language."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "SelLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer

Private Sub chameleonButton1_Click()
    SaveSetting "ComTalk", "Language", "UseAlternateIO", UAIOL.Value
    
    For i = 1 To OptLanguage.Count
        If OptLanguage(i).Value = True Then
            SaveSetting "ComTalk", "Language", "LanguageID", OptLanguage(i).Tag
            SaveSetting "ComTalk", "Language", "LanguageNum", i
            SaveSetting "ComTalk", "Language", "LanguageName", OptLanguage(i).Caption
            Exit For
        End If
    Next
    
    Me.Hide
    GlobalRefreshLang
    On Error GoTo 0
    mainfrm.LoadComTalk True
End Sub

Private Sub Form_Load()
Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'Make form stay on top
        
    On Error Resume Next 'REQUIRED - Stops other forms calling the Character when it is hidden from producing errors
    EnableSupported 'Enable supported languages' option buttons
    
    SaveSetting "ComTalk", "Language", "UseAlternateIO", 1
    
    OptLanguage_Click 1 'Select the English language by default
    
    Static i As Integer
    For i = 1 To OptLanguage.Count
    Debug.Print "DEFAULT LANGUAGE CHECK: " & Int(OptLanguage(i).Tag) & " = " & Int(GetUserDefaultLangID)
    If Int(OptLanguage(i).Tag) = Int(GetUserDefaultLangID) Then ' Found the computer's default language
    If OptLanguage(i).Enabled = True Then ' Make sure the language data is in the LANGUAGE.ini file before selecting
    OptLanguage(i).Value = True             '  \   Select the computer's
    OptLanguage(i).ForeColor = &HFF0000     '  |   default language and
    OptLanguage_Click i                     '  /   Colour it blue
    Exit For
    End If
    End If
    Next
     
    tempint = GetSetting("ComTalk", "Language", "LanguageNum", 1)
    OptLanguage(Int(tempint)).Value = True
    
    FixLangSLang 'Fix form's language
    ReTrans Me 'Make top of form rounded
End Sub

Sub FixLangSLang()
    If FoundLangFile = False Then Exit Sub
    
    Me.Caption = Language.GetLanguage(27)
    Me.Label1.Caption = Language.GetLanguage(26)
    Me.chameleonButton1.Caption = Language.GetLanguage(23)
End Sub

Sub EnableSupported() 'Enable supported languages' option buttons
    Static i As Integer

    If Language.FoundLangFile = False Then
        Me.Hide
    End If
    
    For i = 1 To OptLanguage.Count
        If Language.CheckForSupportedLanguage(OptLanguage(i).Caption) = True Then
            OptLanguage(i).Enabled = True
            OptLanguage(i).ForeColor = RGB(0, 0, 100)
        Else
            OptLanguage(i).Enabled = False
            OptLanguage(i).ForeColor = 0
        End If
    Next
    
    For i = 1 To OptLanguage.Count
    Debug.Print "DEFAULT LANGUAGE CHECK: " & Int(OptLanguage(i).Tag) & " = " & Int(GetUserDefaultLangID)
    If Int(OptLanguage(i).Tag) = Int(GetUserDefaultLangID) Then ' Found the computer's default language
    If OptLanguage(i).Enabled = True Then ' Make sure the language data is in the LANGUAGE.ini file before selecting
    OptLanguage(i).Value = True             '  \   Select the computer's
    OptLanguage(i).ForeColor = &HFF0000     '  |   default language and
    OptLanguage_Click i                     '  /   Colour it blue
    Exit For
    End If
    End If
    Next
     
    tempint = GetSetting("ComTalk", "Language", "LanguageNum", 1)
    OptLanguage(Int(tempint)).Value = True
End Sub

Private Sub OptLanguage_Click(Index As Integer)
Label3.Caption = OptLanguage(Index).Tag
End Sub
