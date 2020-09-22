VERSION 5.00
Begin VB.Form TextFixIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TextFix"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "Icon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Don't Use In ComTalk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use In Comtalk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "By Dean Camera, 1/12/02"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No Requirments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2520
      Picture         =   "Icon.frx":628A
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Icon.frx":BE9C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This Plugin automatically fixes what the character is saying from abbreviations or contractions to proper english."
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label NameOfPlug 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TextFix"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "TextFixIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Agent1_ActivateInput(ByVal CharacterID As String)

End Sub

Private Sub Option1_Click()
plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
For i = 1 To plugcount
If GetSetting("ComTalk", "Plugins", "Plugin " & i, "") = "TextFix.TextFixMain" Then
Exit Sub
End If
Next
SaveSetting "ComTalk", "Plugins", "Count", plugcount + 1
SaveSetting "ComTalk", "Plugins", "Plugin " & plugcount + 1, "TextFix.TextFixMain"
End Sub

Private Sub Option2_Click()
plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
Debug.Print "COUNT: " & plugcount
For i = 1 To plugcount
If GetSetting("ComTalk", "Plugins", "Plugin " & i, "") = "TextFix.TextFixMain" Then
myold = i
Debug.Print "OLD: " & myold
maxnum = plugcount
SaveSetting "ComTalk", "Plugins", "Count", plugcount - 1
DeleteSetting "ComTalk", "Plugins", "Plugin " & myold
Debug.Print myold & "|" & maxnum

For z = myold + 1 To maxnum
Debug.Print "Z: " & z
Temp = GetSetting("ComTalk", "Plugins", "Plugin " & z, "")
Debug.Print "TEMP: " & Temp
SaveSetting "ComTalk", "Plugins", "Plugin " & (z - 1), Temp
Next z

plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
If plugcount > 1 Then
DeleteSetting "ComTalk", "Plugins", "Plugin " & plugcount
End If
End If
Next i
End Sub
