VERSION 5.00
Begin VB.Form options 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ComTalk Options"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Commands"
      Height          =   615
      Index           =   17
      Left            =   1200
      TabIndex        =   75
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
      Begin ComTalk.chameleonButton chameleonButton2 
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Enable/Disable Commands"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   12582912
         BCOLO           =   12582912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "options.frx":0E42
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   795
      TabIndex        =   66
      Top             =   600
      Width           =   855
      Begin VB.Image OptionCatagoryButton 
         Height          =   855
         Index           =   6
         Left            =   120
         Top             =   4320
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   5
         Left            =   120
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   4
         Left            =   120
         Top             =   2880
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   3
         Left            =   120
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   2
         Left            =   120
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   1
         Left            =   120
         Top             =   720
         Width           =   615
      End
      Begin VB.Image OptionCatagoryButton 
         Height          =   735
         Index           =   0
         Left            =   120
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   200
         Picture         =   "options.frx":0E5E
         Top             =   4320
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   200
         Picture         =   "options.frx":1728
         Top             =   3600
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   200
         Picture         =   "options.frx":1FF2
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   195
         Picture         =   "options.frx":28BC
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   200
         Picture         =   "options.frx":3186
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   200
         Picture         =   "options.frx":3A50
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   195
         Picture         =   "options.frx":3D5A
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lock Workstation"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   73
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reminders"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   72
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Appearance"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   71
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "ComTalk"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   70
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Character"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Startup"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Speech"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   615
      Index           =   10
      Left            =   1200
      TabIndex        =   31
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox ufe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Fast Exit"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox spic 
      Height          =   375
      Left            =   5160
      Picture         =   "options.frx":4064
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   65
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComTalk.chameleonButton command2 
      Height          =   375
      Left            =   2640
      TabIndex        =   46
      Top             =   5640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "options.frx":4352
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
      Left            =   1680
      TabIndex        =   45
      Top             =   5640
      Width           =   735
      _ExtentX        =   1296
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
      COLTYPE         =   3
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "options.frx":436E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custom Reminders"
      Height          =   1215
      Index           =   15
      Left            =   1200
      TabIndex        =   54
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox ARR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Also Read Reminder"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox Upop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use popup window to display reminders"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   2415
      End
      Begin VB.Line Line8 
         X1              =   360
         X2              =   240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   240
         Y1              =   600
         Y2              =   960
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lock Workstation"
      Height          =   1455
      Index           =   16
      Left            =   1200
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "New:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Old:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speech Type"
      Height          =   1215
      Index           =   1
      Left            =   1200
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ListBox speechtype 
         ForeColor       =   &H00C00000&
         Height          =   645
         ItemData        =   "options.frx":438A
         Left            =   120
         List            =   "options.frx":4397
         TabIndex        =   30
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Character Options"
      Height          =   2775
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox idleani 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Idle Animations"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox AP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Words as they are Said"
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox STF 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Size Speech Bubble to Fit Text"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox sndeffect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Character Sound Effects"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox ssb 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Speech Bubble"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maintenance"
      Height          =   735
      Index           =   11
      Left            =   1200
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox dsc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check Drive C:\ for Low Disk Space (< 50mb)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speed and  Pitch"
      Height          =   975
      Index           =   3
      Left            =   1200
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin ComTalk.DMSlider CusSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "options.frx":43B6
         MaxValue        =   "200"
         Mask            =   "false"
      End
      Begin ComTalk.chameleonButton command7 
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Defaults"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "options.frx":4474
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ComTalk.chameleonButton command6 
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Test"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "options.frx":4490
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ComTalk.DMSlider CusPitch 
         Height          =   255
         Left            =   1440
         TabIndex        =   62
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "options.frx":44AC
         MaxValue        =   "200"
         Mask            =   "false"
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ComTalk Windows"
      Height          =   735
      Index           =   9
      Left            =   1200
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox OTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stay On Top"
         Height          =   375
         Left            =   480
         TabIndex        =   42
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Silent Speech"
      Height          =   975
      Index           =   2
      Left            =   1200
      TabIndex        =   43
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox usesilent 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use a Thought Bubble Instead of Reading Out Text"
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Button Appearance"
      Height          =   2655
      Index           =   13
      Left            =   1200
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ListBox List1 
         ForeColor       =   &H00C00000&
         Height          =   1620
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2415
      End
      Begin ComTalk.chameleonButton chameleonButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   2160
         Width           =   2415
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Sample Button"
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
         MICON           =   "options.frx":456A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Preview:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reminders"
      Height          =   1695
      Index           =   14
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton Command5 
         Caption         =   "Custom Reminders"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton te 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Every Hour"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton th 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Half Hour"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1080
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton tq 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quarter Hour"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox remindt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remind me of the time every:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   240
         X2              =   840
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   720
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   240
         Y1              =   600
         Y2              =   1560
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Character Size"
      Height          =   1935
      Index           =   5
      Left            =   1200
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin ComTalk.DMSlider cHeight 
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "options.frx":4586
         MaxValue        =   "200"
         Mask            =   "false"
      End
      Begin ComTalk.chameleonButton command8 
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Defaults"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "options.frx":4644
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ComTalk.DMSlider cWidth 
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "options.frx":4660
         MaxValue        =   "200"
         Mask            =   "false"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logging"
      Height          =   735
      Index           =   7
      Left            =   1200
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox logme 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Log all events"
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internet Functions"
      Height          =   1455
      Index           =   12
      Left            =   1200
      TabIndex        =   57
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox DNCFIC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Do Not Check for Internet Connection"
         Height          =   615
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox SIM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show ""Internet"" menu"
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "When the program starts:"
      Height          =   2775
      Index           =   4
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox splash 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Splash Screen"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox usrname 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Tag             =   "Z"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox sn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Speak Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Don't Speak Time or Date"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Speak time and date"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Speak date"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Speak time"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Line Line2 
         X1              =   480
         X2              =   240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   2040
         Y2              =   2160
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speech Volume"
      Height          =   855
      Index           =   0
      Left            =   1200
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin ComTalk.DMSlider SpeechVol 
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "options.frx":471E
         Mask            =   "False"
      End
   End
   Begin VB.Frame OptionsPage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extended Functions"
      Height          =   1695
      Index           =   8
      Left            =   1200
      TabIndex        =   38
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox extended 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Extended Voice Commands"
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"options.frx":47DC
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   1320
      X2              =   3840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Some options may require ComTalk to be restarted to take effect."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1320
      TabIndex        =   77
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a catagory from the menu at the left of the screen."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1200
      TabIndex        =   74
      Top             =   1440
      Width           =   2775
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
      Width           =   4095
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sval
Dim startupMval As String
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub Check1_Click()
    
End Sub


Private Sub chameleonButton2_Click()
DisableCommands.Show
End Sub

Private Sub cHeight_Change(NewValue As Double)
    mainfrm.Agent1.Characters("Genie").Height = cHeight.Value
End Sub

Private Sub cHeight_Click()
    mainfrm.Agent1.Characters("Genie").Height = cHeight.Value
End Sub

Private Sub Command1_Click()

'SAVE OPTIONS TO REGISTRY

    SaveSetting "ComTalk", "Options", "DoNotCheckINet", DNCFIC.Value
    SaveSetting "ComTalk", "Options", "ShowInternetMenu", SIM.Value
    
    SaveSetting "ComTalk", "Options", "RemindPopup", Upop.Value
    SaveSetting "ComTalk", "Options", "ReadPopup", ARR.Value
    
    If SIM.Value = 1 Then CustomMenu.MNUInet.Visible = True Else CustomMenu.MNUInet.Visible = False
    
    SilentTemp = usesilent.Value
    If SilentTemp = 0 Then SilentTemp = False
    If SilentTemp = 1 Then SilentTemp = True
    
    If options.OTop.Value = 1 Then
        OTopW = True
    Else
        OTopW = False
    End If
    
    mainfrm.Agent1.Characters("Genie").AutoPopupMenu = False
    
    UseLog = logme.Value
    
    If Text1.Text = PWORD Then
        password = ""
        EncryptPass Text2.Text
        PWORD = Text2.Text
        Text1.Text = ""
        Text2.Text = ""
    End If
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            SaveSetting "ComTalk", "Options", "ButtonType", i
            If i = 0 Then CmdExt.GoButtonAppearance [STYLE Windows 16-bit]
            If i = 1 Then CmdExt.GoButtonAppearance [STYLE Windows 32-bit]
            If i = 2 Then CmdExt.GoButtonAppearance [STYLE Windows XP]
            If i = 3 Then CmdExt.GoButtonAppearance [STYLE MacOS 9]
            If i = 4 Then CmdExt.GoButtonAppearance [STYLE Java metal]
            If i = 5 Then CmdExt.GoButtonAppearance [STYLE Netscape 6]
            If i = 6 Then CmdExt.GoButtonAppearance [STYLE Simple Flat]
            If i = 7 Then CmdExt.GoButtonAppearance [STYLE Flat Highlight]
            If i = 8 Then CmdExt.GoButtonAppearance [STYLE Office XP]
            If i = 9 Then CmdExt.GoButtonAppearance [STYLE 3D Hover]
            If i = 10 Then CmdExt.GoButtonAppearance [STYLE Oval Flat]
            If i = 11 Then CmdExt.GoButtonAppearance [STYLE KDE 2]
        End If
    Next
    
    SaveSetting "ComTalk", "Options", "Silent", SilentTemp
    SaveSetting "ComTalk", "Options", "OnTopWindows", OTop.Value
    SaveSetting "ComTalk", "Options", "ExtendedVoice", extended.Value
    SaveSetting "ComTalk", "Options", "UseFastExit", ufe.Value
    SaveSetting "ComTalk", "Options", "SpeechVol", SpeechVol.MaxValue - SpeechVol.Value
    SaveSetting "ComTalk", "Options", "LogEvents", logme.Value
    SaveSetting "ComTalk", "Options", "UseIdleAnimations", idleani.Value
    SaveSetting "ComTalk", "Options", "CustomSpeed", CusSpeed.Value
    SaveSetting "ComTalk", "Options", "CustomPitch", CusPitch.Value
    SaveSetting "ComTalk", "Options", "CustomCharHeight", cHeight.Value
    SaveSetting "ComTalk", "Options", "CustomCharWidth", cWidth.Value
    SaveSetting "ComTalk", "Options", "CheckDiskSpace", dsc.Value
    SaveSetting "ComTalk", "Options", "UseSplash", splash.Value
    
    RemindTVal = GetSetting("ComTalk", "Options", "RemindTimeVal", "Quarter")
    
    pitchnum = GetSetting("ComTalk", "Options", "CustomPitch", mainfrm.Agent1.Characters("Genie").Pitch)
    speednum = GetSetting("ComTalk", "Options", "CustomSpeed", mainfrm.Agent1.Characters("Genie").Speed)
    
    If idleani.Value = 1 Then mainfrm.Agent1.Characters("Genie").IdleOn = True
    If idleani.Value = 0 Then mainfrm.Agent1.Characters("Genie").IdleOn = False
    
    If speechtype.Selected(0) = True Then
        SaveSetting "ComTalk", "Options", "SpeechType", 0
    ElseIf speechtype.Selected(1) = True Then
        SaveSetting "ComTalk", "Options", "SpeechType", 1
    Else
        SaveSetting "ComTalk", "Options", "SpeechType", 2
    End If
    
    If usrname.Text = "" Then
        usrname.Text = "User"
    Else
        SaveSetting "ComTalk", "Options", "MyName", usrname.Text
    End If
    
    If Option4.Value = True Then
        startupMval = "N"
    ElseIf Option1.Value = True Then
        startupMval = "T"
    ElseIf Option2.Value = True Then
        startupMval = "D"
    Else
        startupMval = "TD"
    End If
    
    If sndeffect.Value = 1 Then
        mainfrm.Agent1.Characters("Genie").SoundEffectsOn = True
    Else
        mainfrm.Agent1.Characters("Genie").SoundEffectsOn = False
    End If
    
    If STF.Value = 0 Then
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not SizeToText)
    Else
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or SizeToText
    End If
    
    If ssb.Value = 0 Then
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not BalloonOn)
    Else
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or BalloonOn
    End If
    
    If AP.Value = 0 Then
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style And (Not AutoPace)
    Else
        mainfrm.Agent1.Characters("Genie").Balloon.Style = mainfrm.Agent1.Characters("Genie").Balloon.Style Or AutoPace
    End If
    
    SaveSetting "ComTalk", "Options", "ShowSpeechBubble", ssb.Value
    SaveSetting "ComTalk", "Options", "UseCharacterSNDs", sndeffect.Value
    SaveSetting "ComTalk", "Options", "SayNameOnStartup", sn.Value
    SaveSetting "ComTalk", "Options", "StartupMessage", startupMval
    SaveSetting "ComTalk", "Options", "SizeBubbleToText", STF.Value
    SaveSetting "ComTalk", "Options", "AutoWordPace", AP.Value
    SaveSetting "ComTalk", "Options", "RemindTime", remindt.Value
    
    If tq.Value = True Then
        RemindTVal = "Quarter"
    ElseIf th.Value = True Then
        RemindTVal = "Half"
    Else
        RemindTVal = "Hour"
    End If
    
    SaveSetting "ComTalk", "Options", "RemindTimeVal", RemindTVal
    
    'Set reminding time parameter
    doremind = GetSetting("ComTalk", "Options", "RemindTime")
    Unload Me
End Sub

Private Sub Command10_Click()
    Customreminders.Show
End Sub


Private Sub command2_Click()
    Me.Hide
    Form_Load
    Unload Me
End Sub

Private Sub Command5_Click()
    Customreminders.Show
End Sub

Private Sub Command6_Click()
    SpeakText "\pit=" & CusPitch.Value & "\ " & "\spd=" & CusSpeed.Value & "\" & "This is a test"
End Sub

Private Sub Command7_Click()
    CusSpeed.Value = mainfrm.Agent1.Characters("Genie").Speed
    CusPitch.Value = mainfrm.Agent1.Characters("Genie").Pitch
End Sub

Private Sub Command8_Click()
    cHeight.Value = mainfrm.Agent1.Characters("Genie").OriginalHeight
    cWidth.Value = mainfrm.Agent1.Characters("Genie").OriginalWidth
End Sub


Sub OKPress()
    Command1_Click
End Sub

Private Sub FNT_Change()
    
End Sub

Private Sub cWidth_Change(NewValue As Double)
    mainfrm.Agent1.Characters("Genie").Width = cWidth.Value
End Sub

Private Sub cWidth_Click()
    mainfrm.Agent1.Characters("Genie").Width = cWidth.Value
End Sub

Private Sub Form_Load()
Label14.Visible = True
Label15.Visible = True
Line1.Visible = True

Me.Image1.Picture = mainfrm.Image1.Picture ' Show the ComTalk banner at the top of the page
    
    On Error Resume Next
        
    SpeechVol.BackColor = RGB(255, 255, 255)      ' The new sliders are a bit dodgy, so
    SpeechVol.Caption = ""                        ' force them to retain their settings
    SpeechVol.MaxValue = 65535                    '     |
    cHeight.BackColor = RGB(255, 255, 255)        '     |
    cWidth.Caption = ""                           '     |
    cHeight.MaxValue = 200                        '     |
    cWidth.MaxValue = 200                         '     |
    CusSpeed.BackColor = RGB(255, 255, 255)       '     |
    CusPitch.Caption = ""                         '     |
    CusSpeed.MaxValue = 200                       '     |
    CusPitch.MaxValue = 200                       '     |
    Set CusPitch.HPicture = spic.Picture          '     |
    Set CusSpeed.HPicture = spic.Picture          '     |
    Set SpeechVol.HPicture = spic.Picture         '     |
    Set cWidth.HPicture = spic.Picture            '     |
    Set cHeight.HPicture = spic.Picture           ' ____|
    
'GET OPTIONS FROM REGISTRY
    
    ReTrans Me
            
' Check and add options
    
    DNCFIC.Value = GetSetting("ComTalk", "Options", "DoNotCheckINet", 0)
    
    SIM.Value = GetSetting("ComTalk", "Options", "ShowInternetMenu", 1)
    
    For i = 1 To List1.ListCount
        List1.RemoveItem i
    Next
    
    List1.AddItem "Windows 16 Bit", 0
    List1.AddItem "Windows 32 Bit", 1
    List1.AddItem "Windows XP", 2
    List1.AddItem "Mac", 3
    List1.AddItem "Java Metal", 4
    List1.AddItem "Netscape 6", 5
    List1.AddItem "Simple Flat", 6
    List1.AddItem "Flat Highlight", 7
    List1.AddItem "Office XP", 8
    List1.AddItem "3D Hover", 9
    List1.AddItem "Oval Flat", 10
    List1.AddItem "KDE 2", 11
    
    BTYPE = GetSetting("ComTalk", "Options", "ButtonType", 2)
    List1.Selected(BTYPE) = True
    If BTYPE = 0 Then CmdExt.GoButtonAppearance [STYLE Windows 16-bit]
    If BTYPE = 1 Then CmdExt.GoButtonAppearance [STYLE Windows 32-bit]
    If BTYPE = 2 Then CmdExt.GoButtonAppearance [STYLE Windows XP]
    If BTYPE = 3 Then CmdExt.GoButtonAppearance [STYLE MacOS 9]
    If BTYPE = 4 Then CmdExt.GoButtonAppearance [STYLE Java metal]
    If BTYPE = 5 Then CmdExt.GoButtonAppearance [STYLE Netscape 6]
    If BTYPE = 6 Then CmdExt.GoButtonAppearance [STYLE Simple Flat]
    If BTYPE = 7 Then CmdExt.GoButtonAppearance [STYLE Flat Highlight]
    If BTYPE = 8 Then CmdExt.GoButtonAppearance [STYLE Office XP]
    If BTYPE = 9 Then CmdExt.GoButtonAppearance [STYLE 3D Hover]
    If BTYPE = 10 Then CmdExt.GoButtonAppearance [STYLE Oval Flat]
    If BTYPE = 11 Then CmdExt.GoButtonAppearance [STYLE KDE 2]
    
    SilentTemp = GetSetting("ComTalk", "Options", "Silent", False)
    If SilentTemp = False Then SilentTemp = 0
    If SilentTemp = True Then SilentTemp = 1
    
    usesilent.Value = SilentTemp
    OTop.Value = GetSetting("ComTalk", "Options", "OnTopWindows", 0)
    extended.Value = GetSetting("ComTalk", "Options", "ExtendedVoice", 0)
    Text1.Text = ""
    Text2.Text = ""
    Text1_Change
    ufe.Value = GetSetting("ComTalk", "Options", "UseFastExit", 0)
    SpeechVol.Value = SpeechVol.MaxValue - GetSetting("ComTalk", "Options", "SpeechVol", 65535)
    sptype = GetSetting("ComTalk", "Options", "SpeechType", 0)
    speechtype.Selected(sptype) = True
    logme.Value = GetSetting("ComTalk", "Options", "LogEvents", 0)
    cHeight.Value = GetSetting("ComTalk", "Options", "CustomCharHeight", mainfrm.Agent1.Characters("Genie").OriginalHeight)
    cWidth.Value = GetSetting("ComTalk", "Options", "CustomCharWidth", mainfrm.Agent1.Characters("Genie").OriginalWidth)
    idleani.Value = GetSetting("ComTalk", "Options", "UseIdleAnimations", 1)
    CusSpeed.Value = GetSetting("ComTalk", "Options", "CustomSpeed", mainfrm.Agent1.Characters("Genie").Speed)
    CusPitch.Value = GetSetting("ComTalk", "Options", "CustomPitch", mainfrm.Agent1.Characters("Genie").Pitch)
    dsc.Value = GetSetting("ComTalk", "Options", "CheckDiskSpace", 1)
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    splash.Value = GetSetting("ComTalk", "Options", "UseSplash", 1)
    Upop.Value = GetSetting("ComTalk", "Options", "RemindPopup", 0)
    ARR.Value = GetSetting("ComTalk", "Options", "ReadPopup", 0)
    remindt.Value = GetSetting("ComTalk", "Options", "RemindTime", 0)
    ssb.Value = GetSetting("ComTalk", "Options", "ShowSpeechBubble", 1)
    STF.Value = GetSetting("ComTalk", "Options", "SizeBubbleToText", 1)
    sndeffect.Value = GetSetting("ComTalk", "Options", "UseCharacterSNDs", 1)
    sn.Value = GetSetting("ComTalk", "Options", "SayNameOnStartup", 1)
    AP.Value = GetSetting("ComTalk", "Options", "AutoWordPace", 1)
    usrname.Text = Declairs.MyNameToRead
    startupMval = GetSetting("ComTalk", "Options", "StartupMessage", "N")
    If startupMval = "N" Then
        Option4.Value = True
    ElseIf startupMval = "T" Then
        Option1.Value = True
    ElseIf startupMval = "D" Then
        Option2.Value = True
    ElseIf startupMval = "TD" Then
        Option3.Value = True
    End If
    
    RemindTVal = GetSetting("ComTalk", "Options", "RemindTimeVal", "Quarter")
    
    If RemindTVal = "Quarter" Then
        tq.Value = True
    ElseIf RemindTVal = "Half" Then
        th.Value = True
    Else
        te.Value = True
    End If
    
    If Upop.Value = 1 Then
        ARR.Enabled = True
    Else
        ARR.Enabled = False
    End If
    
    List1.Selected(BTYPE) = True
    speechtype.Selected(GetSetting("ComTalk", "Options", "SpeechType", 0)) = True
    
    If OTopW = True Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
End Sub

Private Sub Imagelist1_Change()
    
End Sub

Private Sub HScroll1_Change()
    GoPage HScroll1.Value
End Sub

Private Sub List1_Click()
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then BTYPE = i
    Next
    
    If BTYPE = 0 Then Me.chameleonButton1.ButtonType = [Windows 16-bit]
    If BTYPE = 1 Then Me.chameleonButton1.ButtonType = [Windows 32-bit]
    If BTYPE = 2 Then Me.chameleonButton1.ButtonType = [Windows XP]
    If BTYPE = 3 Then Me.chameleonButton1.ButtonType = Mac
    If BTYPE = 4 Then Me.chameleonButton1.ButtonType = [Java metal]
    If BTYPE = 5 Then Me.chameleonButton1.ButtonType = [Netscape 6]
    If BTYPE = 6 Then Me.chameleonButton1.ButtonType = [Simple Flat]
    If BTYPE = 7 Then Me.chameleonButton1.ButtonType = [Flat Highlight]
    If BTYPE = 8 Then Me.chameleonButton1.ButtonType = [Office XP]
    If BTYPE = 9 Then Me.chameleonButton1.ButtonType = [3D Hover]
    If BTYPE = 10 Then Me.chameleonButton1.ButtonType = [Oval Flat]
    If BTYPE = 11 Then Me.chameleonButton1.ButtonType = [KDE 2]
End Sub

Private Sub OptionCatagoryButton_Click(Index As Integer)
Label14.Visible = False
Label15.Visible = False
Line1.Visible = False
GoPage Index + 1 ' Show the options corresponding to the clicked option catagory button
End Sub

Private Sub remindt_Click()
    If remindt.Value = 0 Then
        tq.Enabled = False
        th.Enabled = False
        te.Enabled = False
    Else
        tq.Enabled = True
        th.Enabled = True
        te.Enabled = True
    End If
End Sub

Sub GoPage(PageNum As Integer)
Label14.Visible = False

' Hide all options pages:
For i = 0 To 17
OptionsPage(i).Visible = False
Next

' Show the correct options pages:

Select Case PageNum
Case 1                              ' SPEECH
OptionsPage(0).Visible = True
OptionsPage(1).Visible = True
OptionsPage(2).Visible = True
OptionsPage(3).Visible = True
Case 2                              ' STARTUP
OptionsPage(4).Visible = True
Case 3                              ' CHARACTER
OptionsPage(5).Visible = True
OptionsPage(6).Visible = True
Case 4                              ' COMTALK
OptionsPage(7).Visible = True
OptionsPage(8).Visible = True
OptionsPage(9).Visible = True
OptionsPage(10).Visible = True
OptionsPage(11).Visible = True
Case 5                              ' APPEARANCE
OptionsPage(12).Visible = True
OptionsPage(13).Visible = True
OptionsPage(17).Visible = True
Case 6                              ' REMINDERS
OptionsPage(14).Visible = True
OptionsPage(15).Visible = True
Case 7                              ' LOCK WORKSTATION
OptionsPage(16).Visible = True
End Select
End Sub

Private Sub Text1_Change()
    If Text1.Text = PWORD Then
        Text2.Enabled = True
        Text2.BackColor = RGB(255, 255, 255)
    Else
        Text2.Enabled = False
        Text2.BackColor = &HE0E0E0
    End If
End Sub

Private Sub Form_Terminate()
    command2_Click
End Sub

Private Sub Upop_Click()
    If Upop.Value = 1 Then
        ARR.Enabled = True
    Else
        ARR.Enabled = False
    End If
End Sub
