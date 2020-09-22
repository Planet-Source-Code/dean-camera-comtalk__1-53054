VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PluginMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name Of Plugin"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "PluginMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1800
      Top             =   1440
   End
   Begin RichTextLib.RichTextBox Log 
      Height          =   1335
      Left            =   4680
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"PluginMain.frx":5D52
   End
   Begin VB.Timer cancelsaid 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1080
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FakeSend"
      Height          =   300
      Left            =   5040
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6255
      Left            =   4800
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer updatesaid 
      Interval        =   3000
      Left            =   1560
      Top             =   1920
   End
   Begin VB.TextBox said 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "<"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "0 Unread E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2355
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Testing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Log:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Output:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Status:"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Users:"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "PluginMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
