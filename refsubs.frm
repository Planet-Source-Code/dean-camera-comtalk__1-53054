VERSION 5.00
Begin VB.Form refsubs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reference Subs"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"refsubs.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "refsubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Links to all ComTalk Subs

' Add a 'L' before sub name

Sub LCheckDiskSpace()
mainfrm.checkdiskspace
End Sub

Sub LEndMe()
mainfrm.endme
End Sub

Sub LPopulateCommands()
mainfrm.populatecommands
End Sub

Sub LPopulateVCommands()
mainfrm.populateVcommands
End Sub

Sub LReload()
mainfrm.reload
End Sub

Sub LStartupMCheck()
mainfrm.startupmcheck
End Sub

Sub LOKPress()
options.OKPress
End Sub

Sub LGetCAFromReg()
Declairs.GetCAFromReg
End Sub

Function LGetWindowsDir()
LGetWindowsDir = Declairs.GetWindowsDir
End Function

Sub LKillProgram()
Declairs.killprogram
End Sub

Function LMyCharacter(Optional LDontUsePath As Boolean)
LMyCharacter = Declairs.MyCharacter(LDontUsePath)
End Function

Function LMyNameToRead()
LMyNameToRead = Declairs.MyNameToRead
End Function

Function LOneLine(Ltxt As String)
LOneLine = Declairs.OneLine(Ltxt)
End Function

Sub LWait(Ldelay As Single)
Declairs.Wait Ldelay
End Sub

Function LGetDiskInfo(LRootPathName As String)
LGetDiskInfo = DiskSubs.GetDiskInfo(LRootPathName)
End Function

Sub LDecodePassword()
LockWorkstation.DecodePassword
End Sub

Sub LLockStation()
LockWorkstation.LockStation
End Sub

Sub LUnlockStation()
LockWorkstation.UnlockStation
End Sub

Sub LAction(LActionName)
LoggingSubs.Action ActionName
End Sub

Sub LAddLogText(LText)
LoggingSubs.AddLogText LText
End Sub

Sub LBeginLog()
LoggingSubs.BeginLog
End Sub

Sub LChangingCharacter()
LoggingSubs.ChangingCharacter
End Sub

Sub LEndLog()
LoggingSubs.EndLog
End Sub

Sub LReq(LRequestName As String)
LoggingSubs.Req LRequestName
End Sub

Sub LSpeakText(Ltts)
LoggingSubs.SpeakText Ltts
End Sub

Sub LWarnUnexpected()
LoggingSubs.WarnUnexpected
End Sub

Sub LExecuteScript()
ScriptSubs.ExecuteScript
End Sub

Sub LInjectCode(LExCode As String)
ScriptSubs.InjectCode LExCode
End Sub

Sub LKillScript()
ScriptSubs.KillScript
End Sub

Sub LSetObjects()
ScriptSubs.SetObjects
End Sub

Sub LCheckForCustomReminders()
TimeSubs.checkforcustomreminders
End Sub

Sub LGetCRFromReg()
TimeSubs.GetCRFromReg
End Sub

Sub LRemind()
TimeSubs.Remind
End Sub

Sub LRemindCompare(Lnum As Integer, Lreminder, Lmins, Lhours)
TimeSubs.REMINDCOMPARE Lnum, Lreminder, Lmins, Lhours
End Sub

Sub LTCompare(LTimeType As String, LCurrentTime)
TimeSubs.TCOMPARE LTimeType, LCurrentTime
End Sub

Sub LIdlingChar()
LoggingSubs.IdlingChar
End Sub

Sub LExtendedCommands()
CmdExt.ExtendCommands
End Sub

Sub LExtendedMenu(LGetCA As Boolean)
CmdExt.ExtendedMenu LGetCA
End Sub

Sub LGoPage(LPageNum As Integer)
options.GoPage LPageNum
End Sub

Sub LSpeakError(LError As String)
SpeakError LError
End Sub
