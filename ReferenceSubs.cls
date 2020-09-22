VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTLink"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' LINK SUBROUTINES - USED BY COMTALK PLUGINS


' THIS SHOULD HAVE IT'S "INSTANCING" PROPERTY SET TO "GLOBAL MULTI USE"
' TO MAKE IT ACCESABLE TO COMTALK PLUGINS.


Public Enum LCharmelionButtonTypes
    [LSTYLE Windows 16-bit] = 1    'the old-fashioned Win16 button
    [LSTYLE Windows 32-bit] = 2    'the classic windows button
    [LSTYLE Windows XP] = 3        'the new brand XP button totally owner-drawn
    [LSTYLE MacOS 9] = 4           'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [LSTYLE Java metal] = 5        'there are also other styles but not so different from windows one
    [LSTYLE Netscape 6] = 6        'this is the button displayed in web-pages, it also appears in some java apps
    [LSTYLE Simple Flat] = 7       'the standard flat button seen on toolbars
    [LSTYLE Flat Highlight] = 8    'again the flat button but this one has no border until the mouse is over it
    [LSTYLE Office XP] = 9         'the new Office XP button
    [LSTYLE Transparent] = 11      'suggested from a user...
    [LSTYLE 3D Hover] = 12         'took this one from "Noteworthy Composer" toolbal
    [LSTYLE Oval Flat] = 13        'a simple Oval Button
    [LSTYLE KDE 2] = 14            'the great standard KDE2 button!
End Enum

Public Enum LLogMyText
    LYesLog = 0
    LNoLog = 1
End Enum

Public Enum LOverideSilentOption
    LNoOveride = 0
    LYesSilent = 1
    LYesNotSilent = 2
End Enum

' Links to ComTalk Subs
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
    Options.OKPress
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

Sub LEncryptPass(LPW As String)
    LockWorkstation.EncryptPass LPW
End Sub

Sub LLockStation()
    LockWorkstation.LockStation
End Sub

Sub LUnlockStation()
    LockWorkstation.UnlockStation
End Sub

Sub LAction(LActionName)
    LoggingSubs.Action LActionName
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

Sub LSpeakText(Ltts, Optional LOverideSilent As LOverideSilentOption, Optional LLogText As LLogMyText)
    LoggingSubs.SpeakText Ltts, LOverideSilent
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

Sub LGoPage(LPageNum As Integer)
    Options.GoPage LPageNum
End Sub

Sub LSpeakError(lError As String)
    SpeakError lError
End Sub

Function LIsFirstRun() As Boolean
    LIsFirstRun = FirstRun.IsFirstRun
End Function

Sub LFirstRunText()
    FirstRun.FirstRunText
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    PlugIns.KillPlugs
End Sub

Sub LGoButtonAppearance(LButtonAppearanceType As LCharmelionButtonTypes)
    CmdExt.GoButtonAppearance LButtonAppearanceType
End Sub

Sub LShowBox(LBoxText As String, LBoxTitle As String)
    onnow.ShowBox LBoxText, LBoxTitle
End Sub
