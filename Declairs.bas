Attribute VB_Name = "Declairs"
'   General Declares and Subs

Option Explicit
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Global PopupDLLFail As Boolean

Global CMDDDC As Boolean
Global DNCC As Integer
Global NoLockDown As Boolean

Global AllowSpashHide As Boolean

Global UseComMenu As Boolean
Global ComMenu As Integer

Global OTopW As Boolean
Global PWORD
Global UseLog As Integer
Global TimeInterval As String
Global ReminderTime As String
Global RemindTVal As String
Global ReminderMsg As String
Global MyName As String
Global pitchnum As Integer
Global speednum As Integer
Global GoCharacter As Boolean
Global LDUseIdle
Global CurrentDisk As DISKSPACEINFO
Global CTTempFile As Integer

Global ONEpath, ONEname, ONEcommand
Global TWOpath, TWOname, TWOcommand
Global THREEpath, THREEname, THREEcommand
Global FOURpath, FOURname, FOURcommand
Global FIVEpath, FIVEname, FIVEcommand
Global SIXpath, SIXname, SIXcommand
Global SEVENpath, SEVENname, SEVENcommand
Global EIGHTpath, EIGHTname, EIGHTcommand
Global NINEpath, NINEname, NINEcommand
Global TENpath, TENname, TENcommand

Global NoDiskSpaceCheck As Boolean

Public Const SizeToText = 2
Public Const BalloonOn = 1
Public Const AutoPace = 8

Public Const scUserAgent = "ComTalk iUpdate"

Const INTERNET_FLAG_RELOAD = &H80000000

Dim Temp
Dim starttime
Dim dirtemp As String

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
                                      
Public Sub ReTrans(RForm As Form) ' Makes top of form rounded
    ' Rounded edges code "Borrowed" from Olsen XP Components
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    X = RForm.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = RForm.Height / Screen.TwipsPerPixelY  'form in pixels
    
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn RForm.hWnd, Sum, True   'Sets corners transparent
End Sub

Public Function GetWindowsDir() As String ' Get Windows directory
    Dim Temp As String
    Dim ret As Long
    Const MAX_LENGTH = 145
    
    Temp = String$(MAX_LENGTH, 0)
    ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If
End Function

Public Function MyNameToRead() ' Get the name for ComTalk to read out
    MyNameToRead = GetSetting("ComTalk", "Options", "MyName", "User")
End Function

Public Function MyCharacter(Optional DONTUSEPATH As Boolean) ' Get selected character's filename
    MyCharacter = GetSetting("ComTalk", "Options", "MyCharacter", "CharacterFail")
    
    dirtemp = Not (Dir(Declairs.GetWindowsDir & "\MsAgent\Chars\" & MyCharacter) = "")
    
    If dirtemp = False Then MyCharacter = "CharacterFail"
    
    If MyCharacter = "CharacterFail" Then
        mainfrm.Character.Path = Declairs.GetWindowsDir & "MSAgent\Chars"
        If mainfrm.Character.ListCount > 0 Then
            mainfrm.Character.Selected(0) = True
            MyCharacter = mainfrm.Character.filename
            SaveSetting "ComTalk", "Options", "MyCharacter", mainfrm.Character.filename
        Else
            frmSplash.Hide
                MsgBox "No Characters found. Please install some and try again.", vbCritical + vbOKOnly + vbSystemModal, "ComTalk - Error"
            End
        End If
    End If
    
    If DONTUSEPATH = True Then MyCharacter = GetSetting("ComTalk", "Options", "MyCharacter", "CharacterFail")
End Function

Public Sub Wait(delay As Single)
    starttime = Timer
    Do Until Timer >= starttime + delay
    Loop
End Sub

Public Function OneLine(Txt)
    Static i, temptxt
    temptxt = ""
    
    For i = 1 To Len(Txt)
        If Mid(Txt, i, 1) = Chr(13) Or Mid(Txt, i, 1) = Chr(10) Then
            If Mid(Txt, i, 1) = Chr(10) Then
            Else
                temptxt = temptxt & " (New Line) "
            End If
        Else
            temptxt = temptxt & Mid(Txt, i, 1)
        End If
    Next
    OneLine = temptxt
End Function

Public Sub killprogram()
    On Error Resume Next
    mainfrm.Agent1.Characters.Unload "Genie"
    LoggingSubs.EndLog
    Dim i As Integer
    While Forms.Count > 1
        If Forms(i).Caption <> mainfrm.Caption Then
            Unload Forms(i)
        End If
        i = i + 1
    Wend
    
    Temp = GetSetting("ComTalk", "Program", "IsOpen", 0)
    If Temp = 1 Then
        SaveSetting "ComTalk", "Program", "IsOpen", 0
        EndLog
        SaveSetting "ComTalk", "Program", "IsOpen", 0
    End If
    
    End
End Sub

Public Sub GetCAFromReg() ' Get Custom Commands from registry
    ONEpath = GetSetting("ComTalk", "Vcommand1", "Path", "")
    ONEname = GetSetting("ComTalk", "Vcommand1", "Name", "")
    ONEcommand = GetSetting("ComTalk", "Vcommand1", "Command", "")
    TWOpath = GetSetting("ComTalk", "Vcommand2", "Path", "")
    TWOname = GetSetting("ComTalk", "Vcommand2", "Name", "")
    TWOcommand = GetSetting("ComTalk", "Vcommand2", "Command", "")
    THREEpath = GetSetting("ComTalk", "Vcommand3", "Path", "")
    THREEname = GetSetting("ComTalk", "Vcommand3", "Name", "")
    THREEcommand = GetSetting("ComTalk", "Vcommand3", "Command", "")
    FOURpath = GetSetting("ComTalk", "Vcommand4", "Path", "")
    FOURname = GetSetting("ComTalk", "Vcommand4", "Name", "")
    FOURcommand = GetSetting("ComTalk", "Vcommand4", "Command", "")
    FIVEpath = GetSetting("ComTalk", "Vcommand5", "Path", "")
    FIVEname = GetSetting("ComTalk", "Vcommand5", "Name", "")
    FIVEcommand = GetSetting("ComTalk", "Vcommand5", "Command", "")
    SIXpath = GetSetting("ComTalk", "Vcommand6", "Path", "")
    SIXname = GetSetting("ComTalk", "Vcommand6", "Name", "")
    SIXcommand = GetSetting("ComTalk", "Vcommand6", "Command", "")
    SEVENpath = GetSetting("ComTalk", "Vcommand7", "Path", "")
    SEVENname = GetSetting("ComTalk", "Vcommand7", "Name", "")
    SEVENcommand = GetSetting("ComTalk", "Vcommand7", "Command", "")
    EIGHTpath = GetSetting("ComTalk", "Vcommand8", "Path", "")
    EIGHTname = GetSetting("ComTalk", "Vcommand8", "Name", "")
    EIGHTcommand = GetSetting("ComTalk", "Vcommand8", "Command", "")
    NINEpath = GetSetting("ComTalk", "Vcommand9", "Path", "")
    NINEname = GetSetting("ComTalk", "Vcommand9", "Name", "")
    NINEcommand = GetSetting("ComTalk", "Vcommand9", "Command", "")
    TENpath = GetSetting("ComTalk", "Vcommand10", "Path", "")
    TENname = GetSetting("ComTalk", "Vcommand10", "Name", "")
    TENcommand = GetSetting("ComTalk", "Vcommand10", "Command", "")
End Sub


Public Sub SpeakError(ErrorText As String, Optional JUNK As VbMsgBoxStyle, Optional JUNK2 As String)
On Error GoTo dispMSG
    mainfrm.Agent1.Characters("Genie").StopAll
    SpeakText "An Error has occurred: " & ErrorText, NoOveride
Exit Sub
dispMSG:
MsgBox "An Error has occurred: " & ErrorText
End Sub

Public Function OpenURL(ByVal sUrl As String) As String
    '****************************************************
    'From http://www.freevbcode.com/ShowCode.Asp?ID=1252
    '*****************************************************
    
    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String
    
    hOpen = InternetOpen(scUserAgent, 1, _
        vbNullString, vbNullString, 0)
    
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
        INTERNET_FLAG_RELOAD, 0)
    
    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
            Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
            lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer
    
End Function

