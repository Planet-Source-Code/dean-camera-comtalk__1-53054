VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Mail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CS Email Monitor"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Now"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   4440
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   2535
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   4575
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0 KB"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Total Size"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total Emails"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Last Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   4575
      Begin VB.CheckBox Check3 
         Caption         =   "Enable Sound"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Launch At Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "5"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check For Email Every"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3960
         Picture         =   "Mail.frx":0442
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3360
         Picture         =   "Mail.frx":074C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "min"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Login Password"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Login Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "POP3 Mail Sever"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "http://www.croftssoftware.com"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   210
      TabIndex        =   22
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MenuCheck 
         Caption         =   "Check Now"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205


Private Const HKEY_DYN_DATA = &H80000006

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTON3STATE = &H10

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTTOPMOST = -2

Private Const ILD_TRANSPARENT = &H1

Private Type NOTIFYICONDATA
    cbSize As Long
    mhWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim TheForm As NOTIFYICONDATA

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number


Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_TOP
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State         As POP3States
Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private FilePathName As String
Dim X As Long
Dim Y As Long
Dim wavSetup As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Sub SaveSettings()
Dim fFile As Integer
fFile = FreeFile
'save settings
Open App.Path & "\Settings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "server=" & Text1.Text
Print #fFile, "loginname=" & Text2.Text
Print #fFile, "loginpass=" & Text3.Text
Print #fFile, "checkeveryon=" & Check1.Value
Print #fFile, "checkeverymin=" & Text5.Text
Print #fFile, "launchatstartup=" & Check2.Value
Print #fFile, "enablesound=" & Check3.Value
Close fFile
DoEvents
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Call AddToRun("CS Email Monitor Light", App.Path & "\" & App.EXEName & ".exe")
End If
If Check2.Value = 0 Then
Call RemoveFromRun("CS Email Monitor Light")
End If
End Sub

Private Sub Command1_Click()
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
DoEvents
Call Me.CheckEmail
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim AppDir As String
AppDir = App.Path

Me.Caption = "CS Email Monitor v." & App.Major & "." & App.Minor & "." & App.Revision & " - Light"
X = 0
Y = 0

DoEvents

FilePathName = AppDir + "\Settings.inf"
server = GetPrivateProfileString("settings", "server", "", FilePathName)
loginname = GetPrivateProfileString("settings", "loginname", "", FilePathName)
loginpass = GetPrivateProfileString("settings", "loginpass", "", FilePathName)
checkeveryon = GetPrivateProfileString("settings", "checkeveryon", "", FilePathName)
checkeverymin = GetPrivateProfileString("settings", "checkeverymin", "", FilePathName)
launchatstartup = GetPrivateProfileString("settings", "launchatstartup", "", FilePathName)
enablesound = GetPrivateProfileString("settings", "enablesound", "", FilePathName)

DoEvents

Text1.Text = server
Text2.Text = loginname
Text3.Text = loginpass
Check1.Value = checkeveryon
Text5.Text = checkeverymin
Check2.Value = launchatstartup
Check3.Value = enablesound
DoEvents

SysTray


If Text1.Text = "" Then
Exit Sub
End If

If Text2.Text = "" Then
Exit Sub
End If

If Text3.Text = "" Then
Exit Sub
End If

Me.Hide

End Sub

Public Sub CheckEmail()
'On Error Resume Next
Dim YY As String
Dim ZZ As String


If Text1.Text = "" Then
Text4.Text = "Please Enter A Mail Server!"
Exit Sub
End If

If Text2.Text = "" Then
Text4.Text = "Please Enter A User Name!"
Exit Sub
End If

If Text3.Text = "" Then
Text4.Text = "Please Enter A Password!"
Exit Sub
End If


Text4.Text = ""

DoEvents
    'Check the emptiness of all the text fields except for the txtBody
    '
    'Change the value of current session state
    m_State = POP3_Connect
    '
    'Close the socket in case it was opened while another session
    Winsock1.Close
    '
    'reset the value of the local port in order to let to the
    'Windows Sockets select the new one itself
    'It's necessary in order to prevent the "Address in use" error,
    'which can appear if the Winsock Control has already used while the 
    'previous session
    Winsock1.LocalPort = 0
    '
    'POP3 server waits for the connection request at the port 110.
    'According with that we want the Winsock Control to be connected to
    'the port number 110 of the server we have supplied in combo1 field
    Winsock1.Connect Text1.Text, 110
End Sub
Public Sub DisconnectMe()
On Error Resume Next
'm_State = POP3_QUIT
m_colMessages.Clear
Winsock1.SendData "QUIT" & vbCrLf
Text4.Text = Text4.Text & "QUIT" & vbCrLf
Text4.SelStart = Len(Text4.Text)
DoEvents
DoEvents
End Sub
Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub


Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Cancel = True
End If
Me.Hide
Me.SaveSettings
End Sub

Private Sub Label8_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)
End Sub
Private Sub MenuCheck_Click()
Command1_Click
End Sub

Private Sub MenuExit_Click()
Call FrmMain.SaveSettings
Shell_NotifyIcon NIM_DELETE, TheForm
DoEvents
End
End Sub

Private Sub Timer1_Timer()

If Check1.Value = 1 Then

X = X + 1

If X >= 60 Then
Y = Y + 1
X = 0
End If

If Y >= Val(Text5.Text) Then
Command1_Click
Y = 0
End If

Else

X = 0
Y = 0

End If

End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    Static TotalSize            As Long
    Static TotalSize2            As Long
    Dim EmailNum                As Long
    Dim wavSetup As String
    '
    'Save the received data into strData variable
If Text7.Text = 0 Then
SysTray
TheForm.szTip = Text7.Text & " Messages" & " as of " & Text6.Text
Shell_NotifyIcon NIM_MODIFY, TheForm
Else
SysTray2

End If
    Winsock1.GetData strData
    Text4.Text = Text4.Text & strData & vbCrLf
    Text4.SelStart = Len(Text4.Text)
    
    If Left$(strData, 1) = "+" Or m_State = POP3_TOP Then
        'If the first character of the server's response is "+" then
        'server accepted the client's command and waits for the next one
        'If this symbol is "-" then here we can do nothing
        'and execution skips to the Else section of the code
        'The first symbol may differ from "+" or "-" if the received
        'data are the part of the message's body, i.e. when
        'm_State = POP3_TOP (the loading of the message state)
        Select Case m_State
            Case POP3_Connect
                '
                'Reset the number of messages
                intMessages = 0
                intCurrentMessage = 0
                '
                'Change current state of session
                m_State = POP3_USER
                '
                'Send to the server the USER command with the parameter.
                'The parameter is the name of the mail box
                'Don't forget to add vbCrLf at the end of the each command!
                Winsock1.SendData "USER " & Text2.Text & vbCrLf
                Text4.Text = Text4.Text & "USER " & Text2.Text & vbCrLf
                Text4.SelStart = Len(Text4.Text)
                'Here is the end of Winsock1_DataArrival routine until the
                'next appearing of the DataArrival event. But next time this
                'section will be skipped and execution will start right after
                'the Case POP3_USER section.
            Case POP3_USER
                '
                'This part of the code runs in case of successful response to
                'the USER command.
                'Now we have to send to the server the user's password
                '
                'Change the state of the session
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & Text3 & vbCrLf
                Text4.Text = Text4.Text & "PASS ***** " & vbCrLf
                Text4.SelStart = Len(Text4.Text)
            Case POP3_PASS
                '
                'The server answered positively to the process of the
                'identification and now we can send the STAT command. As a
                'response the server is going to return the number of
                'messages in the mail box and its size in octets
                '
                ' Change the state of the session
                m_State = POP3_STAT
                '
                'Send STAT command to know how many
                'messages in the mailbox
                Winsock1.SendData "STAT" & vbCrLf
                Text4.Text = Text4.Text & "STAT" & vbCrLf
                Text4.SelStart = Len(Text4.Text)
            Case POP3_STAT
                '
                'The server's response to the STAT command looks like this:
                '"+OK 0 0" (no messages at the mailbox) or "+OK 3 7564"
                '(there are messages). Evidently, the first of all we have to
                'find out the first numeric value that contains in the
                'server's response
                intMessages = Get_After_Seperator(strData, 1, " ")
                TotalSize = Get_After_Seperator(strData, 2, " ")
                Text6.Text = Date & " at " & Time
                Text7.Text = intMessages
                Text8.Text = Format(TotalSize / 1024, 0) & " KB"
                DoEvents
                If intMessages = 0 Then
                Text4.Text = Text4.Text & "There are no messages on the server." & vbCrLf
                Winsock1.SendData "QUIT" & vbCrLf
                Text4.Text = Text4.Text & "QUIT" & vbCrLf
                Text4.SelStart = Len(Text4.Text)
                Exit Sub
                End If
                Winsock1.SendData "QUIT" & vbCrLf
                
            Case POP3_QUIT
                'No matter what data we've received it's important
                'to close the connection with the mail server
                Winsock1.Close
                Call DisconnectMe

                'Now we're calling the ListMessages routine in order to
                'fill out the ListView control with the messages we've          
                'downloaded
                'Call ListMessages
        End Select
    Else
        'As you see, there is no sophisticated error
        'handling. We just close the socket and show the server's response
        'That's all. By the way even fully featured mail applications
        'do the same.
            Winsock1.Close
            Text4.Text = "POP3 Error: " & strData & vbExclamation & "POP3 Error"
    End If
End Sub
Function Get_After_Seperator(ByVal strString As String, ByVal intNthOccurance As Integer, ByVal strSeperator As String) As String
    'On Error Resume Next

    
    'check for intNthOccurance = 0--i.e. fir
    '     st one


    If (intNthOccurance = 0) Then


        If (InStr(strString, strSeperator) > 0) Then
                Get_After_Seperator = Left(strString, InStr(strString, strSeperator) - 1)
        Else
                Get_After_Seperator = strString
        End If
    Else
        'not the first one
        'init start of string on first comma
        intStartOfString = InStr(strString, strSeperator)
        
        'place start of string after intNthOccur
        '     ance-th comma (-1 since
        'already did one
        boolNotFound = 0


        For intIndex = 1 To intNthOccurance - 1
            'get next comma
            intStartOfString = InStr(intStartOfString + 1, strString, strSeperator)
            'check for not found


            If (intStartOfString = 0) Then
                boolNotFound = 1
            End If
        Next intIndex
        
        'put start of string past 1st comma
        intStartOfString = intStartOfString + 1
        
        'check for ending in a comma


        If (intStartOfString > Len(strString)) Then
            boolNotFound = 1
        End If
        


        If (boolNotFound = 1) Then
            Get_After_Seperator = "NOT FOUND"
        Else
            intEndOfString = InStr(intStartOfString, strString, strSeperator)
            
            ' check for no second comma (i.e. end of
            '     string)


            If (intEndOfString = 0) Then
                intEndOfString = Len(strString) + 1
            Else
                intEndOfString = intEndOfString - 1
            End If
            Get_After_Seperator = Mid$(strString, intStartOfString, intEndOfString - intStartOfString + 1)
        End If
    End If
End Function
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim ErrorStats As String
Winsock1.Close
    ErrorStats = Number & " : " & Description
    Text4.Text = Text4.Text & ErrorStats & vbCrLf
    Text4.SelStart = Len(Text4.Text)
End Sub
Public Function SysTray()
TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Me.hwnd
    TheForm.hIcon = Image1.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    TheForm.szTip = Me.Caption
    
    Shell_NotifyIcon NIM_ADD, TheForm
    Shell_NotifyIcon NIM_MODIFY, TheForm
End Function

Public Function SysTray2()
TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Me.hwnd
    TheForm.hIcon = Image2.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    TheForm.szTip = Text7.Text & " Messages" & " as of " & Text6.Text
    
    Shell_NotifyIcon NIM_MODIFY, TheForm
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
    
Select Case msg
Case WM_LBUTTONDBLCLK    '515 restore form window
If Me.Visible = True Then
Me.Visible = False
Else
Me.Visible = True
Me.SetFocus
Text4.SelStart = Len(Text4.Text)
End If
Case WM_RBUTTONUP        '517 display popup menu
Me.PopupMenu Me.MenuFile
End Select
End Sub
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

End Function
