VERSION 5.00
Begin VB.UserControl USEFULLctrl 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Usefull.ctx":0000
   ScaleHeight     =   645
   ScaleWidth      =   360
   ToolboxBitmap   =   "Usefull.ctx":06BA
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "API"
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "USEFULLctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'           ----------------------------------
'    ______| Usefull Ctrl By Dean Camera 2002 |_____
'   |  Comprised of misc. code found on the internet|
'    -----------------------------------------------


Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (Os As OsVersionInfo) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemMetrics Lib "User32" (ByVal WhatInfo As Integer)
Private Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function SetSysColors Lib "User32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function ShowCursor& Lib "User32" (ByVal bShow As Long)
Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FlashWindow Lib "User32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function ExitWindowsEx Lib "User32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SwapMouseButton& Lib "User32" (ByVal bSwap As Long)
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal Text As String, ByVal hIcon As Long) As Long
Private Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowWord Lib "User32" (ByVal hWnd As Long, ByVal _
    nIndex As Long, ByVal nNewWord As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FindWindowEx Lib "User32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
Private Declare Function SHFileOperation Lib _
    "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const HWND_TOP = 0
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const LANG_USER_DEFAULT = &H400&
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const COLOR_SCROLLBAR = 0 'The Scrollbar color
Private Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Private Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Private Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Private Const COLOR_MENU = 4 'Menu
Private Const COLOR_WINDOW = 5 'Windows background
Private Const COLOR_WINDOWFRAME = 6 'Window frame
Private Const COLOR_MENUTEXT = 7 'Window Text
Private Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Private Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Private Const COLOR_ACTIVEBORDER = 10 'Border of active window
Private Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Private Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Private Const COLOR_HIGHLIGHT = 13 'Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Private Const COLOR_BTNFACE = 15 'Button
Private Const COLOR_BTNSHADOW = 16 '3D shading of button
Private Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Private Const COLOR_BTNTEXT = 18 'Button text
Private Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Private Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_LWIN = &H5B
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LogOff As Long = 0
Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TRANSPARENT = &H20&
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SCREENSAVE = &HF140&
Private Const ALTERNATE = 1
Private Const WINDING = 2
Private Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Private Const SPI_SCREENSAVERRUNNING = 97
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPI_SETDESKWALLPAPER = 20
Private Const MAX_PATH = 260

Dim lngReturnValue As Long
Dim nid As NOTIFYICONDATA
Dim strBuffer As String
Dim lngBufSize As Long
Dim lngStatus As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Type OsVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Enum CDState
    TrayOpen = 1
    TrayClosed = 2
End Enum

Public Enum RemoveAction
    Recycle = 1
    Delete = 2
End Enum

Public Enum DT
    AddRemove = 1
    Display = 2
    Accessability = 3
    Regional = 4
    Joystick = 5
    Mouse = 6
    Keyboard = 7
    Printers = 8
    Fonts = 9
    Multimedia = 10
    Modem = 11
    Network = 12
    Passwords = 13
    System = 14
    AddNewHardware = 15
    DateAndTime = 16
    Sounds = 17
    QuickTime = 18
    Dialing = 19
    Power = 20
    Internet = 21
End Enum





Function ACTN_WIN_ShutDown()
    ExitWindowsEx EWX_SHUTDOWN, 0&
End Function

Function ACTN_WIN_Reboot()
    ACTN_WIN_Reboot = ExitWindowsEx(EWX_REBOOT, 0&)
End Function

Function ACTN_PRG_AlternateFlash()
    ACTN_PRG_AlternateFlash = FlashWindow(UserControl.Parent.hWnd, True)
End Function

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 345
    UserControl.Height = 585
End Sub

Function ACTN_PRG_MakeFormCircular(RedrawYesNo As Boolean)
    Dim lngRegion As Long
    Dim lngReturn As Long
    Dim lngFormWidth As Long
    Dim lngFormHeight As Long
    lngFormWidth = UserControl.Parent.Width / Screen.TwipsPerPixelX
    lngFormHeight = UserControl.Parent.Height / Screen.TwipsPerPixelY
    lngRegion = CreateEllipticRgn(1, 1, lngFormWidth, lngFormHeight)
    lngReturn = SetWindowRgn(UserControl.Parent.hWnd, lngRegion, RedrawYesNo)
End Function

Function ACTN_PRG_MakeFormTransparent()
    SetWindowLong UserControl.Parent.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    SetWindowPos UserControl.Parent.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
End Function

Function ACTN_WIN_LogOff()
    ACTN_WIN_LogOff = ExitWindowsEx(EWX_LogOff, 0&)
End Function

Function ACTN_PRG_StayOnTop(YesNo As Boolean)
    If YesNo = True Then
        ACTN_PRG_StayOnTop = SetWindowPos(UserControl.Parent.hWnd, -1, 0, 0, 0, 0, 3)
    Else
        ACTN_PRG_StayOnTop = SetWindowPos(UserControl.Parent.hWnd, -2, 0, 0, 0, 0, 3)
    End If
End Function

Function ACTN_WIN_ShowInTaskList(YesNo As Boolean)
    Static lngProcessID As Long
    Static lngReturn As Long
    If YesNo = True Then
        lngProcessID = GetCurrentProcessId()
        lngReturn = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
    Else
        lngProcessID = GetCurrentProcessId()
        lngReturn = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
    End If
End Function

Function ACTN_WIN_MinimizeAllWindows()
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function

Function ACTN_WIN_ShowMouseCursor(YesNo As Boolean)
    If YesNo = True Then
        ShowCursor (True)
    Else
        ShowCursor (False)
    End If
End Function

Function ACTN_WIN_ChangeTitleBarColour(RedAmount As Integer, GreenAmount As Integer, BlueAmount As Integer)
    Static lngReturn As Long
    ACTN_WIN_ChangeTitleBar = SetSysColors(1, COLOR_ACTIVECAPTION, RGB(RedAmount, GreenAmount, BlueAmount))
End Function

Function ACTN_DRV_ChangeDriveName(Drivepath As String, NewName As String)
    SetVolumeLabel Drivepath, NewName
End Function

Function ACTN_WIN_SetCursorPos(Xpos As Integer, Ypos As Integer)
    Dim lpRect As RECT
    GetWindowRect UserControl.Parent.hWnd, lpRect
    SetCursorPos Xpos, Ypos
End Function

Function ACTN_PRG_SaveTextToFile(ByVal filename As String, ByVal Text As String) As Boolean
    Dim Handle As Integer
    Dim isOpen As Boolean
    On Error GoTo SaveTextFile_ErrHandler
    Handle = FreeFile
    Open filename For Output As #Handle
    isOpen = True
    Print #Handle, Text;
    SaveTextFile = True
SaveTextFile_ErrHandler:
    If isOpen Then Close #Handle
End Function

Function ACTN_WIN_OpenStartMenu()
    keybd_event VK_LWIN, 0, 0, 0
    keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
End Function


Function PROP_WIN_WindowsDirectory()
    PROP_WIN_WindowsDirectory = Environ("windir")
End Function

Function PROP_WIN_Username()
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    PROP_WIN_Username = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

Function PROP_WIN_Resolution()
    Dim intWidth As Integer
    Dim intHeight As Integer
    intWidth = Screen.Width \ Screen.TwipsPerPixelX
    intHeight = Screen.Height \ Screen.TwipsPerPixelY
    PROP_WIN_Resolution = Str$(intWidth) + " x" + Str$(intHeight)
End Function

Function ACTN_WIN_HideTaskBar(YesNo As Boolean)
    If YesNo = True Then
        ACTN_WIN_HideTaskBar = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    Else
        ACTN_WIN_HideTaskBar = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Function


Function ACTN_PRG_ScrollTextOnePlace(strText As String)
    strText = (Right$(strText, Len(strText) - 1)) & Left$(strText, 1)
    ACTN_PRG_ScrollTextOnePlace = strText
End Function

Function ACTN_WIN_SetCDState(State As CDState)
    If State = TrayOpen Then
        Call mciSendString("Set CDAudio Door Open", 0&, 0&, 0&)
    ElseIf State = TrayClosed Then
        Call mciSendString("Set CDAudio Door Closed", 0&, 0&, 0&)
    End If
End Function

Function ACTN_WIN_MonitorStandby()
    ACTN_WIN_MonitorStandby = SendMessage(UserControl.Parent.hWnd, &H112, &HF170, 1)
End Function

Function ACTN_WIN_SwapMouseButtons(YesNo As Boolean)
    SwapMouseButton YesNo
End Function

Function ACTN_WIN_StartScreenSaver()
    Call SendMessage(UserControl.Parent.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Function

Function ACTN_WIN_ShutdownDlg()
    SHShutDownDialog 0
End Function

Function ACTN_WIN_DesktopPropertiesDlg()
    Shell "C:\Windows\control.exe" & " desk.cpl", 1
End Function

Function ACTN_WIN_ChangeComName(NewComName As String)
    ACTN_WIN_ChangeComName = SetComputerName(NewComName)
End Function

Function PROP_SND_SoundCardEnabled()
    sndinfo = waveOutGetNumDevs()
    If sndinfo > 1 Then
        PROP_SND_SoundCardEnabled = True
    Else
        PROP_SND_SoundCardEnabled = False
    End If
End Function

Function ACTN_PRG_CoverEntireScreen()
    Dim cx As Long
    Dim cy As Long
    Dim RetVal As Long
    If UserControl.Parent.WindowState = vbMaximized Then
        UserControl.Parent.WindowState = vbNormal
    End If
    RetVal = SetWindowPos(UserControl.Parent.hWnd, HWND_TOP, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
End Function

Function ACTN_WIN_ShowStartButton(YesNo As Boolean)
    If YesNo = True Then
        OurParent& = FindWindow("Shell_TrayWnd", "")
        OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
        ShowWindow OurHandle&, 5
    Else
        OurParent& = FindWindow("Shell_TrayWnd", "")
        OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
        ShowWindow OurHandle&, 0
    End If
End Function


Function ACTN_WIN_ShowSystemClock(YesNo As Boolean)
    Dim FindClass As Long, FindParent As Long, Handle As Long
    If YesNo = True Then
        FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
        FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
        Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
        ShowWindow Handle&, 1
    Else
        FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
        FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
        Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
        ShowWindow Handle&, 0
    End If
End Function

Function ACTN_WIN_DisableCTRLALTDEL(YesNo As Boolean)
    Dim ret As Integer
    Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, YesNo, pOld, 0)
End Function

Function PROP_WIN_WindowsRunTimeMinutes()
    PROP_WIN_WindowsRunTimeMinutes = Format(GetTickCount / 60000, "0")
End Function

Function ACTN_FLE_RecycleFile(FilePathAndName As String, Action As RemoveAction) As Boolean
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    On Error GoTo RemoveFile_Err
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = filename
        If Action = rfRecycle Then
            .fFlags = FOF_ALLOWUNDO + FOF_CREATEPROGRESSDLG
        Else
            .fFlags = FO_DELETE + FOF_CREATEPROGRESSDLG
        End If
    End With
    lReturn = SHFileOperation(FileOperation)
    If lReturn <> 0 Then
        ACTN_FLE_RecycleFile = False
    Else
        ACTN_FLE_RecycleFile = True
    End If
    Exit Function
RemoveFile_Err:
    ACTN_FLE_RecycleFile = False
End Function

Function ACTN_PRG_AboutDlg(Title As String, AboutText As String)
    ShellAbout UserControl.Parent.hWnd, Title, AboutText, UserControl.Parent.Icon
End Function

Function ACTN_PRG_FatalExit(ErrorText As String)
    FatalAppExit 1, ErrorText
End Function

Function PROP_WIN_SystemDirectory()
    strBuffer = Space$(MAX_PATH)
    buffset = GetSystemDirectory(strBuffer, MAX_PATH)
    PROP_WIN_SystemDirectory = Left$(strBuffer, Len(strBuffer) - 1)
End Function

Function ACTN_PRG_BringWindowToTop()
    BringWindowToTop UserControl.Parent.hWnd
End Function

Function PROP_FLE_FileOrPathExists(Path As String, Optional File As String)
    If File = "" Then
        PROP_FLE_FileExists = Not (Dir(Path) = "")
    Else
        PROP_FLE_FileExists = Not (Dir(Path & File) = "")
    End If
End Function

Function ACTN_WIN_OpenControlPanel(TypeOfFocus As VbAppWinStyle)
    Shell "rundll32.exe shell32.dll,Control_RunDLL", TypeOfFocus
End Function

Function ACTN_WIN_OpenControlPanelDialog(DialogName As DT, TypeOfFocus As VbAppWinStyle)
    Select Case DialogName
    Case 1
        Shell "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", TypeOfFocus
    Case 2
        Shell "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,5", TypeOfFocus
    Case 3
        Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", TypeOfFocus
    Case 4
        Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", TypeOfFocus
    Case 5
        Shell "rundll32.exe shell32.dll,Control_RunDLL joy.cpl", TypeOfFocus
    Case 6
        Shell "rundll32.exe shell32.dll,Control_RunDLL mainfrm.cpl @0", TypeOfFocus
    Case 7
        Shell "rundll32.exe shell32.dll,Control_RunDLL mainfrm.cpl @1", TypeOfFocus
    Case 8
        Shell "rundll32.exe shell32.dll,Control_RunDLL mainfrm.cpl @2", TypeOfFocus
    Case 9
        Shell "rundll32.exe shell32.dll,Control_RunDLL mainfrm.cpl @3", TypeOfFocus
    Case 10
        Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", TypeOfFocus
    Case 11
        Shell "rundll32.exe shell32.dll,Control_RunDLL modem.cpl", TypeOfFocus
    Case 12
        Shell "rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", TypeOfFocus
    Case 13
        Shell "rundll32.exe shell32.dll,Control_RunDLL password.cpl", TypeOfFocus
    Case 14
        Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", TypeOfFocus
    Case 15
        Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", TypeOfFocus
    Case 16
        Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", TypeOfFocus
    Case 17
        Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", TypeOfFocus
    Case 18
        Shell "rundll32.exe shell32.dll,Control_RunDLL quicktime.cpl", TypeOfFocus
    Case 19
        Shell "rundll32.exe shell32.dll,Control_RunDLL Telephon.cpl", TypeOfFocus
    Case 20
        Shell "rundll32.exe shell32.dll,Control_RunDLL powercfg.cpl", TypeOfFocus
    Case 21
        Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl", TypeOfFocus
    End Select
End Function

Function PROP_WIN_ErrorDescription(ErrorNumber As Integer)
    PROP_WIN_ErrorDescription = GetLastErrorStr(Val(ErrorNumber))
End Function

Function PROP_DRV_CheckKBFreeSpace(DriveName As String)
    Dim free_Space As Long
    ChDrive DriveName
    Dim numSectorsPerCluster As Long
    Dim numBytesPerSector As Long
    Dim numFreeClusters As Long
    Dim numTotalClusters As Long
    Dim success As Boolean
    success = GetDiskFreeSpace(DriveName, numSectorsPerCluster, numBytesPerSector, numFreeClusters, numTotalClusters)
    free_Space = numSectorsPerCluster * numBytesPerSector * numFreeClusters
    PROP_DRV_CheckFreeSpace = Format(Str$(free_Space / 1024), "###,### ")
End Function


Private Function GetLastErrorStr(dwErrCode As Long) As String
    Static sMsgBuf As String * 257, dwLen As Long
    dwLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM _
        Or FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, _
        dwErrCode, LANG_USER_DEFAULT, ByVal sMsgBuf, 256&, 0&)
    If dwLen Then GetLastErrorStr = Left$(sMsgBuf, dwLen)
End Function
