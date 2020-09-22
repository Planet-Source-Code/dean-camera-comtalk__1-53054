Attribute VB_Name = "Resolution"
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
 
Const HWND_BROADCAST = &HFFFF&
Const WM_DISPLAYCHANGE = &H7E&
Const SPI_SETNONCLIENTMETRICS = 42
Const CCDEVICENAME = 32
Const CCFORMNAME = 32

Private Type DEVMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Sub ChangeRes(Dimensions As String)
Debug.Print "Change resolution: " & Dimensions

Static ResHeight As Integer
Static ResWidth As Integer

For d = 1 To Len(Dimensions)
    If Mid(Dimensions, d, 1) = "x" Then
        ResWidth = Mid(Dimensions, 1, d - 1)
        ResHeight = Mid(Dimensions, d + 1)
        SetRes ResHeight, ResWidth
        Exit For
    End If
Next
End Sub

Private Sub SetRes(Height As Integer, Width As Integer)
Dim DevM    As DEVMODE
Dim lResult As Long
Dim iAns    As Integer
'
' Retrieve info about the current graphics mode
' on the current display device.
'
lResult = EnumDisplaySettings(0, 0, DevM)
'
' Set the new resolution. Don't change the color
' depth so a restart is not necessary.
'
With DevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
        .dmPelsWidth = Width  'ScreenWidth
        .dmPelsHeight = Height 'ScreenHeight
End With
'
' Change the display settings to the specified graphics mode.
'
lResult = ChangeDisplaySettings(DevM, CDS_TEST)
Select Case lResult
    Case DISP_CHANGE_RESTART
        iAns = MsgBox("You must restart your computer to apply these changes." & _
            vbCrLf & vbCrLf & "Do you want to restart now?", _
            vbYesNo + vbSystemModal, "Screen Resolution")
        If iAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
        Call ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        Call SendMessage(HWND_BROADCAST, WM_DISPLAYCHANGE, SPI_SETNONCLIENTMETRICS, ByVal 0&)
    Case Else
        MsgBox "Screen resolution not supported", vbSystemModal, "Error"
End Select
End Sub

