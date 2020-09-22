Attribute VB_Name = "AtomicClock"
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" (ByVal sAgent As String, _
    ByVal lAccessType As Long, ByVal sProxyName As String, _
    ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
    Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
    ByVal sUrl As String, ByVal sHeaders As String, _
    ByVal lLength As Long, ByVal lFlags As Long, _
    ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As Long, ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
    As Integer

Private Declare Function InternetCloseHandle _
    Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function GetTimeZoneInformation& Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION)


Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public m_StrAtomicTime As String

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpm_ndwsFlags As Long, ByVal dwReserved As Long) As Long

Private Function IsNetConnectOnline() As Boolean
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function

Private Function DoAtomicTime() As Boolean ' Get current time from Atomic server
    Static TZI As TIME_ZONE_INFORMATION    ' This sub is not properly commented as
    Static x As Single                     ' it was not writen by me, I only modified it
    Static Ntime As Date
    Static i, n As Integer
    Static lRet As Long
    Static sRet As String
    Static bDaylightSavings As Boolean
    Static iMonthNow As Integer, iDayNow As Integer
    Static iStandardMonth As Integer, iStandardDay As Integer
    Static iDaylightMonth As Integer, iDaylightDay As Integer
    
    m_StrAtomicTime = ""                            ' \
    Call GetTimeZoneInformation(TZI)                ' |
                                                    ' |
    iStandardMonth = TZI.StandardDate.wMonth        ' |
    iStandardDay = TZI.StandardDate.wDay            ' |
                                                    ' |- Get daylight savings time data
    iDaylightMonth = TZI.DaylightDate.wMonth        ' |
    iDaylightDay = TZI.DaylightDate.wDay            ' |
                                                    ' |
    iMonthNow = Month(Now)                          ' |
    iDayNow = Day(Now)                              ' /
    
    If iStandardMonth = iMonthNow Then                          ' \
        bDaylightSavings = (iDayNow < iStandardDay)             ' |
    ElseIf iDaylightMonth = iMonthNow Then                      ' |
        bDaylightSavings = (iDayNow >= iDaylightDay)            ' |
    Else                                                        ' | - See if user's computer is on Daylight Savings Time (+1 Hour)
        If iDaylightMonth < iStandardMonth Then                 ' |
            bDaylightSavings = iMonthNow > iDaylightMonth _
            And iMonthNow < iStandardMonth                      ' |
        Else
            bDaylightSavings = iMonthNow > iDaylightMonth _
            Or iMonthNow < iStandardMonth                       ' /
        End If
    End If
    lRet = TZI.Bias
    
    If bDaylightSavings Then
        lRet = lRet + TZI.DaylightBias
    Else
        lRet = lRet + TZI.StandardBias
    End If
    
    x = lRet / 60
    x = (x / 24)
    
    sRet = OpenURL("http://tycho.usno.navy.mil/cgi-bin/timer.pl")   ' Get atomic time data from website
    i = InStr(2, sRet, GetSearchString) ' CHANGE THIS IF PAGE LAYOUT CHANGES - MUST BE THE TEXT AFTER THE UNIVERSAL TIME (UTC)
                               ' This has been "Universal", "U" and "UTC". For some reason, the webmasters keep changing the
                               ' page format, making it VERY hard for Austomated programs
    If i <> 0 Then
        sRet = Left$(sRet, i)
        Debug.Print sRet
        n = InStrRev(sRet, ",")
        If n <> 0 Then
            Ntime = CDate(Trim(Mid$(sRet, (n + 1), (i - (n + 1))))) ' Set the time
            Time = (Ntime - x)
            m_StrAtomicTime = CStr(Time)
            DoAtomicTime = True
        End If
    End If
End Function

Public Function GetAtomicClock()
    DNCC = GetSetting("ComTalk", "Options", "DoNotCheckINet", 0)
    If DNCC = 0 Then ' See if ComTalk should check if user is online
        If IsNetConnectOnline = False Then ' Check if user is online
                SpeakText "You do not seem to be currently connected to the net. Please connect and try again."
            Exit Function
        End If
    End If
    
        SpeakText "Attempting to Synchronize System Clock to U.S. Naval Atomic Clock..."
    
    If DoAtomicTime Then
        mainfrm.UC.ACTN_WIN_HideTaskBar True  '   \ Refresh taskbar
        mainfrm.UC.ACTN_WIN_HideTaskBar False '   / to show new time
            SpeakText "Your system time was changed to " & m_StrAtomicTime
    Else
            SpeakText "The attempt to synchronize your system time failed. " & Err.Description
    End If
End Function

Function GetSearchString() As String ' Get the string that ComTalk should search for to find the Universal Time
If INI.FoundExtraFile = True Then ' ComTalk.ini exists?
GetSearchString = INI.GetExtra("AtomSync", "SearchInPage")

For i = 1 To Len(GetSearchString)
If Mid(GetSearchString, i, 1) = "\" Then GetSearchString = Mid(GetSearchString, 1, i - 1) & vbNewLine & Mid(GetSearchString, i + 1) ' "\" Means VBNewLine
If Mid(GetSearchString, i, 1) = "|" Then GetSearchString = Mid(GetSearchString, 1, i - 1) & vbCrLf & Mid(GetSearchString, i + 1) ' "|" Means New Line (different from VBNewLine, put here just in case)
If Mid(GetSearchString, i, 1) = "_" Then GetSearchString = Mid(GetSearchString, 1, i - 1) & " " & Mid(GetSearchString, i + 1) ' "_" means space, not really needed but makes it look better
Next
Else
GetSearchString = " UTC" ' Use standard string
End If

Debug.Print "AtomSync String to look for on page: " & GetSearchString
End Function
