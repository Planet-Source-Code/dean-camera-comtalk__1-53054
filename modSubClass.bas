Attribute VB_Name = "modSubClass"
Option Explicit
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Constans
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*

'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== APIs
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'-- hook
Private Const WH_KEYBOARD = 2
Private Const HC_ACTION = 0
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private m_KeyHookPtr()     As Long
Private m_KeyHookCount     As Long
Private m_HookAddress      As Long
Private m_oldHook          As Long

Public Function HookKeyboard(ByVal objThis As cMenus)
    If (m_HookAddress = 0) Then
        m_HookAddress = pvLongFromLong(AddressOf KeyboardProc)
        m_oldHook = SetWindowsHookEx(WH_KEYBOARD, m_HookAddress, 0&, GetCurrentThreadId())
    End If
    Dim lPtr As Long
    lPtr = ObjPtr(objThis)
    If (m_KeyHookCount > 0) Then
        Dim I As Long
        For I = 1 To m_KeyHookCount
            If (m_KeyHookPtr(I) = lPtr) Then
                Exit Function
            End If
        Next
    End If
    m_KeyHookCount = m_KeyHookCount + 1
    ReDim Preserve m_KeyHookPtr(1 To m_KeyHookCount) As Long
    m_KeyHookPtr(m_KeyHookCount) = lPtr
End Function
Public Function RemoveHookKeyboard(ByVal objThis As cMenus)
    Dim lPtr As Long
    lPtr = ObjPtr(objThis)
    If (m_KeyHookCount > 0) Then
        Dim I As Long
        Dim bFound As Boolean
        Dim hHook
        For I = 1 To m_KeyHookCount
            If (bFound) Then
                If (I <> m_KeyHookCount) Then
                    m_KeyHookPtr(I - 1) = m_KeyHookPtr(I)
                End If
            ElseIf (m_KeyHookPtr(I) = lPtr) Then
                bFound = True
            End If
        Next
        m_KeyHookCount = m_KeyHookCount - 1
        If (m_KeyHookCount = 0) Then
            Erase m_KeyHookPtr
            Call UnhookWindowsHookEx(m_oldHook)
            m_oldHook = 0
        Else
            ReDim Preserve m_KeyHookPtr(1 To m_KeyHookCount) As Long
        End If
    End If
End Function
Private Function pvObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    If lPtr <> 0 Then
        Call CopyMemory(oTemp, lPtr, 4)
        Set pvObjectFromPtr = oTemp
        Call CopyMemory(oTemp, 0&, 4)
    End If
End Function
Private Function pvLongFromLong(ByVal lngThis As Long) As Long
    pvLongFromLong = lngThis
End Function
Private Function KeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If ncode = HC_ACTION Then
        If (Not ((lParam And &H80000000) = &H80000000)) Then
            Dim I As Long
            Dim currClass As cMenus
            Dim bShift As Boolean
            Dim bAlt As Boolean
            Dim bCtrl As Boolean
            Dim lKeyStr As String
            bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
            bAlt = ((lParam And &H20000000) = &H20000000)
            bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
            lKeyStr = GetKeyName(wParam, bAlt, bCtrl, bShift)
            For I = 1 To m_KeyHookCount
                Set currClass = pvObjectFromPtr(m_KeyHookPtr(I))
                If (currClass.KeyAccelPressed(lKeyStr)) Then
                    If (currClass.ConsumeKeys) Then
                        KeyboardProc = 1
                        GoTo gTerminate
                    End If
                End If
            Next
            KeyboardProc = CallNextHookEx(m_oldHook, ncode, wParam, lParam)
gTerminate:
            Set currClass = Nothing
        End If
    End If
End Function
Public Function GetKeyName(ByVal KeyCode As KeyCodeConstants, Optional Alt As Boolean, Optional Ctrl As Boolean, Optional Shift As Boolean) As String
    If ((KeyCode >= vbKeyF1) And (KeyCode <= vbKeyF16)) Then
        GetKeyName = "F" & (KeyCode - vbKeyF1) + 1
    ElseIf ((KeyCode >= vbKeyA) And (KeyCode <= vbKeyZ)) Then
        GetKeyName = Chr((KeyCode - vbKeyA) + 65)
    ElseIf ((KeyCode >= vbKey0) And (KeyCode <= vbKey9)) Then
        GetKeyName = (KeyCode - vbKey0)
    ElseIf ((KeyCode >= vbKeyNumpad0) And (KeyCode <= vbKeyNumpad9)) Then
        GetKeyName = "Numpad" & (KeyCode - vbKeyNumpad0)
    ElseIf (KeyCode = vbKeyDelete) Then
        GetKeyName = "Delete"
    ElseIf (KeyCode = vbKeyTab) Then
        GetKeyName = "Tab"
    ElseIf (KeyCode = vbKeyEscape) Then
        GetKeyName = "Escape"
    End If
    If (GetKeyName <> vbNullString) Then
        Dim strLeft As String
        If (Alt) Then
            strLeft = "Alt"
        End If
        If (Ctrl) Then
            If (strLeft = vbNullString) Then
                strLeft = "Ctrl"
            Else
                strLeft = strLeft & "+" & "Ctrl"
            End If
        End If
        If (Shift) Then
            If (strLeft = vbNullString) Then
                strLeft = "Shift"
            Else
                strLeft = strLeft & "+" & "Shift"
            End If
        End If
        If (strLeft <> vbNullString) Then
            GetKeyName = strLeft & "+" & GetKeyName
        End If
    End If
End Function
