Attribute VB_Name = "INI"
Option Explicit
Dim TempLan As String
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Function ReadWriteINI(Mode As String, tmpSecname As String, tmpKeyname As String, iniFile As String, Optional tmpKeyValue) As String ' Read the strings from the INI file
    Dim tmpString As String
    Dim filename As String
    Dim secname As String
    Dim keyname As String
    Dim keyvalue As String
    Dim anInt
    Dim defaultkey As String
    
    On Error GoTo ReadWriteINIError
    '
    ' *** set the return value to OK
    'ReadWriteINI = "OK"
    ' *** test for good data to work with
    If IsNull(Mode) Or Len(Mode) = 0 Then
        ReadWriteINI = "ERROR MODE"    ' Set the return value
        Exit Function
    End If
    If IsNull(tmpSecname) Or Len(tmpSecname) = 0 Then
        ReadWriteINI = "ERROR Secname" ' Set the return value
        Exit Function
    End If
    If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then
        ReadWriteINI = "ERROR Keyname" ' Set the return value
        Exit Function
    End If
    ' *** set the ini file name
    filename = iniFile
    '
    '
    ' ******* WRITE MODE *************************************
    If UCase(Mode) = "WRITE" Then
        If IsNull(tmpKeyValue) Or Len(tmpKeyValue) = 0 Then
            ReadWriteINI = "ERROR KeyValue"
            Exit Function
        Else
            
            secname = tmpSecname
            keyname = tmpKeyname
            keyvalue = tmpKeyValue
            anInt = WritePrivateProfileString(secname, keyname, keyvalue, filename)
        End If
    End If
    ' *******************************************************
    '
    ' *******  READ MODE *************************************
    If UCase(Mode) = "GET" Then
        
        secname = tmpSecname
        keyname = tmpKeyname
        defaultkey = "Failed"
        keyvalue = String$(50, 32)
        anInt = GetPrivateProfileString(secname, keyname, defaultkey, keyvalue, Len(keyvalue), filename)
        If Left(keyvalue, 6) <> "Failed" Then        ' *** got it
            tmpString = keyvalue
            tmpString = RTrim(tmpString)
            tmpString = Left(tmpString, Len(tmpString) - 1)
        Else
            tmpString = "Failed"
        End If
        ReadWriteINI = tmpString
    End If
    Exit Function
    
    ' *******
ReadWriteINIError:
    SpeakError Error
End Function

Private Function NeedSlash() As String ' Check if the path needs a slash at the end
    If Right(App.Path, 1) = "\" Then
        NeedSlash = ""
    Else
        NeedSlash = "\"
    End If
End Function

Function FoundExtraFile() As Boolean ' Check if the ComTalk.ini file exists
    FoundExtraFile = FileExist(App.Path & NeedSlash & "ComTalk.ini")
    Debug.Print "Found ComTalk.ini File: " & FoundExtraFile
End Function

Function GetExtra(iniSection As String, Name As String) As String ' Get a non-numbered string (e.g. "SearchInPage=") from the ComTalk.ini file
    Static PartA As String
    
    PartA = ReadWriteINI("GET", iniSection, Name, App.Path & NeedSlash & "ComTalk.ini")
    
    GetExtra = PartA
End Function

Private Function FileExist(Fname As String) As Boolean ' Check if a file exists
    On Local Error Resume Next
    FileExist = (Dir(Fname) <> "")
End Function
