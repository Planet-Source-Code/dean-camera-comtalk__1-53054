Attribute VB_Name = "Juked"
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_LENGTH = 145
    Dim Temp
    Dim ret As String
    
Function GetWindowsDir()
        GetWindowsDir = Environ("windir")
End Function

Function MyCharacter(Optional DONTUSEPATH As Boolean)
MyCharacter = GetSetting("ComTalk", "Options", "MyCharacter", "CharacterFail")
End Function
