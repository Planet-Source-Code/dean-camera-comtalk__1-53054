Attribute VB_Name = "LockWorkstation"
Dim EDSBS As New CCrypto

Public Sub LockStation() ' Locks the user's computer
If NoLockDown = True Then Exit Sub ' Don't load if command line parameter to disable it is used
    
    If PWORD = "" Then ' Check if a password has been set
            SpeakError "No password set. Please set a password from the options screen.", vbCritical, "ComTalk"
        Exit Sub
    Else
        Lockdown.FixMyForm
        Lockdown.Show
    End If
    
    If mainfrm.Agent1.Characters("Genie").IdleOn = True Then
        LDUseIdle = True
        mainfrm.Agent1.Characters("Genie").IdleOn = False
    End If
    
    mainfrm.Agent1.Characters("Genie").Hide
    
    GoCharacter = False
End Sub

Public Sub UnlockStation() ' Unlock user's computer
If NoLockDown = True Then Exit Sub ' Don't load if command line parameter to disable it is used
    
    On Error Resume Next
    GoCharacter = True
    mainfrm.Agent1.Characters("Genie").Show
        SpeakText "Workstation Unlocked."
    Lockdown.Hide
    
    If LDUseIdle = True Then
        mainfrm.Agent1.Characters("Genie").IdleOn = True
    End If
End Sub

Public Function DecodePassword()
On Error Resume Next
    password = ""
    Temp = GetSetting("ComTalk", "Options", "Lock Station PWord", "")
    password = EDSBS.Decrypt(Temp, "CTALK")
    DecodePassword = password
    PWORD = password
    Debug.Print "LDPASS DECODE: " & Temp
End Function

Public Function EncryptPass(PW As String)
        password = ""
        password = EDSBS.Encrypt(PW, "CTALK")
        Debug.Print "LDPASS ENCODE: " & password
        SaveSetting "ComTalk", "Options", "Lock Station PWord", password
End Function

