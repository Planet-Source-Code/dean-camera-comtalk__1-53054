Attribute VB_Name = "FirstRun"
Global FRunning As Boolean

Function IsFirstRun() As Boolean ' Check if ComTalk has been run before
    TempSet = GetSetting("ComTalk", "Program", "FirstRun", True)
    If TempSet = True Then IsFirstRun = True Else IsFirstRun = False
End Function

Function CheckForOldEncMethod() ' Checks to see if user has been running a version pervious to 3.0.2, and therefore uses the old ASCII shift encryption method.
    AW = GetSetting("ComTalk", "Program", "Ver302UP", False) ' Check if the user ran an older ComTalk version
    If AW = False Then ' Not first run, but not version 3.0.2 or higher = Older Version ran (warn user)
        elp = GetSetting("ComTalk", "Options", "Lock Station PWord", "") ' get encoded password
        If elp <> "" Then ' If there is a password, warn that it is the ASCII shift type
            SpeakText "\Map=""Warning:""=""WARNING:""\ System detects that this is not your first time running ComTalk, but is your first time running this version. Versions previous to 3.0.2 uses a \map=""As-Key""=""ASCII""\ shift encription method for the LockDown password. This version uses a 40 Bit Cipher, which will not work with the old password method. Your LockDown password has been deleted, please re-enter it into the options screen to convert it into a 40 bit encrypted password."
            SaveSetting "ComTalk", "Options", "Lock Station PWord", "" ' Delete old password
            PWORD = ""
            onnow.ShowBox "Your Lockdown Password has been deleted, due to version incompatibilities.", "WARNING:"
        End If
    End If
    SaveSetting "ComTalk", "Program", "Ver302UP", True
End Function

Function FirstRunText() ' Intro
    SaveSetting "ComTalk", "Program", "Ver302UP", True ' Save setting that ComTalk is Version 3.0.2 or higher

    FRunning = True

        SaveSetting "ComTalk", "Program", "FirstRun", False
        onnow.ShowBox "Thankyou for downloading ComTalk." & vbNewLine & vbNewLine & "Before you begin, you may want to set ComTalk's options.", "Welcome!", "Set Options"

        Action "Blink"
        If FRunning = True Then SpeakText "Hello, and welcome to ComTalk. System detects that this is your first time running the ComTalk program. \pau=900\"
        Action "Blink"
        If FRunning = True Then SpeakText "I'm the character. You can speak to me by holding down the \Map=""Scroll Lock""=""SCROLL-LOCK""\ key, and speaking into the microphone. You can get the phrases from the ComTalk read me file. Before you begin, please make sure you have all the necessary components (Speech-Recognition Engine, etc.) and that you have properly trained your voice. \pau=1500\"
        Action "Blink"
        If FRunning = True Then SpeakText "ComTalk's commands can be accessed by the menu (right-click on the character) and clicking on the command. Plug-Ins appear under the 'Plug-Ins' menu, after installing them. Not all plug-ins appear, depending on their settings, some are hidden.\pau=900\"
        Action "Blink"
        If FRunning = True Then SpeakText "You can set ComTalk to remind you the time every quarter hour, half hour or hour. Other settings are changeable in the option page, including character options. \pau=1500\"
        Action "Blink"
        If FRunning = True Then SpeakText "If you dislike the current character (i.e. me), you can change it (provided you have other characters installed) by right-clicking on the character, and selecting 'Change Character'. Extra characters can be downloaded from such sites as theagentry.com if you have the internet. \pau=900\"
        Action "Blink"
        If FRunning = True Then SpeakText "ComTalk offers many different commands, such as Computer LockDown, Shutdown, Restart and Logoff. To make ComTalk as powerful as possible, you can set your own commands (click 'Voice Actions' from the menu) or use plug-ins (downloadable from the internet). \pau=1000\"
        Action "Blink"
        If FRunning = True Then SpeakText "Thank you for your time. If you wish to interrupt the character from speaking, simply left click on it. \pau=800\"
        Action "Blink"
        If FRunning = True Then SpeakText "ComTalk is © Dean Camera, 2001-2004 \Mrk=1\"
        FRunning = False
End Function
