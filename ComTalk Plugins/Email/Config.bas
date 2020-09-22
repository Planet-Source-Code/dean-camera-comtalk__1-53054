Attribute VB_Name = "Config"
Const PACN = "PopMail.MailMain" ' Project & Class Name

Sub Main()
plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
For i = 1 To plugcount
If GetSetting("ComTalk", "Plugins", "Plugin " & i, "") = PACN Then
MailIcon.Option1.Value = True
End If
Next

CTOpen = GetSetting("ComTalk", "Program", "IsOpen", 0)
If CTOpen = 0 Then
MailIcon.Show
End If
End Sub
