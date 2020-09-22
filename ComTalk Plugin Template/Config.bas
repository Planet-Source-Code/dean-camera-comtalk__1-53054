Attribute VB_Name = "Config"
Global Const PACN = "TextFix.TextFixMain" ' Project & Class Name
Global Const MEMADEBY = "Dean Camera, 2002" ' Plugin Author
Global Const MEREQIREMENTS = "None" ' Plugin Requirements

Sub Main()
plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
For i = 1 To plugcount
If GetSetting("ComTalk", "Plugins", "Plugin " & i, "") = PACN Then
TextFixIcon.Option1.Value = True
End If
Next

CTOpen = GetSetting("ComTalk", "Program", "IsOpen", 0)
If CTOpen = 0 Then
TextFixIcon.Show
End If
End Sub
