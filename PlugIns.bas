Attribute VB_Name = "PlugIns"
' Plugin Subs - (C) Dean Camera, 2004

Global PlugInsLST As New Collection
Dim LoadPlugTemp As Object
Dim PlugNameTemp

Private Function LoadPlug(ClassName As String)
    Set LoadPlugTemp = CreateObject(ClassName)
    PlugNameTemp = LoadPlugTemp.FriendlyName
End Function

Function GoPlug(Key, Parametres As String)
    On Error GoTo AlertError
    For I = 1 To PlugInsLST.Count
        If PlugInsLST.Item(I).FriendlyName = Key Then
            PlugInsLST.Item(I).doaction (Parametres)
        End If
    Next
    
    Exit Function
AlertError:
    Debug.Print "Unable to commence action - " & Err.Description
End Function

Sub LoadPlugs()
    On Error GoTo quitsub
    
    Static PTA As String
    plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
    If plugcount = 0 Then
        Debug.Print "No Plugins Found."
        Exit Sub
    Else
        Debug.Print plugcount & " Plugin(s) Found."
        LoggingSubs.AddLogText plugcount & " Plugin(s) Found."
    End If
    
    On Error GoTo 0
    
    CustomMenu.MNUPlugEmpty.Visible = True
    
On Error Resume Next
    For I = 1 To plugcount
        PTA = GetSetting("ComTalk", "Plugins", "Plugin " & I, "")
        Debug.Print "Loading Plugin No. " & I & " (" & PTA & ")"
        If PTA = "" Then Exit Sub
        LoadPlug PTA
        LoggingSubs.AddLogText "Loaded Plugin No. " & I & " (" & PTA & ")"
        PlugInsLST.Add LoadPlugTemp, PlugNameTemp
        If PlugInsLST(I).ShowInMenu = True Then
            Load CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count + 1)
            CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count).Caption = PlugNameTemp
            CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count).Visible = True
            CustomMenu.MNUPlugEmpty.Visible = False
        End If
    Next
    
quitsub:
End Sub

Function KillPlugs()
    On Error Resume Next
    For I = 1 To PlugInsLST.Count
        PlugInsLST.Item(I).KillMe
        PlugInsLST.Item(I) = Nothing
    Next
End Function

Function CheckPassText(Txt As String)
    On Error Resume Next
    For I = 1 To PlugInsLST.Count
        If PlugInsLST(I).PassBeforeSay = True Then
            texttemp = PlugInsLST(I).doaction(Txt)
        End If
    Next
    
    If texttemp <> "" Then CheckPassText = texttemp Else CheckPassText = Txt
End Function

Sub PutPlugsInMenu()
On Error Resume Next
    plugcount = GetSetting("ComTalk", "Plugins", "Count", 0)
    For I = 1 To plugcount
        PTA = GetSetting("ComTalk", "Plugins", "Plugin " & I, "")
        If PTA = "" Then Exit Sub
        If PlugInsLST(I).ShowInMenu = True Then
            Load CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count + 1)
            CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count).Caption = PlugInsLST(I).FriendlyName
            CustomMenu.MNUPlugs(CustomMenu.MNUPlugs.Count).Visible = True
            CustomMenu.MNUPlugEmpty.Visible = False
            Debug.Print "Placed """ & PlugInsLST(I).FriendlyName & """ in Plugin menu."
        End If
    Next

For I = 1 To CustomMenu.MNUPlugs.Count
If CustomMenu.MNUPlugs(I).Caption = "PLUG" Then CustomMenu.MNUPlugs(I).Visible = False ' Remove any failed plugins that made new plugin menu items
Next
End Sub
