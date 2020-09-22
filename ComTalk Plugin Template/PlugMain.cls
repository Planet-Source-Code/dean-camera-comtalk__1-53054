VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginClassName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get FriendlyName() As String
FriendlyName = "MyName" ' Friendly Name (Shows in Plugin Menu)
End Property

Public Function DoAction(Parametres)
' Actions Executed

DoAction = "Hello!"
End Function

Public Function KillMe()
On Error Resume Next
Unload PluginMain
Unload PluginMainIcon
End
End Function

Public Property Get ShowInMenu() As Boolean
ShowInMenu = True ' Show Plugin In Menu
End Property

Public Property Get PassBeforeSay() As Boolean
PassBeforeSay = False ' Pass Text to be spoken through plugin "DoAction" first
End Property

Public Property Get ClassName() As String
ClassName = PACN
End Property
