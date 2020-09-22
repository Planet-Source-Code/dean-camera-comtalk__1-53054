VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreMenu"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ID            As Long
Public OwnerDrawID   As Long
Public Caption       As String
Public Name          As String
Public HelpText      As String
Public Image         As String
Public ParentHandle  As Long
Public KeyAccel      As String
Public OwnerDraw     As Boolean
Public CheckedStyle  As MenuCheckedStyle
Public HasSub        As Boolean
Public AtRunTime     As Boolean
Public OwnerDrawData As Long
Public Index         As Long
Public Break         As Long
Public hMenu         As Long
Public Separator     As Boolean
Public TopMenu       As Boolean
Public hSubMenu      As Long
Public ParentID      As Long
Public DestroyMe     As Boolean
Public AccelChar     As String
Public Description   As String
Public Tag           As String
Public InVisible     As Boolean
Private Sub Class_Terminate()
    KeyAccel = vbNullString
    Caption = vbNullString
    Name = vbNullString
    Image = vbNullString
    HelpText = vbNullString
    AccelChar = vbNullString
    Description = vbNullString
    Tag = vbNullString
End Sub
