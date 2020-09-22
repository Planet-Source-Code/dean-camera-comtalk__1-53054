VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenus"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'           Owner Draw Menus Version 1.0
'               This class written By
'                  Abdalla Mahmoud
'               (la3toot@hotmail.com)
'       You are freely to use this class in
'        any projects while these comments
'                above are displayed
'-----------------------------------------
'   You have to include also 4 classes with this
'      1)cSuperClass     By : Paul_Caton@hotmail.com
'      2)ISuperClas      By : Paul_Caton@hotmail.com
'      3)cMemDC          By : wqweto@myrealbox.com
'      4)clsStoreMenu    By : Me
'    With this class you can easily make your
'    own menus style , this class gives you
'    a default 3D-Style and calculate menus
'    size itself , but you can use your own
'    styles and sizes your self
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Option Explicit
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Constans
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENOUTER = &H2
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const ODT_MENU = 1
Private Const ODS_SELECTED = &H1
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_MENU = 4
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const DSS_DISABLED = &H20
Private Const DST_ICON = &H3
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_TYPE = &H10
Private Const MF_ItemOwnerDraw = &H100&
Private Const MF_SEPARATOR = &H800
Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUSELECT = &H11F
Private Const WM_ERASEBKGND = &H14
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_COMMAND = &H111
Private Const WM_MENUCHAR = &H120
Private Const MF_MENUBREAK = &H40&
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_DEFAULT = &H1000&
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_DEFAULT = MF_DEFAULT
Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MF_BYCOMMAND = &H0&
Private Const MIIM_CHECKMARKS = &H8&
Private Const MF_SYSMENU = &H2000&
'Private Const WM_INITMENUPOPUP = &H117 ' Used to grab the menus and make them owner-drawn
Private Const MIIM_DATA = &H20&
'===
'Private Const m_ExtraSize     As Long = 36&
'Private Const m_LeftIndent    As Long = 15&
Private Const m_RightIndent     As Long = 9&
Private Const m_RightIndentXP   As Long = 20&
Private Const m_IconSize        As Long = 21&
Private Const m_IconSizeXP      As Long = 24&
Private Const m_CaptionIndent   As Long = 2&
Private Const m_CaptionIndentXP As Long = 8&
Private Const m_AccelIndent     As Long = 6&
Private Const m_SubTriSize      As Long = 15&
'------ Enums
Public Enum MenuCheckedStyle
    mcs_Normal = 0
    mcs_Radio = 1
    mcs_Icon = 2
End Enum
Public Enum MenuDrawStyle
    mds_3D = 0
    mds_XP = 1
    'mcs_Icon = 2
End Enum
Public Enum SpecialColors
    SC_XPBack = 16251903
    SC_XPBorder = 7021576
    SC_XPHighlight = 14073525
    SC_XPFill = 14604246
    SC_XPShadow = 10260108
    SC_XPMargin = 14080990
    SC_XPDarkFill = 11899524
    SC_XPDisabled = 10856101
End Enum
Public Enum MenuOwnerDrawFlags
    modf_Highlighted = &H0
    modf_Checked = &H1
    modf_Enabled = &H2
    modf_TopMenu = &H3
    modf_Clicked = &H4
End Enum
Public Enum ShowPopupMenuConstants
    ' Track popup menu constants:
    TPM_CENTERALIGN = &H4&
    TPM_LEFTALIGN = &H0&
    TPM_LEFTBUTTON = &H0&
    TPM_RIGHTALIGN = &H8&
    TPM_RIGHTBUTTON = &H2&

    TPM_HORIZONTAL = &H0                                   '/* Horz alignment matters more */
    TPM_VERTICAL = &H40                                    '/* Vert alignment matters more */

    ' Win98/2000 menu animation and menu within menu options:
    TPM_RECURSE = &H1&
    TPM_HORPOSANIMATION = &H400&
    TPM_HORNEGANIMATION = &H800&
    TPM_VERPOSANIMATION = &H1000&
    TPM_VERNEGANIMATION = &H2000&
    ' Win2000 only:
    TPM_NOANIMATION = &H4000&
End Enum
'Const m_XP_BackColor As Long = 16251903
'Const m_XP_Border As Long = 7021576
'Const m_XP_Highlight As Long = 14073525
'Const m_XP_Shadow As Long = 10260108
'Const m_XP_Margin As Long = 14080990
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Types
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    itemAction As Long
    itemState As Long
    hWndItem As Long
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    ItemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type
Private Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As String
     cch As Long
End Type
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== API Declartion
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal byPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As Any) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
'Private Declare Function GetMenuInfo Lib "user32.dll" (ByVal hmenu As Long, ByRef LPMENUINFO As MenuInfo) As Long
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Variables
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private m_SubClass As cSuperClass
Implements ISuperClass

Private m_VBMenus      As Collection
Private m_CollMenus    As Collection
Private m_MemDC        As cMemDC
Private m_TmpMemDC     As cMemDC

Private m_hMenu  As Long
Private m_hWnd   As Long

Private m_MenusCount         As Long
Private m_ItemHeight         As Long
Private m_ImageList          As Object
Private m_CurrentMenuIndex   As Long
Private m_hSystemMenu        As Long
Private m_BackColor          As OLE_COLOR
Private m_ForeColor          As OLE_COLOR
Private m_HighlightColor     As OLE_COLOR
Private m_HighlightForeColor As OLE_COLOR
Private m_Font               As StdFont
Private m_ItemOwnerDrawAll   As Boolean
Private m_BackGroundPicture  As StdPicture
Private m_BackGroundWidth    As Long
Private m_BackGroundHeight   As Long
Private m_DrawStyle          As MenuDrawStyle
Private m_TopMenus           As Boolean
Private m_DrawBackGround     As Boolean
Private m_RightToLeft        As Boolean
Private m_LastID             As Long
Private m_IsCreated          As Boolean
Private m_MenusToDestroy()   As Long
Private m_MenusToDestroyDx() As Long
Private m_MenusToDestroyCnt  As Long
Private m_FromForm           As Boolean
Private m_AllSupportedKeys   As String
Private m_hWndKeyOwner       As Long
Private m_ConsumeKeys        As Boolean
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Events
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Public Event MenuHighlight(ByVal Index As Long)   ', ByVal IsSeparator As Boolean)
Public Event MenuExit()
Public Event DrawItem(ByRef Cancel As Boolean, ByVal Index As Long, ByVal hDC As Long, ByVal bSelected As Boolean, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Public Event MeasureItem(ByRef Cancel As Boolean, ByVal Index As Long, ByRef rWidth As Long, ByRef rHeight As Long)
Public Event Click(ByVal Index As Long)
Private Function pvInitMenuPopUp(ByVal hMenu As Long)
    'Call pvEnumMenus(hMenu)
End Function

'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Class Events
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub Class_Initialize()
    Set m_SubClass = New cSuperClass
    Set m_VBMenus = New Collection
    Set m_CollMenus = New Collection
    Set m_MemDC = New cMemDC
    Set m_TmpMemDC = New cMemDC
    '----
    m_BackColor = vbMenuBar
    m_ForeColor = vbMenuText
    m_HighlightColor = vbHighlight
    m_HighlightForeColor = vbWhite
    With m_TmpMemDC
        Call .Init(100, 100)
        Set m_Font = .SystemMenuFont
        Set .Font = m_Font
        'm_ItemHeight = .FontHeight + 6
        Call pvCalculateItemHeight
    End With
    'hook keyboard
    Call HookKeyboard(Me)
    m_ConsumeKeys = True
End Sub
Private Sub Class_Terminate()
    Dim I As Long
    For I = 1 To m_CollMenus.Count
        Call DestroyMenu(m_CollMenus.Item(I).hMenu)
    Next
    Set m_SubClass = Nothing
    Set m_VBMenus = Nothing
    Set m_CollMenus = Nothing
    Set m_MemDC = Nothing
    Set m_TmpMemDC = Nothing
    Set m_ImageList = Nothing
    Set m_Font = Nothing
    Set m_BackGroundPicture = Nothing
    m_AllSupportedKeys = vbNullString
    Call RemoveHookKeyboard(Me)
End Sub
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Properties
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'-- Move pointer to a menu so that we can
'call properties without telling the index
Public Property Get CurrentMenuIndex() As Long
    CurrentMenuIndex = m_CurrentMenuIndex
End Property
Public Property Let CurrentMenuIndex(ByVal vNewValue As Long)
    m_CurrentMenuIndex = vNewValue
End Property
'-- Source of imaged
Public Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property
Public Property Set ImageList(ByVal vNewValue As Object)
    Set m_ImageList = vNewValue
End Property
'-- Set\Clear an image for a menu
Public Property Get ItemImage(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemImage = ItemClass.Image
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemImage(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Image = vNewValue
        Set ItemClass = Nothing
    End If
End Property
'-- Set\Get the help text for a menu
Public Property Get ItemHelpText(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemHelpText = ItemClass.HelpText
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemHelpText(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.HelpText = vNewValue
        Set ItemClass = Nothing
    End If
End Property
'-- Get\Set caption of an item
Public Property Get ItemCaption(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemCaption = ItemClass.Caption
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemCaption(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.Caption = vNewValue
        Call pvSetCaption(cIndex, vNewValue)
        'remove item - insert again
        Call ReplaceItem(cIndex)
        Set ItemClass = Nothing
    End If
End Property
'-- Get\Set if the item is a menu-break
Public Property Get ItemBreak(Optional Index) As Boolean
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        ItemBreak = ((MII.fType And MF_MENUBREAK) = MF_MENUBREAK)
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemBreak(Optional Index, ByVal vNewValue As Boolean)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Break = vNewValue
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        If (vNewValue) Then
            MII.fType = MII.fType Or MF_MENUBREAK
        Else
            MII.fType = MII.fType And Not MF_MENUBREAK
        End If
        Call SetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        Set ItemClass = Nothing
    End If
End Property
'-- Get\Set if the item is a menu-break
Public Property Get ItemSeparator(Optional Index) As Boolean
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        ItemSeparator = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemSeparator(Optional Index, ByVal vNewValue As Boolean)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        If (ItemClass.Separator <> vNewValue) Then
            ItemClass.Separator = vNewValue
            'apply
            Dim MII As MENUITEMINFO
            MII.cbSize = Len(MII)
            MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            MII.fType = 0
            MII.dwTypeData = sBuffer
            MII.cch = 80
            Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
            If (vNewValue) Then
                MII.fType = MII.fType Or MF_SEPARATOR
            Else
                MII.fType = MII.fType And Not MF_SEPARATOR
            End If
            Call SetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
            Call ReplaceItem(cIndex)
        End If
        Set ItemClass = Nothing
    End If
End Property
Public Property Get ItemEnabled(Optional Index) As Boolean
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        ItemEnabled = Not ((MII.fState And MFS_DISABLED) = MFS_DISABLED)
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemEnabled(Optional Index, ByVal vNewValue As Boolean)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.X = vNewValue
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        If (vNewValue) Then
            MII.fState = MII.fState And Not MFS_DISABLED
        Else
            MII.fState = MII.fState Or MFS_DISABLED
        End If
        'Debug.Print MII.cch
        'Dim Temp As String
        'MsgBox InStr(1, MII.dwTypeData, vbNullChar)
        'Temp = Left(Replace(MII.dwTypeData, Chr(0), ""), MII.cch)
        'Temp = "success" & vbNullChar & "Ctrl+T" & vbNullChar
        'MII.dwTypeData = Temp & String(80 - Len(Temp), 0)
        'sBuffer = "New" & vbNullChar
        Call SetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        If (ItemClass.TopMenu) Then
            Call DrawMenuBar(m_hWnd)
        End If
        Set ItemClass = Nothing
    End If
End Property
Public Property Get ItemChecked(Optional Index) As Boolean
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        ItemChecked = ((MII.fState And MFS_CHECKED) = MFS_CHECKED)
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemChecked(Optional Index, ByVal vNewValue As Boolean)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    Dim sBuffer   As String * 80
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.X = vNewValue
        'apply
        Dim MII As MENUITEMINFO
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        If (vNewValue) Then
            MII.fState = MII.fState Or MFS_CHECKED
        Else
            MII.fState = MII.fState And Not MFS_CHECKED
        End If
        Call SetMenuItemInfo(ItemClass.hMenu, ItemClass.Index, True, MII)
        Set ItemClass = Nothing
    End If
End Property
'-- Just to store the HELP-TEXT for an item
Public Property Get HelpText(Index) As String
    On Error Resume Next
    Dim ItemClass As clsStoreMenu
    Set ItemClass = m_CollMenus.Item(Index)
    If (ItemClass Is Nothing) Then
        Exit Property
    End If
    HelpText = ItemClass.HelpText
End Property
Public Property Let HelpText(Index, ByVal vNewValue As String)
    On Error Resume Next
    Dim ItemClass As clsStoreMenu
    Set ItemClass = m_CollMenus.Item(Index)
    If (ItemClass Is Nothing) Then
        Exit Property
    End If
    ItemClass.HelpText = vNewValue
End Property
'-- Get\Set if the item is ItemOwnerDraw by you
'if true the event MeasureItem & DrawItem
'will be called before processing
Public Property Get ItemOwnerDraw(Optional Index) As Boolean
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDraw = ItemClass.OwnerDraw
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemOwnerDraw(Optional Index, ByVal vNewValue As Boolean)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        If (ItemClass.OwnerDraw <> vNewValue) Then
            ItemClass.OwnerDraw = vNewValue
            'remove item - insert again
            Call ReplaceItem(cIndex)
        End If
        Set ItemClass = Nothing
    End If
End Property
'-- Just to store a number which descripe the
'way you want to draw your item
Public Property Get ItemOwnerDrawID(Optional Index) As Long
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDrawID = ItemClass.OwnerDrawID
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemOwnerDrawID(Optional Index, ByVal vNewValue As Long)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.OwnerDrawID = vNewValue
        Set ItemClass = Nothing
    End If
End Property
'-- Just to store the number value which will
'help you to draw your item
Public Property Get ItemOwnerDrawData(Optional Index) As Long
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDrawData = ItemClass.OwnerDrawData
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemOwnerDrawData(Optional Index, ByVal vNewValue As Long)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.OwnerDrawData = vNewValue
        Set ItemClass = Nothing
    End If
End Property
Public Property Get ItemDescription(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemDescription = ItemClass.Description
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemDescription(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Description = vNewValue
        Set ItemClass = Nothing
    End If
End Property
Public Property Get ItemTag(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemTag = ItemClass.Tag
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemTag(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Tag = vNewValue
        Set ItemClass = Nothing
    End If
End Property
'-- Get\Set the shortcut of a menu
'this is a great property not included in
'others classes , ENJIY ;)
Public Property Get ItemKeyAccel(Optional Index) As String
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemKeyAccel = ItemClass.KeyAccel
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemKeyAccel(Optional Index, ByVal vNewValue As String)
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.KeyAccel = vNewValue
        If (Not (ItemClass.TopMenu)) Then
            If (StrComp(ItemClass.KeyAccel, vNewValue) <> 0) Then
                If pvSetKeyAccel(cIndex, vNewValue) Then
                    'remove item - insert again
                    Call ReplaceItem(cIndex)
                End If
            End If
        End If
        Set ItemClass = Nothing
    End If
End Property
Private Function pvSetKeyAccel(ByVal Index As Long, ByVal vNewValue As String) As Boolean
    If (Not (pvCheckKeys(vNewValue))) Then
        Exit Function
    End If
    With m_CollMenus.Item(Index)
        Call pvRemoveSupportedKeys(.KeyAccel)
        If (vNewValue <> vbNullString) Then
            Call pvAddSupportedKeys(vNewValue)
        End If
        .KeyAccel = vNewValue
    End With
    pvSetKeyAccel = True
End Function
Private Function pvCheckKey(ByRef lStr As String) As Boolean
    pvCheckKey = True
    'Dim keyArr() As String
    'Dim I As Long
    'Dim repStr As String
    'keyArr = Split(lStr, "+")
    'For I = 0 To UBound(keyArr)
    '
    'Next
    'Erase keyArr
End Function
Private Function pvCheckKeys(ByRef lStr As String) As Boolean
    Dim keyArr() As String
    Dim I As Long
    Dim repStr As String
    keyArr = Split(lStr, "|")
    pvCheckKeys = True
    For I = 0 To UBound(keyArr)
        If (pvCheckKey(keyArr(I))) Then
            If (I = 0) Then
                repStr = keyArr(I)
            Else
                repStr = repStr & "|" & keyArr(I)
            End If
        Else
            pvCheckKeys = False
        End If
    Next
    Erase keyArr
End Function
Private Function pvAddSupportedKeys(ByVal lKeys As String)
    If (m_AllSupportedKeys = vbNullString) Then
        m_AllSupportedKeys = "|" & lKeys & "|"
    Else
        m_AllSupportedKeys = Replace(m_AllSupportedKeys, "|", "|" & lKeys & "|", , 1)
    End If
End Function
Private Function pvRemoveSupportedKeys(ByVal lKeys As String)
    m_AllSupportedKeys = Replace(m_AllSupportedKeys, "|" & lKeys & "|", "|")
    If (Len(m_AllSupportedKeys) = 1) Then
        m_AllSupportedKeys = vbNullString
    End If
End Function
Private Function pvKeyAccelExists(ByVal KeyAccel As String) As Boolean
    Dim Fnd As Long
    Fnd = InStr(1, m_AllSupportedKeys, "|" & KeyAccel & "|")
    pvKeyAccelExists = (Fnd <> 0)
End Function
Public Function IndexForKeyAccel(ByVal KeyAccel As String) As Long
    If (Not (pvKeyAccelExists(KeyAccel))) Then
        Exit Function
    End If
    Dim I As Long
    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If (InStr(1, "|" & .KeyAccel & "|", "|" & KeyAccel & "|")) Then
                IndexForKeyAccel = I
                Exit For
            End If
        End With
    Next
End Function
'-- Get\Set the checked style of an item
Public Property Get ItemCheckedStyle(Optional Index) As MenuCheckedStyle
    On Error Resume Next
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemCheckedStyle = ItemClass.CheckedStyle
        Set ItemClass = Nothing
    End If
End Property
Public Property Let ItemCheckedStyle(Optional Index, ByVal vNewValue As MenuCheckedStyle)
    On Error Resume Next
    Dim ItemClass As clsStoreMenu
    Dim cIndex    As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.CheckedStyle = vNewValue
        Set ItemClass = Nothing
    End If
End Property
'-- Get\Set the Style of menus
Public Property Get DrawStyle() As MenuDrawStyle
    DrawStyle = m_DrawStyle
End Property
Public Property Let DrawStyle(ByVal vNewValue As MenuDrawStyle)
    If (m_DrawStyle <> vNewValue) Then
        m_DrawStyle = vNewValue
        Call pvCalculateItemHeight
        Call pvReplaceAllItems
    End If
End Property
'-- Get\Set the Right to Left
Public Property Get RightToLeft() As OLE_COLOR
    RightToLeft = m_RightToLeft
End Property
Public Property Let RightToLeft(ByVal vNewValue As OLE_COLOR)
    m_RightToLeft = vNewValue
End Property
'-- Get\Set
Public Property Get ConsumeKeys() As Boolean
    ConsumeKeys = m_ConsumeKeys
End Property
Public Property Let ConsumeKeys(ByVal vNewValue As Boolean)
    m_ConsumeKeys = vNewValue
End Property
'-- Get\Set the BackColor of menus
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    m_BackColor = vNewValue
End Property
'-- Get\Set the ForeColor of menus
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
End Property
'-- Get\Set the HighlightColor of menus
Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = m_HighlightColor
End Property
Public Property Let HighlightColor(ByVal vNewValue As OLE_COLOR)
    m_HighlightColor = vNewValue
End Property
'-- Get\Set the HighlightForeColor of menus
Public Property Get HighlightForeColor() As OLE_COLOR
    HighlightColor = m_HighlightForeColor
End Property
Public Property Let HighlightForeColor(ByVal vNewValue As OLE_COLOR)
    m_HighlightForeColor = vNewValue
End Property
Public Property Get hWndKeyOwner() As Long
    hWndKeyOwner = m_hWndKeyOwner
End Property
Public Property Let hWndKeyOwner(ByVal vNewValue As Long)
    m_hWndKeyOwner = vNewValue
End Property
'-- Get\Set the Font of menus
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property
Public Property Set Font(ByVal vNewValue As StdFont)
    With m_TmpMemDC
        If (vNewValue Is Nothing) Then
            Set m_Font = .SystemMenuFont
        ElseIf (vNewValue = m_Font) Then
            vNewValue.Bold = False
            Exit Property
        Else
            vNewValue.Bold = False
            Set m_Font = vNewValue
        End If
        Set .Font = m_Font
        Call pvCalculateItemHeight
        Call pvReplaceAllItems
        If (m_FromForm) Then
            Call DrawMenuBar(m_hWnd)
        End If
    End With
End Property
Private Sub pvCalculateItemHeight()
    If (m_DrawStyle = mds_3D) Then
        m_ItemHeight = m_TmpMemDC.FontHeight + 6
    ElseIf (m_DrawStyle = mds_XP) Then
        m_ItemHeight = m_TmpMemDC.FontHeight + 9
    End If
End Sub
'-- Get the menu index of a menu name
Public Property Get ItemOwnerDrawAll() As Boolean
    ItemOwnerDrawAll = m_ItemOwnerDrawAll
End Property
Public Property Let ItemOwnerDrawAll(ByVal vNewValue As Boolean)
    m_ItemOwnerDrawAll = vNewValue
End Property
'-- Get the backgorund of menus
Public Property Get BackGroundPicture() As StdPicture
    Set BackGroundPicture = m_BackGroundPicture
End Property
Public Property Let BackGroundPicture(ByVal vNewValue As StdPicture)
    Set m_BackGroundPicture = vNewValue
    If (m_BackGroundPicture Is Nothing) Then
        m_DrawBackGround = False
    Else
        With m_TmpMemDC
            m_BackGroundWidth = .HM2Pix(m_BackGroundPicture.Width)
            m_BackGroundHeight = .HM2Pix(m_BackGroundPicture.Height)
        End With
        m_DrawBackGround = True
    End If
End Property
'-- Get the menu index of a menu name
Public Function IndexForKey(Key) As Long
    If (IsMissing(Key)) Then
        IndexForKey = m_CurrentMenuIndex
    Else
        If (IsNumeric(Key)) Then
            IndexForKey = CLng(Key)
            If (IndexForKey < 1 Or IndexForKey > m_CollMenus.Count) Then
                IndexForKey = 0
            End If
        Else
            Dim I As Long
            For I = 1 To m_CollMenus.Count
                If StrComp(Key, m_CollMenus.Item(I).Name, vbTextCompare) = 0 Then
                    IndexForKey = I
                End If
            Next
        End If
    End If
End Function
Public Property Get ItemCount() As Long
    ItemCount = m_CollMenus.Count
End Property
Public Property Get ItemKey(ByVal Index As Long) As String
    ItemKey = m_CollMenus.Item(Index).Name
End Property
Public Property Get ItemTopMenu(Optional Index) As Boolean
    Dim cIndex As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        ItemTopMenu = m_CollMenus.Item(cIndex).TopMenu
    End If
End Property
Public Property Get ItemID(Optional Index) As Long
    Dim cIndex As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        ItemID = m_CollMenus.Item(cIndex).ID
    End If
End Property
Public Property Get ItemParentID(Optional Index) As Long
    Dim cIndex As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        ItemParentID = m_CollMenus.Item(cIndex).ParentID
    End If
End Property
Public Property Get ItemParentIndex(Optional Index) As Boolean
    Dim cIndex As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        ItemParentIndex = IndexForID(m_CollMenus.Item(cIndex).ParentID)
    End If
End Property
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Functions
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Public Sub CreateFromForm(oForm As Object, Optional TopMenus As Boolean = True)
    If (m_IsCreated) Then
        Exit Sub
    End If
    Dim CTL As Control
    Dim cIndex As Long
    Dim hMenu As Long
    m_TopMenus = TopMenus
    If Not (TypeOf oForm Is Form) Then
        Exit Sub
    End If
    m_hWnd = oForm.hWnd
    hWndKeyOwner = m_hWnd
    m_FromForm = True
    For Each CTL In oForm.Controls
        If TypeOf CTL Is Menu Then
            Call m_VBMenus.Add(CTL)
        End If
    Next
    m_hMenu = GetMenu(m_hWnd)
    m_hSystemMenu = GetSystemMenu(m_hWnd, 0)
    m_RightToLeft = oForm.RightToLeft
    Call pvEnumMenus(m_hMenu, 0)
    'With m_TmpMemDC
    '    Call .Init(100&, 100&)
        'Set .Font = .SystemMenuFont
        'm_ItemHeight = .FontHeight + 6
    'End With
    'Set m_TmpMemDC.Font = m_Font
    Call pvSubClass
    m_IsCreated = True
End Sub
Public Function CreateFromNothing(Optional hWnd) As Boolean
    If (m_IsCreated) Then
        Exit Function
    End If
    m_hMenu = CreatePopupMenu()
    If (Not (IsMissing(hWnd))) Then
        m_hWnd = CLng(hWnd)
        Call pvSubClass
        'With m_SubClass
        '    Call .AddMsg(WM_MEASUREITEM, True)
        '    Call .AddMsg(WM_DRAWITEM, True)
        '   'Call .AddMsg(WM_ERASEBKGND, True)
        '    'Call .AddMsg(WM_INITMENUPOPUP, True)
        '    Call .SubClass(m_hWnd, Me)
        'End With
        m_IsCreated = True
    End If
End Function
Friend Sub MenuClicked(ByVal MenuName As String)
    RaiseEvent Click(MenuName)
End Sub
'=========================
'==== PRIVATE
'=========================
Private Sub pvEnumMenus(ByVal hMenu As Long, ByVal ParentID As Long, Optional ItemOwnerDraw As Boolean = True)
    On Error Resume Next
    Dim I          As Long
    Dim hSubMenu   As Long
    Dim MenuCount  As Long
    Dim Store_Menu As clsStoreMenu
    Dim Chk        As Long
    Dim cKey       As String
    Dim VBMenu     As Menu
    MenuCount = GetMenuItemCount(hMenu)
    For I = 0 To MenuCount - 1
        Set Store_Menu = pvGetMenu(hMenu, I, ParentID, ItemOwnerDraw)   ' New clsStoreMenu
        With Store_Menu
            'Set .Parent = Me
            m_MenusCount = m_MenusCount + 1
            If (m_FromForm) Then
                Set VBMenu = m_VBMenus(m_MenusCount)
                .Tag = VBMenu.Tag
                cKey = VBMenu.Name
                Chk = VBMenu.Index
            End If
            '.ID = GetMenuItemID(hMenu, I)
            'Call Err.Clear
            '.Caption = .vbMenuID.Caption
            If Err.Number = 0 Then
                cKey = cKey & "(" & Chk & ")"
            End If
            .Name = cKey
            '.hMenu = hMenu
            '.Index = I
            Call m_CollMenus.Add(Store_Menu, "#" & .ID)
            hSubMenu = GetSubMenu(hMenu, I)
            If hSubMenu > 0 Then
                .HasSub = True
                Call pvEnumMenus(hSubMenu, .ID, True)
            End If
        End With
    Next
    Set VBMenu = Nothing
    Set Store_Menu = Nothing
End Sub
Private Function pvGetMenu(ByVal hMenu As Long, ByVal Index As Long, ByVal ParentID As Long, Optional ByVal ItemOwnerDraw As Boolean = True) As clsStoreMenu
    Dim New_Class As clsStoreMenu
    Dim MII As MENUITEMINFO
    Dim MI(0 To 1023) As Byte
    Dim sBuffer As String * 80
    Dim Temp As String
    Dim IsSep As Boolean
    Set New_Class = New clsStoreMenu
    With New_Class
    '********************
        MII.cbSize = Len(MII)
        MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
        MII.fType = 0
        MII.dwTypeData = sBuffer
        MII.cch = 80
        Call GetMenuItemInfo(hMenu, Index, True, MII)
        'mnuID = MII.wID
        .ID = MII.wID
        .ParentID = ParentID
        IsSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
        .Separator = IsSep
        Temp = Left(Replace(MII.dwTypeData, Chr(0), ""), MII.cch)
        If InStr(1, Temp, vbTab) Then
            .Caption = Split(Temp, vbTab)(0)
            .KeyAccel = Split(Temp, vbTab)(1)
        Else
            .Caption = Temp
            .KeyAccel = ""
        End If
        .hSubMenu = MII.hSubMenu
        .hMenu = hMenu
        .Index = Index
        .TopMenu = (hMenu = m_hMenu)
        .CheckedStyle = mcs_Normal
        'If it's not a main menu or a separator then convert to 'ItemOwnerDrawn'
        'If ItemOwnerDraw And IsSep = False Then
        If ItemOwnerDraw Then
            MII.fType = MII.fType Or MF_ItemOwnerDraw
            Call SetMenuItemInfo(hMenu, Index, True, MII)
        End If
    '********************
    End With
    Set pvGetMenu = New_Class
    Set New_Class = Nothing
End Function
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== SubClassing
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub ISuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'hold place
End Sub
Private Sub ISuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    Select Case uMsg
    Case WM_MEASUREITEM
        If (pvMeasureItem(wParam, lParam)) Then
            lHandled = True
        End If
    Case WM_DRAWITEM
        '************************************
        'Debug.Print "DrawItem"
        If (pvDrawItem(wParam, lParam)) Then
            lHandled = True
        End If
    Case WM_ERASEBKGND
        '************************************
        'Debug.Print "TILE"
        lHandled = True
    Case WM_INITMENUPOPUP
        '************************************
        Call pvInitMenuPopUp(wParam)
    Case WM_MENUSELECT
        '************************************
        Call pvMenuSelect(wParam, lParam)
    Case WM_COMMAND
        '************************************
        If (pvCommand(wParam, lParam)) Then
            lHandled = True
        End If
    Case WM_MENUCHAR
        '************************************
        If (((wParam \ &H10000) And MF_SYSMENU) <> MF_SYSMENU) Then
            lReturn = pvMenuChar(lParam, wParam And &HFFFF&)
            lHandled = True
        End If
    End Select
End Sub
Private Function pvMenuChar(ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim I As Long
    Dim cChar As String
    cChar = UCase$(Chr$(lParam))
    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If (.hMenu = wParam) Then
                If (.AccelChar = cChar) Then
                    pvMenuChar = &H20000 Or .Index
                    Exit For
                End If
            End If
        End With
    Next
End Function
Private Function pvCommand(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    On Error Resume Next
    Dim lHiWord    As Long
    Dim lMenuId    As Long
    Dim currMenu   As clsStoreMenu
    If (lParam = 0) Then
        ' Low order word of the wParam item is the menu item id:
        lMenuId = (wParam And &HFFFF&)
        Set currMenu = m_CollMenus.Item("#" & lMenuId)
        RaiseEvent Click(IndexForKey(currMenu.Name))
        Set currMenu = Nothing
        'If (currMenux) Then
        '    ' Don't send on the WM_COMMAND if the item
        '    ' wasn't a VB menu, it might interfere
        '    ' with some other control items!
        '    m_emr = emrConsume
        'Else
        '    ' Otherwise allow the message to parse through
        '    ' to the click event on the VB menu so your old
        '    ' code continues to work:
        '    m_emr = emrPostProcess
        'End If
    End If
End Function
Private Function pvMenuSelect(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    On Error Resume Next
    Dim lHiWord    As Long
    Dim lMenuId    As Long
    Dim bSeparator As Boolean
    Dim bEnabled   As Boolean
    Dim currMenu   As clsStoreMenu
    Dim hMenu      As Long
    lHiWord = wParam \ &H10000
    lMenuId = wParam And &HFFFF&
    
    bSeparator = ((lHiWord And MF_SEPARATOR) = MF_SEPARATOR)
    bEnabled = ((lHiWord And MF_DISABLED) = MF_DISABLED) Or ((lHiWord And MF_GRAYED) = MF_GRAYED)
    ' Menu handle is passed in as lParam:
    hMenu = lParam
    ' Now check if the message is a menu item higlight,
    ' or whether it is indicating exit from the menu:
    Set currMenu = m_CollMenus.Item("#" & lMenuId)
    If (currMenu Is Nothing) Then
        RaiseEvent MenuExit
    Else
        'Debug.Print currMenu.Name
        RaiseEvent MenuHighlight(IndexForKey(currMenu.Name))
        Set currMenu = Nothing
    End If
End Function
Private Function pvMeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '************************************
    'Set menu size
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim currMenu   As clsStoreMenu
    Dim bCancel    As Boolean
    Dim bSeparator As Boolean
    Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo)) 'get the UDT for the menuitem's dimensions
    If MeasureInfo.CtlType <> ODT_MENU Then GoTo Ignore 'bail - not 'ItemOwnerDrawn'
    'If (MeasureInfo.ItemID = 100) Then
    '    MsgBox "Hi"
    'End If
    pvMeasureItem = True
    Set currMenu = m_CollMenus.Item("#" & MeasureInfo.ItemID)  'get menu details from the class
    'If (MeasureInfo.ItemID = 100) Then
    '    Debug.Print "hey"
    'End If
    'Adjust dimensions as neccessary
    If (currMenu.OwnerDraw Or m_ItemOwnerDrawAll) Then
        Dim rWidth  As Long
        Dim rHeight As Long
        rHeight = m_ItemHeight
        RaiseEvent MeasureItem(bCancel, IndexForKey(currMenu.Name), rWidth, rHeight)
        If (bCancel) Then
            MeasureInfo.ItemWidth = rWidth
            MeasureInfo.ItemHeight = rHeight
        Else
            GoTo AutoMesaure
        End If
    Else
AutoMesaure:
        'Debug.Print "MeasureItem"
        'If (currMenu.Separator) Then
        '    MeasureInfo.ItemWidth = 0 ' m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        '    MeasureInfo.ItemHeight = 9
        'Else
        '    MeasureInfo.ItemHeight = m_ItemHeight ' ParForm.TextHeight(BB.Caption) + 6
        '    If (currMenu.KeyAccel = "") Then
        '        MeasureInfo.ItemWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        ''    Else
        '        MeasureInfo.ItemWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent + m_TmpMemDC.TextWidth(currMenu.KeyAccel) + m_RightIndent
        ' ''   End If
        '  '  'If (currMenu.Name = "mnuFilePrint") Then
        '  '  '    MeasureInfo.ItemData = MF_MENUBREAK
        '    'End If
        'End If
        If (m_DrawStyle = mds_3D) Then
            Call pvMeasureItem3D(currMenu, MeasureInfo.ItemWidth, MeasureInfo.ItemHeight)
        ElseIf (m_DrawStyle = mds_XP) Then
            Call pvMeasureItemXP(currMenu, MeasureInfo.ItemWidth, MeasureInfo.ItemHeight)
        End If
    End If
    'Return the UDT with the new values
    Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
Ignore:
    Set currMenu = Nothing
End Function
Private Sub pvMeasureItem3D(ByVal currMenu As clsStoreMenu, ByRef rWidth As Long, ByRef rHeight As Long)
    If (currMenu.TopMenu) Then
        rHeight = m_ItemHeight ' ParForm.TextHeight(BB.Caption) + 6
        rWidth = m_TmpMemDC.TextWidth(currMenu.Caption)   '+ 2
    ElseIf (currMenu.Separator) Then
        rWidth = 0 ' m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        rHeight = 9
    Else
        rHeight = m_ItemHeight ' ParForm.TextHeight(BB.Caption) + 6
        If (currMenu.KeyAccel = "") Then
            rWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        Else
            rWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent + m_TmpMemDC.TextWidth(pvExtractAccel(currMenu.KeyAccel)) + m_RightIndent
        End If
    End If
End Sub
Private Sub pvMeasureItemXP(ByVal currMenu As clsStoreMenu, ByRef rWidth As Long, ByRef rHeight As Long)
    If (currMenu.TopMenu) Then
        rHeight = m_ItemHeight ' ParForm.TextHeight(BB.Caption) + 6
        'Debug.Print m_TmpMemDC.Font.Name
        rWidth = m_TmpMemDC.TextWidth(currMenu.Caption)   '+ 2
    ElseIf (currMenu.Separator) Then
        rWidth = 0 ' m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        rHeight = 9
    Else
        rHeight = m_ItemHeight ' ParForm.TextHeight(BB.Caption) + 6
        If (currMenu.KeyAccel = "") Then
            rWidth = m_IconSizeXP + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndentXP '+ m_IconSize 'm_ExtraSize
        Else
            rWidth = m_IconSizeXP + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndentXP + m_TmpMemDC.TextWidth(pvExtractAccel(currMenu.KeyAccel)) + m_RightIndentXP
        End If
    End If
End Sub
Private Function pvDrawItem(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    Dim DrawInfo As DRAWITEMSTRUCT
    Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo)) 'get the UDT for the menuitem's appearance
    If (DrawInfo.CtlType <> ODT_MENU) Then
        Exit Function
    End If
    pvDrawItem = True
    'Call pvDrawItem3D(DrawInfo)
    If (m_DrawStyle = mds_3D) Then
        Call pvDrawItem3D(DrawInfo)
    ElseIf (m_DrawStyle = mds_XP) Then
        Call pvDrawItemXP(DrawInfo)
    End If
End Function
Private Sub pvDrawItem3D(DrawInfo As DRAWITEMSTRUCT)  'ByVal wParam As Long, ByVal lParam As Long)
    'Dim DrawInfo   As DRAWITEMSTRUCT
    Dim IsSelected As Boolean
    Dim IsDisabled As Boolean
    Dim IsChecked  As Boolean
    'Dim IsClicked  As Boolean
    Dim HasIcon    As Boolean
    Dim HasSub     As Boolean
    Dim currMenu   As clsStoreMenu
    Dim bCancel    As Boolean
    Dim xMemDC     As cMemDC
    Dim cTop       As Long
    Dim cLeft      As Long
    Dim cText      As String
    Dim cAccel     As String
    '************************************
    'Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo)) 'get the UDT for the menuitem's appearance
    '************************************
    If DrawInfo.CtlType <> ODT_MENU Then GoTo Ignore 'bail - not 'ItemOwnerDrawn'
    '************************************
    IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED) 'selected ?
    IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED) 'disabled ?
    IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED) 'checked ?
    Set currMenu = m_CollMenus.Item("#" & DrawInfo.ItemID)  'get menu details from the class
    HasIcon = currMenu.Image <> vbNullString
    HasSub = currMenu.HasSub
    '************************************
    If (currMenu.OwnerDraw Or m_ItemOwnerDrawAll) Then
        RaiseEvent DrawItem(bCancel, IndexForID(currMenu.ID), DrawInfo.hDC, IsSelected, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Right, DrawInfo.rcItem.Bottom)
        If (bCancel) Then GoTo Ignore
    End If
    '************************************
    Set xMemDC = New cMemDC
    cText = currMenu.Caption
    cAccel = pvExtractAccel(currMenu.KeyAccel)
    With xMemDC
        Call .Init(DrawInfo.rcItem.Right - DrawInfo.rcItem.Left, DrawInfo.rcItem.Bottom - DrawInfo.rcItem.Top)
        Set .Font = m_Font ' .SystemMenuFont
        .BackStyle = BS_TRANSPARENT
        Call .FillRect(0, 0, .Width, .Height, m_BackColor)
        '-- if top menu
        If (currMenu.TopMenu) Then
            Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            'cLeft = (.Width - .TextWidth(cText)) / 2
            'cTop = (.Height - .TextHeight(cText)) / 2
            If (IsDisabled) Then
                .ForeColor = vbWhite
                Call .DrawText(currMenu.Caption, 1, 1, .Width + 1, .Height + 1, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
                .ForeColor = vbButtonShadow
                Call .DrawText(currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
            Else
                If ((DrawInfo.itemState And &H1) = &H1) Then
                    'clicked
                    Call .DrawEdge(0, 0, .Width, .Height, BDR_SUNKENOUTER)
                ElseIf ((DrawInfo.itemState And &H40) = &H40) Then
                    'hilghited
                    Call .DrawEdge(0, 0, .Width, .Height, BDR_RAISEDINNER)
                Else
                    'normal
                End If
                .ForeColor = vbBlack
                Call .DrawText(currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
            End If
            GoTo Display
        End If
        '--------------
        If (m_DrawBackGround) Then
            Dim sX As Long
            Dim sY As Long
            'm_BackGroundWidth
            'm_BackGroundHeight
            sX = -(DrawInfo.rcItem.Left Mod m_BackGroundWidth)
            sY = -(DrawInfo.rcItem.Top Mod m_BackGroundHeight)
            'Debug.Print sY
            'sX = 0: sY = 0
            Dim tI  As Long
            Dim tII As Long
            For tI = sX To .Width Step m_BackGroundWidth
                For tII = sY To .Height Step m_BackGroundHeight
                    Call .PaintPicture(m_BackGroundPicture, tI, tII)
                Next
            Next
        End If
        If currMenu.Separator Then
            Call .DrawLine(0, 4, .Width, 4, vbButtonShadow)
            Call .DrawLine(0, 5, .Width, 5, vbWhite)
            GoTo Display
        End If
        '=========== TEXT ==============
        'cLeft = m_IconSize + m_CaptionIndent
        If (m_RightToLeft) Then
            cLeft = .Width - (m_IconSize + m_CaptionIndent)
        Else
            cLeft = m_IconSize + m_CaptionIndent
        End If
        If (IsDisabled) Then
            'draw menu disabled
            .BackColor = m_BackColor ' vbButtonFace
            .ForeColor = vbWhite
            'Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            'caption
            '---
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                'Call .DrawText(cText, .Width - cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
                Call .DrawText(cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
            '---
            .ForeColor = vbButtonShadow
            'cTop = (.Height - .TextHeight(cText)) / 2
            'cLeft = cLeft - 1
            'cTop = cTop - 1
            'Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            If (m_RightToLeft) Then
                'Call .DrawText(cText, .Width - cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
                Call .DrawText(cText, -1, cTop - 1, .Width - m_IconSize - m_CaptionIndent - 1, .Height - 1, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft - 1, cTop - 1, .Width = 1, .Height - 1, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    Call .DrawText(cAccel, 8, -1, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 10, -1, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
        ElseIf (IsSelected) Then
            'draw menu selected
            .BackColor = m_HighlightColor ' vbHighlight
            .ForeColor = m_HighlightForeColor ' vbWhite
            Dim fLeft  As Long
            Dim fRight As Long
            If (HasIcon) Then
                'Call .FillRect(m_IconSize, 0, .Width, .Height, vbHighlight)
                If (m_RightToLeft) Then
                    fLeft = 0
                    fRight = .Width - m_IconSize ' - m_SubTriSize - 1
                Else
                    fLeft = m_IconSize
                    fRight = .Width ' - m_SubTriSize - 1
                End If
            Else
                'Call .FillRect(0, 0, .Width, .Height, vbHighlight)
                fLeft = 0
                fRight = .Width ' - m_SubTriSize - 1
            End If
            'If (HasSub) Then
            '    fRight = .Width - m_SubTriSize - 1
            'Else
            'End If
            'If (m_tighttoleft) Then
            Call .FillRect(fLeft, 0, fRight, .Height, m_HighlightColor) ' vbHighlight)
            'Else
            '    Call .FillRect(fLeft, 0, fRight, .Height, m_HighlightColor) ' vbHighlight)
            'End If
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                'Call .DrawText(cText, .Width - cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
                Call .DrawText(cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
            'End If
            'accel
            'If cAccel <> "" Then
            '    If (m_RightToLeft) Then
            '        Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
        
        Else
            'draw menu normally
            .BackColor = 0 ' m_BackColor ' vbButtonFace
            .ForeColor = m_ForeColor ' vbBlack
            'Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                'Call .DrawText(cText, .Width - cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
                Call .DrawText(cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
        End If
        '=========== ICON\CHECKED ============
        Dim pDisableDrawIcon As Boolean
        If (IsChecked) Then
            If (IsSelected) Then
                If (m_RightToLeft) Then
                    Call .FillRect(.Width - (m_IconSize - 1), 0, .Width - 1, .Height, LighterColour(m_BackColor))
                Else
                    Call .FillRect(1, 0, m_IconSize - 1, .Height, LighterColour(m_BackColor))
                End If
            Else
                If (m_RightToLeft) Then
                    Call .FillRect(.Width - (m_IconSize - 1), 0, .Width - 1, .Height, vbButtonFace)
                Else
                    Call .FillRect(1, 0, m_IconSize - 1, .Height, vbButtonFace)
                End If
                'Call .FillRect(1, 0, m_IconSize - 1, .Height, vbButtonFace)
            End If
            If (m_RightToLeft) Then
                Call .DrawEdge(.Width - (m_IconSize - 1), 0, .Width - 1, .Height, BDR_SUNKENOUTER, BF_RECT)
            Else
                Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_SUNKENOUTER, BF_RECT)
            End If
            'Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_SUNKENOUTER, BF_RECT)
        End If
        If (IsChecked) Then
            Dim cS As MenuCheckedStyle
            Dim bChar As String
            cS = currMenu.CheckedStyle
            If (cS = mcs_Normal) Then
                bChar = "b"
            ElseIf (cS = mcs_Radio) Then
                bChar = "h"
            End If
            If (bChar <> vbNullString) Then
                pDisableDrawIcon = True
                Dim New_Font As StdFont
                Set New_Font = New StdFont
                With New_Font
                    .Name = "Marlett"
                    .Size = 12
                End With
                Set .Font = New_Font
                .ForeColor = vbBlack
                If (m_RightToLeft) Then
                    Call .DrawText(bChar, .Width - (m_IconSize - 1), 0, .Width - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
                Else
                    Call .DrawText(bChar, 1, 0, m_IconSize - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
                End If
                Set New_Font = Nothing
            End If
        End If
        If (HasIcon And (Not (pDisableDrawIcon))) Then
            Dim oPic As StdPicture
            Set oPic = m_ImageList.ListImages.Item(currMenu.Image).Picture
            If Not (oPic Is Nothing) Then
                If (Not IsDisabled) Then
                    If (IsSelected) Then
                        If (Not IsChecked) Then
                            If (m_RightToLeft) Then
                                Call .DrawEdge(.Width - (m_IconSize - 1), 0, .Width - 1, .Height, BDR_RAISEDINNER)
                            Else
                                Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_RAISEDINNER)
                            End If
                            'Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_RAISEDINNER)
                        End If
                    End If
                End If
                If (IsDisabled) Then
                    If (m_RightToLeft) Then
                        Call .PaintDisabledPicture(oPic, .Width - (16 + 3), (.Height - 16) / 2, 16, 16) ' 2)
                    Else
                        Call .PaintDisabledPicture(oPic, 3, (.Height - 16) / 2, 16, 16)  ' 2)
                    End If
                    'Call .PaintDisabledPicture(oPic, 3, (.Height - 16) / 2, 16, 16) ' 2)
                Else
                    If (m_RightToLeft) Then
                        Call .PaintPicture(oPic, .Width - (16 + 3), (.Height - 16) / 2, 16, 16)  ' 2)
                    Else
                        Call .PaintPicture(oPic, 3, (.Height - 16) / 2, 16, 16) ' 2)
                    End If
                    'Call .PaintPicture(oPic, 3, (.Height - 16) / 2, 16, 16) ' 2)
                End If
                Set oPic = Nothing
            End If
        End If
        '=========== EXTRA ==============
        'If (HasSub) Then
        '    If (IsSelected) Then
        '        'Call .DrawEdge(.Width - m_SubTriSize, 0, .Width, .Height, BDR_INNER)
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width, 0, vbButtonShadow)
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width - m_SubTriSize, .Height, vbButtonShadow)
        '    Else
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width - m_SubTriSize, .Height, vbButtonShadow)
        '        Call .DrawLine(.Width - m_SubTriSize + 1, 0, .Width - m_SubTriSize + 1, .Height, vbWhite)
        '    End If
        '    Call .FillRect(.Width - m_SubTriSize, 0, .Width, .Height, vbButtonShadow)
        'End If
        '===============================
Display:
        Call .BitBlt(DrawInfo.hDC, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, .Width, .Height, 0, 0)
        'Debug.Print "Owner Draw"
    End With
Ignore:
    Set xMemDC = Nothing
    Set currMenu = Nothing
End Sub
Private Sub pvDrawItemXP(DrawInfo As DRAWITEMSTRUCT)  'ByVal wParam As Long, ByVal lParam As Long)
    'Dim DrawInfo   As DRAWITEMSTRUCT
    Dim IsSelected As Boolean
    Dim IsDisabled As Boolean
    Dim IsChecked  As Boolean
    Dim HasIcon    As Boolean
    Dim HasSub     As Boolean
    Dim currMenu   As clsStoreMenu
    Dim bCancel    As Boolean
    Dim xMemDC     As cMemDC
    Dim cTop       As Long
    Dim cLeft      As Long
    Dim cText      As String
    Dim cAccel     As String
    '------
    '************************************
    'Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo)) 'get the UDT for the menuitem's appearance
    '************************************
    If DrawInfo.CtlType <> ODT_MENU Then GoTo Ignore 'bail - not 'ItemOwnerDrawn'
    '************************************
    IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED) 'selected ?
    IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED) 'disabled ?
    IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED) 'checked ?
    Set currMenu = m_CollMenus.Item("#" & DrawInfo.ItemID)  'get menu details from the class
    'If (currMenu Is Nothing) Then
    '    Debug.Print "Error"
    'End If
    HasIcon = currMenu.Image <> vbNullString
    HasSub = currMenu.HasSub
    '************************************
    If (currMenu.OwnerDraw Or m_ItemOwnerDrawAll) Then
        RaiseEvent DrawItem(bCancel, IndexForID(currMenu.ID), DrawInfo.hDC, IsSelected, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Right, DrawInfo.rcItem.Bottom)
        If (bCancel) Then GoTo Ignore
    End If
    '************************************
    Set xMemDC = New cMemDC
    cText = currMenu.Caption '.vbMenuID.Caption
    cAccel = pvExtractAccel(currMenu.KeyAccel)
    With xMemDC
        '-- Init
        Call .Init(DrawInfo.rcItem.Right - DrawInfo.rcItem.Left, DrawInfo.rcItem.Bottom - DrawInfo.rcItem.Top)
        .BackStyle = BS_TRANSPARENT
        Set .Font = m_Font
        '-- if top menu draw and good bye
        If (currMenu.TopMenu) Then
            .ForeColor = vbBlack
            If ((DrawInfo.itemState And &H1) = &H1) Then
                'Clicked
                Call .FillRect(0, 0, .Width, .Height, SC_XPMargin)
                Call .DrawLine(0, 0, .Width, 0, vbButtonShadow)
                Call .DrawLine(0, 0, 0, .Height, vbButtonShadow)
                Call .DrawLine(.Width - 1, 0, .Width - 1, .Height, vbButtonShadow)
            ElseIf ((DrawInfo.itemState And &H40) = &H40) Then
                'Highlighted
                Call .Rectangle(0, 0, .Width, .Height, SC_XPHighlight, , SC_XPBorder)
            Else
                'Normal
                Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            End If
            'draw the caption
            Call .DrawText(currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
            'display
            GoTo Display
        End If
        '--------------
        'draw the background
        Call .FillRect(0, 0, .Width, .Height, SC_XPBack)
        'margin
        If (m_RightToLeft) Then
            Call .FillRect(.Width - m_IconSizeXP, 0, .Width, .Height, SC_XPMargin)
        Else
            Call .FillRect(0, 0, m_IconSizeXP, .Height, SC_XPMargin)
        End If
        '-- if is separator draw and good bye
        If currMenu.Separator Then
            If (m_RightToLeft) Then
                Call .DrawLine(.Width - (m_IconSizeXP + 5), 4, 0, 4, vbButtonShadow)
            Else
                Call .DrawLine(m_IconSizeXP + 5, 4, .Width, 4, vbButtonShadow)
            End If
            'Call .DrawLine(0, 5, .Width, 5, vbWhite)
            GoTo Display
        End If
        '=========== TEXT ==============
        If (m_RightToLeft) Then
            cLeft = .Width - (m_IconSizeXP + m_CaptionIndentXP)
        Else
            cLeft = m_IconSizeXP + m_CaptionIndentXP
        End If
        If (IsDisabled) Then
            'draw menu disabled
            '.BackColor = m_BackColor ' vbButtonFace
            .ForeColor = vbButtonShadow
            'Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            'caption
            '---
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                Call .DrawText(cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
            '---
            '.ForeColor = vbButtonShadow
            ''cTop = (.Height - .TextHeight(cText)) / 2
            'cLeft = cLeft - 1
            'cTop = cTop - 1
            'Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            ''accel
            'If cAccel <> "" Then
            '    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9 - 1, -1, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
            'End If
        ElseIf (IsSelected) Then
            'draw menu selected
            '.BackColor = m_HighlightColor ' vbHighlight
            .ForeColor = vbBlack ' vbWhite
            Dim fLeft  As Long
            Dim fRight As Long
            If (HasIcon) Then
                'Call .FillRect(m_IconSize, 0, .Width, .Height, vbHighlight)
                fLeft = m_IconSize
            Else
                'Call .FillRect(0, 0, .Width, .Height, vbHighlight)
                fLeft = 0
            End If
            'If (HasSub) Then
            '    fRight = .Width - m_SubTriSize - 1
            'Else
            fRight = .Width ' - m_SubTriSize - 1
            'End If
            Call .Rectangle(1, 0, .Width - 1, .Height, SC_XPHighlight, , SC_XPBorder)  ' vbHighlight)
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                Call .DrawText(cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
            End If
            'If cAccel <> "" Then
            '    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
            'End If
        Else
            'draw menu normally
            '.BackColor = 0 ' m_BackColor ' vbButtonFace
            .ForeColor = vbBlack ' vbBlack
            'Call .FillRect(0, 0, .Width, .Height, vbButtonFace)
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If (m_RightToLeft) Then
                Call .DrawText(cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP)
            Else
                Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            End If
            'Call .DrawText(cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
            'accel
            If cAccel <> "" Then
                If (m_RightToLeft) Then
                    Call .DrawText(cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                Else
                    Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                End If
                'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
            End If
        End If
        '=========== ICON\CHECKED ============
        Dim pDisableDrawIcon As Boolean
        If (IsChecked) Then
            If (Not IsDisabled) Then
                If (IsSelected) Then
                    If (m_RightToLeft) Then
                        Call .Rectangle(.Width - (m_IconSizeXP - 1), 1, .Width - 2, .Height - 1, SC_XPHighlight, , SC_XPBorder)
                    Else
                        Call .Rectangle(2, 1, m_IconSizeXP - 1, .Height - 1, SC_XPHighlight, , SC_XPBorder)
                    End If
                Else
                    If (m_RightToLeft) Then
                        Call .Rectangle(.Width - (m_IconSizeXP - 1), 1, .Width - 2, .Height - 1, SC_XPFill, , SC_XPBorder)
                    Else
                        Call .Rectangle(2, 1, m_IconSizeXP - 1, .Height - 1, SC_XPFill, , SC_XPBorder)
                    End If
                    'Call .Rectangle(2, 1, m_IconSizeXP - 1, .Height - 1, SC_XPFill, , SC_XPBorder)
                End If
                'Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_SUNKENOUTER, BF_RECT)
                'Call .Rectangle(1, 0, m_IconSize - 1, .Height, SC_XPBack, , SC_XPBorder)
            End If
        End If
        If (IsChecked) Then
            Dim cS As MenuCheckedStyle
            Dim bChar As String
            cS = currMenu.CheckedStyle
            If (cS = mcs_Normal) Then
                bChar = "b"
            ElseIf (cS = mcs_Radio) Then
                bChar = "h"
            End If
            If (bChar <> vbNullString) Then
                pDisableDrawIcon = True
                Dim New_Font As StdFont
                Set New_Font = New StdFont
                With New_Font
                    .Name = "Marlett"
                    .Size = 12
                End With
                Set .Font = New_Font
                If (IsDisabled) Then
                    .ForeColor = vbButtonShadow
                Else
                    .ForeColor = vbBlack
                End If
                If (m_RightToLeft) Then
                    Call .DrawText(bChar, .Width - m_IconSize - 1 - 2, 0, .Width - 2, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
                Else
                    Call .DrawText(bChar, 2, 0, m_IconSize - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)
                End If
                Set New_Font = Nothing
            End If
        End If
        If (HasIcon And (Not (pDisableDrawIcon))) Then
            Dim oPic As StdPicture
            Set oPic = m_ImageList.ListImages.Item(currMenu.Image).Picture
            If Not (oPic Is Nothing) Then
                If (Not IsDisabled) Then
                    If (IsSelected) Then
                        If (Not IsChecked) Then
                            'Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_RAISEDINNER)
                        End If
                    End If
                End If
                If (IsDisabled) Then
                    If (m_RightToLeft) Then
                        Call .PaintDisabledPicture(oPic, .Width - 20, (.Height - 16) / 2, 16, 16) ' 2)
                    Else
                        Call .PaintDisabledPicture(oPic, 4, (.Height - 16) / 2, 16, 16) ' 2)
                    End If
                Else
                    If (m_RightToLeft) Then
                        Call .PaintPicture(oPic, .Width - 20, (.Height - 16) / 2, 16, 16) ' 2)
                        'Dim oMemDC As cMemDC
                        'Set oMemDC = New cMemDC
                        'With oMemDC
                        '    Call .Init(16, 16)
                        '    Call .PaintPicture(oPic)  ', \ .Width - 20, (.Height - 16) / 2, 16, 16) ' 2)
                        '    Call .BitBlt(xMemDC.hDC, 4, (.Height - 16) / 2, 16, 16, 0, 0)
                        'End With
                        'Debug.Print "Special"
                        'Set oMemDC = Nothing
                    Else
                        'Dim oMemDC As cMemDC
                        If (IsSelected) Then
                            Dim oPic2 As StdPicture
                            Set oPic2 = pvCreateShadowPicture(oPic)
                            '    Set oMemDC = New cMemDC
                            ''    With oMemDC
                            '        Call .Init(16, 16)
                            '        Call .FillRect(0, 0, .Width, .Height, vbButtonShadow)
                            '        Call .PaintPicture(oPic, 0, 0, 16, 16, 0, 0, vbDstInvert) ', vbSrcInvert) ', vbButtonShadow) ', \ .Width - 20, (.Height - 16) / 2, 16, 16) ' 2)
                            '        'Call .BitBlt(xMemDC.hDC, 4, (xMemDC.Height - 16) / 2, 16, 16)
                            '    End With
                            '    Call .PaintPicture(oMemDC.Image, 4, (.Height - 16) / 2, 16, 16, , , , vbBlack)  ' 2)
                            '    Debug.Print "Special"
                            '    Set oMemDC = Nothing
                            Call .PaintPicture(oPic2, 4, (.Height - 16) / 2 + 2, 16, 16, , , , vbBlack) ' 2)
                            Set oPic2 = Nothing
                            Call .PaintPicture(oPic, 4 - 2, (.Height - 16) / 2, 16, 16)  ' 2)
                        Else
                            Call .PaintPicture(oPic, 4, (.Height - 16) / 2, 16, 16) ' 2)
                        End If
                        'End If
                    End If
                    'Call .PaintPicture(oPic, 4, (.Height - 16) / 2, 16, 16) ' 2)
                End If
                Set oPic = Nothing
            End If
        End If
        '=========== EXTRA ==============
        'If (HasSub) Then
        '    If (IsSelected) Then
        '        'Call .DrawEdge(.Width - m_SubTriSize, 0, .Width, .Height, BDR_INNER)
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width, 0, vbButtonShadow)
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width - m_SubTriSize, .Height, vbButtonShadow)
        '    Else
        '        Call .DrawLine(.Width - m_SubTriSize, 0, .Width - m_SubTriSize, .Height, vbButtonShadow)
        '        Call .DrawLine(.Width - m_SubTriSize + 1, 0, .Width - m_SubTriSize + 1, .Height, vbWhite)
        '    End If
        '    Call .FillRect(.Width - m_SubTriSize, 0, .Width, .Height, vbButtonShadow)
        'End If
        '===============================
Display:
        Call .BitBlt(DrawInfo.hDC, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, .Width, .Height, 0, 0)
        'Debug.Print "Owner Draw"
    End With
Ignore:
    Set xMemDC = Nothing
    Set currMenu = Nothing
End Sub
Public Function IndexForID(ByVal ID As Long) As Long
    Dim I As Long
    For I = 1 To m_CollMenus.Count
        If (m_CollMenus.Item(I).ID) = ID Then
            IndexForID = I
            Exit For
        End If
    Next
End Function
Private Function pvCreateShadowPicture(ByVal picThis As StdPicture) As StdPicture
    Dim xMemDC As cMemDC
    Set xMemDC = New cMemDC
    With xMemDC
        Call .Init(.HM2Pix(picThis.Width), .HM2Pix(picThis.Height))
        Call .PaintPicture(picThis)
        Dim sX As Long
        Dim sY As Long
        For sX = 0 To .Width
            For sY = 0 To .Height
                If (.GetPixel(sX, sY)) Then
                    Call .SetPixel(sX, sY, vbButtonShadow)
                End If
            Next
        Next
        Set pvCreateShadowPicture = .Image
    End With
    Set xMemDC = Nothing
End Function
Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
lCFrom = m_TmpMemDC.TranslateColor(oColorFrom)
lCTo = m_TmpMemDC.TranslateColor(oColorTo)

Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
      
End Property
Private Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long
Dim lC As Long
Dim h As Single, S As Single, l As Single
Dim lR As Long, lG As Long, lB As Long
Static s_lColLast As Long
Static s_lLightColLast As Long
   
   lC = m_TmpMemDC.TranslateColor(oColor)
   If (lC <> s_lColLast) Then
      s_lColLast = lC
      RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, S, l
      If (l > 0.99) Then
         l = l * 0.8
      Else
         l = l * 1.2
         If (l > 1) Then
            l = 1
         End If
      End If
      HLSToRGB h, S, l, lR, lG, lB
      s_lLightColLast = RGB(lR, lG, lB)
   End If
   LighterColour = s_lLightColLast
End Property
Private Sub HLSToRGB( _
    ByVal h As Single, ByVal S As Single, ByVal l As Single, _
    R As Long, G As Long, B As Long _
    )
 Dim rR As Single, rG As Single, rB As Single
 Dim Min As Single, Max As Single

     If S = 0 Then
     ' Achromatic case:
     rR = l: rG = l: rB = l
     Else
     ' Chromatic case:
     ' delta = Max-Min
     If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - S)
     Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - S * (1 - l)
     End If
     ' Get the Max value:
     Max = 2 * l - Min
     
     ' Now depending on sector we can evaluate the h,l,s:
     If (h < 1) Then
         rR = Max
         If (h < 0) Then
             rG = Min
             rB = rG - h * (Max - Min)
         Else
             rB = Min
             rG = h * (Max - Min) + rB
         End If
     ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
             rB = Min
             rR = rB - (h - 2) * (Max - Min)
         Else
             rR = Min
             rB = (h - 2) * (Max - Min) + rR
         End If
     Else
         rB = Max
         If (h < 4) Then
             rR = Min
             rG = rR - (h - 4) * (Max - Min)
         Else
             rG = Min
             rR = (h - 4) * (Max - Min) + rG
         End If
         
     End If
             
     End If
     R = rR * 255: G = rG * 255: B = rB * 255
 End Sub
Private Sub RGBToHLS( _
     ByVal R As Long, ByVal G As Long, ByVal B As Long, _
     h As Single, S As Single, l As Single _
     )
 Dim Max As Single
 Dim Min As Single
 Dim delta As Single
 Dim rR As Single, rG As Single, rB As Single

     rR = R / 255: rG = G / 255: rB = B / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = Maximum(rR, rG, rB)
         Min = Minimum(rR, rG, rB)
             l = (Max + Min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = Min Then
             'begin {Acrhomatic case}
             S = 0
             h = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             If l <= 0.5 Then
                 S = (Max - Min) / (Max + Min)
             Else
                 S = (Max - Min) / (2 - Max - Min)
             End If
             '{Next calculate the hue.}
             delta = Max - Min
             If rR = Max Then
                     h = (rG - rB) / delta '{Resulting color is between yellow and magenta}
             ElseIf rG = Max Then
                 h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
 End Sub
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
If (rR > rG) Then
    If (rR > rB) Then
        Maximum = rR
    Else
        Maximum = rB
    End If
Else
    If (rB > rG) Then
        Maximum = rB
    Else
        Maximum = rG
    End If
End If
End Function
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
If (rR < rG) Then
    If (rR < rB) Then
        Minimum = rR
    Else
        Minimum = rB
    End If
Else
    If (rB < rG) Then
        Minimum = rB
    Else
        Minimum = rG
    End If
End If
End Function
Public Function ReplaceItem(Index) As Boolean
    Dim cIndex As Long
    Dim currMenu As clsStoreMenu
    Dim I As Long
    Dim tMI As MENUITEMINFO
    Dim hMenu As Long
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set currMenu = m_CollMenus.Item(cIndex)
    Else
        Exit Function
    End If
    tMI.cbSize = Len(tMI)
    tMI.fMask = MIIM_SUBMENU Or MIIM_TYPE Or MIIM_DATA Or MIIM_ID Or MIIM_CHECKMARKS Or MIIM_STATE ' MIIM_SUBMENU
    Call GetMenuItemInfo(currMenu.hMenu, currMenu.Index, True, tMI)
    Call RemoveMenu(currMenu.hMenu, currMenu.ID, MF_BYCOMMAND)
    Call InsertMenuItem(currMenu.hMenu, currMenu.Index, True, tMI)
    Set currMenu = Nothing
End Function
Public Function RemoveItem(Index) As Boolean
    Dim cIndex As Long
    Dim currMenu As clsStoreMenu
    Dim I As Long
    Dim B As Long
    Dim hMenu As Long
    Dim bTopMenu As Boolean
    cIndex = IndexForKey(Index)
    If (cIndex > 0) Then
        Set currMenu = m_CollMenus.Item(cIndex)
    Else
        Exit Function
    End If
    With currMenu
        Call RemoveMenu(.hMenu, .ID, MF_BYCOMMAND)
        .DestroyMe = True
        bTopMenu = .TopMenu
        Call pvRemoveSubMenus(.ID)
        Call pvMoveIndex(.hMenu, .Index, -1)
        'Call DestroyMenu(.hMenu)
        'hMenu = .hMenu
        'Call m_CollMenus.Remove(cIndex)
    End With
    Set currMenu = Nothing
    Do
        I = I + 1
        If (I > m_CollMenus.Count) Then
            Exit Do
        End If
        With m_CollMenus.Item(I)
            If (.DestroyMe) Then
                Call DestroyMenu(.hMenu)
                'Call DestroyMenu(.hSubMenu)
                'MsgBox .Name
                Call m_CollMenus.Remove(I)
                I = I - 1
            End If
            'B = B + 1
        End With
    Loop
    If (bTopMenu) Then
        Call DrawMenuBar(m_hWnd)
    End If
    'Erase m_MenusToDestroy
    'Erase m_MenusToDestroyDx
    'm_MenusToDestroyCnt = 0
End Function
Private Function pvRemoveSubMenus(ByVal ID As Long)
    'm_MenusToDestroyCnt
    Dim I As Long
    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If (.ParentID = ID) Then
                'm_MenusToDestroyCnt = m_MenusToDestroyCnt + 1
                'ReDim Preserve m_MenusToDestroy(1 To m_MenusToDestroyCnt) As Long
                'ReDim Preserve m_MenusToDestroyDx(1 To m_MenusToDestroyCnt) As Long
                'm_MenusToDestroy(m_MenusToDestroyCnt) = .hMenu
                'm_MenusToDestroyDx(m_MenusToDestroyCnt) = I
                .DestroyMe = True
                Call pvRemoveSubMenus(.ID)
            End If
        End With
    Next
End Function
Public Function AddItem(ByVal Parent As Long, Optional ByVal Caption As String, Optional KeyAccel, Optional Before, Optional Key, Optional Enabled As Boolean = True, Optional Checked As Boolean, Optional CheckedStyle As MenuCheckedStyle, Optional Image, Optional HelpText, Optional OwnerDraw As Boolean, Optional OwnerDrawID, Optional OwnerDrawData, Optional Break As Boolean, Optional Description) As Long
    'add
    On Error Resume Next
    Dim tMParent  As MENUITEMINFO
    Dim tMNew As MENUITEMINFO
    Dim cIndex As Long
    Dim currMenu As clsStoreMenu
    Dim NewMenu  As clsStoreMenu
    Dim parentMenu  As clsStoreMenu
    Dim lR As Long
    Dim hMenu As Long
    Dim bTopMenu As Boolean
    Dim NewIndex As Long
    Dim IsSep   As Boolean
    IsSep = (Trim(Caption) = "-")
    If (Parent = 0) Then
        'top menu
        hMenu = m_hMenu
        If (m_FromForm) Then
            bTopMenu = True
        End If
    Else
        Set currMenu = m_CollMenus.Item(Parent)
        If (currMenu Is Nothing) Then
            Exit Function
        End If
        tMParent.fMask = MIIM_SUBMENU
        tMParent.cbSize = LenB(tMParent)
        Call GetMenuItemInfo(currMenu.hMenu, currMenu.Index, True, tMParent)
        If (tMParent.hSubMenu = 0) Then
            hMenu = CreatePopupMenu()
            tMParent.hSubMenu = hMenu
            Set parentMenu = m_CollMenus.Item(Parent)
            parentMenu.hSubMenu = hMenu
            Set parentMenu = Nothing
        Else
            hMenu = tMParent.hSubMenu
        End If
        ' Now set the parent item so it has a popup menu:
        tMParent.fMask = MIIM_SUBMENU
        tMParent.cbSize = LenB(tMParent)
        'tMParent.hSubMenu = hMenu
        Call SetMenuItemInfo(currMenu.hMenu, currMenu.Index, True, tMParent)
    End If
    If (hMenu <> 0) Then
        With tMNew
            .cbSize = Len(tMNew)
            .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_DATA Or MIIM_STATE
            .wID = pvFreeID
        End With
        tMNew.fType = tMNew.fType Or MF_ItemOwnerDraw
        If (IsSep) Then
            tMNew.fType = tMNew.fType Or MF_SEPARATOR
        End If
        If (Checked) Then
            tMNew.fState = tMNew.fState Or MFS_CHECKED
        End If
        If (Not Enabled) Then
            tMNew.fState = tMNew.fState Or MFS_DISABLED
        End If
        If (Break) Then
            tMNew.fType = tMNew.fType Or MF_MENUBREAK
        End If
        Set NewMenu = New clsStoreMenu
        With NewMenu
            .ID = tMNew.wID
            If (Not bTopMenu) Then
                .ParentID = currMenu.ID
            End If
            .hMenu = hMenu
            .TopMenu = bTopMenu
            '.Caption = Caption
            .Description = Description
            .Separator = IsSep
            If (Not (IsMissing(Key))) Then
                .Name = CStr(Key)
            End If
            .CheckedStyle = CheckedStyle
            NewIndex = -1
            If (Not (IsMissing(Before))) Then
                Dim tstIndex As Long
                tstIndex = IndexForKey(Before)
                If (tstIndex > 0) Then
                    Dim tstMenu  As clsStoreMenu
                    Set tstMenu = m_CollMenus.Item(tstIndex)
                    If (Not (tstMenu Is Nothing)) Then
                        If (tstMenu.hMenu = hMenu) Then
                            Call pvMoveIndex(tstMenu.hMenu, .Index)
                            NewIndex = .Index
                        End If
                    End If
                End If
            End If
            Debug.Print Caption
            If (NewIndex = -1) Then
                Dim iCount As Long
                iCount = GetMenuItemCount(hMenu)
                .Index = iCount '+ 1
            Else
                .Index = NewIndex
            End If
            .OwnerDraw = OwnerDraw
            If (Not (IsMissing(HelpText))) Then
                .HelpText = CStr(HelpText)
            End If
            If (IsMissing(Description)) Then
                .Description = .HelpText
            Else
                .Description = CStr(Description)
            End If
            If (Not (IsMissing(Image))) Then
                .Image = Image
            End If
            If (Not (IsMissing(OwnerDrawID))) Then
                .OwnerDrawID = CLng(OwnerDrawID)
            End If
            If (Not (IsMissing(OwnerDrawData))) Then
                .OwnerDrawData = CLng(OwnerDrawData)
            End If
            Call m_CollMenus.Add(NewMenu, "#" & .ID)
            Call pvSetCaption(m_CollMenus.Count, Caption)
            If (Not (IsMissing(KeyAccel))) Then
                Call pvSetKeyAccel(m_CollMenus.Count, CStr(KeyAccel))
            End If
        End With
        Set NewMenu = Nothing
        If InsertMenuItem(hMenu, NewIndex, True, tMNew) > 0 Then
            AddItem = m_CollMenus.Count
        End If
    End If
    Set currMenu = Nothing
    If (bTopMenu) Then
        Call DrawMenuBar(m_hWnd)
    End If
End Function
Public Function SaveShortcuts(ByVal FileName As String) As Boolean
    Dim hFile As Long
    Dim currMenu As clsStoreMenu
    If (pvIsFileExists(FileName)) Then
        Call Kill(FileName)
    End If
    hFile = FreeFile
    Open FileName For Binary Lock Write As hFile
        Dim I As Long
        Put #hFile, , vbNullChar
        Put #hFile, , "[Boody-Menus-Shortcuts]"
        Put #hFile, , vbNullChar
        Put #hFile, , "[V-1-0-0]"
        Put #hFile, , vbNullChar
        For I = 1 To m_CollMenus.Count
            Set currMenu = m_CollMenus.Item(I)
            With currMenu
                If (.Name = vbNullString) Then
                    Put #hFile, , "I"
                    Put #hFile, , vbNullChar
                    Put #hFile, , CStr(I)
                    Put #hFile, , vbNullChar
                    Put #hFile, , .KeyAccel
                    Put #hFile, , vbNullChar
                Else
                    Put #hFile, , "K"
                    Put #hFile, , vbNullChar
                    Put #hFile, , .Name
                    Put #hFile, , vbNullChar
                    Put #hFile, , .KeyAccel
                    Put #hFile, , vbNullChar
                End If
            End With
        Next
    Close hFile
    Set currMenu = Nothing
    SaveShortcuts = True
End Function
Private Function pvIsFileExists(ByVal FileName As String) As Boolean
    On Local Error GoTo errHandle
    Call FileLen(FileName)
    pvIsFileExists = True
errHandle:
End Function
Public Function LoadShortcuts(ByVal FileName As String) As Boolean
    On Error GoTo errHandle
    Dim hFile As Long
    Dim lStr  As String
    If (Not (pvIsFileExists(FileName))) Then
        Exit Function
    End If
    hFile = FreeFile
    Open FileName For Binary As hFile
        lStr = String(LOF(hFile), 0)
        Get #hFile, , lStr
    Close hFile
    Dim lStrArr() As String
    Dim I As Long
    Dim Cnt  As Long
    lStrArr = Split(lStr, vbNullChar)
    Cnt = UBound(lStrArr)
    If (lStrArr(0) <> vbNullString) Then
        GoTo errHandle
    End If
    If (lStrArr(1) <> "[Boody-Menus-Shortcuts]") Then
        GoTo errHandle
    End If
    If (lStrArr(2) <> "[V-1-0-0]") Then
        GoTo errHandle
    End If
    On Error GoTo IgnoreError
    Dim oldIndex As Long
    Dim currStep As Byte
    Dim useKey   As Boolean
    For I = 3 To Cnt
        If (currStep = 0) Then
            'type
            If (lStrArr(I) = "I") Then
                useKey = False
            ElseIf (lStrArr(I) = "K") Then
                useKey = True
            Else
                GoTo IgnoreError
            End If
            currStep = currStep + 1
        ElseIf (currStep = 1) Then
            'index
            If (useKey) Then
                oldIndex = IndexForKey(lStrArr(I))
            Else
                oldIndex = CLng(lStrArr(I))
            End If
            currStep = currStep + 1
        ElseIf (currStep = 2) Then
            'accel
            If (oldIndex <> 0) Then
                ItemKeyAccel(oldIndex) = CStr(lStrArr(I))
            End If
            currStep = 0
        End If
    Next
IgnoreError:
    LoadShortcuts = True
errHandle:
    lStr = vbNullString
End Function
Public Function XPPopUpMenu(Optional Index, Optional X, Optional Y, Optional ByVal Options As ShowPopupMenuConstants = TPM_LEFTALIGN Or TPM_HORIZONTAL) As Long
    Dim PO As POINTAPI
    Dim CPO As POINTAPI
    Dim CPB As Boolean
    Dim hMenu As Long
    Dim currMenu As clsStoreMenu
    Dim cIndex As Long
    Dim rID    As Long
    Dim tRect  As RECT
    Const TPM_RETURNCMD = &H100
    If (IsMissing(Index)) Then
        hMenu = m_hMenu
    Else
        cIndex = IndexForKey(Index)
        If (cIndex < 1) Then
            Exit Function
        End If
    End If
    If (IsMissing(X)) Then
        Call GetCursorPos(CPO)
        PO.X = CPO.X
        CPB = True
    Else
        PO.X = CLng(X)
    End If
    If (IsMissing(Y)) Then
        If (Not (CPB)) Then
            Call GetCursorPos(CPO)
        End If
        PO.Y = CLng(CPO.Y)
    Else
        PO.Y = CLng(Y)
    End If
    If (hMenu = 0) Then
        Set currMenu = m_CollMenus.Item(cIndex)
        hMenu = currMenu.hSubMenu
    End If
    If (hMenu <> 0) Then
        Options = Options Or TPM_RETURNCMD
        rID = TrackPopupMenu(hMenu, Options, PO.X, PO.Y, 0, m_hWnd, tRect)
    End If
    'MsgBox rID
    Call pvRaiseClick(, rID)
    Set currMenu = Nothing
    XPPopUpMenu = rID
End Function
Private Function pvRaiseClick(Optional Index, Optional ID)
    If (Not (IsMissing(Index))) Then
    
    ElseIf (Not (IsMissing(ID))) Then
        Dim cID As Long
        Dim I As Long
        cID = CLng(ID)
        For I = 1 To m_CollMenus.Count
            With m_CollMenus.Item(I)
                If (.ID = cID) Then
                    'MsgBox .Name
                    RaiseEvent Click(I)
                End If
            End With
        Next
    End If
End Function
Private Function pvMoveIndex(ByVal hMenu As Long, ByVal lStart As Long, Optional lValue As Integer = 1)
    Dim I As Long
    Dim currMenu As clsStoreMenu
    For I = 1 To m_CollMenus.Count
        Set currMenu = m_CollMenus.Item(I)
        With currMenu
            If (hMenu = .hMenu) Then
                If (.Index >= lStart) Then
                    .Index = .Index + lValue
                End If
            End If
        End With
    Next
    Set currMenu = Nothing
End Function
Private Function pvAddNewItem(ByVal currMenu As clsStoreMenu)
    Dim IsTopLevel As Boolean
    Dim hMenu      As Long
    Dim tMII       As MENUITEMINFO
    Dim lR         As Long
    
    If (currMenu.ParentID = 0) Then
        ' This is a new top level menu item:
        hMenu = m_hMenu
        IsTopLevel = True
    Else
        ' We are adding to an existing menu:
        ' First we need to determine if there is already a sub menu for the parent item:
        Dim parentMenu As clsStoreMenu
        Set parentMenu = m_CollMenus.Item("#" & currMenu.ParentID)
        'lPIndex = plGetIndexForId(tMI.lParentId)

        ' Determine if the parent menu has a sub-menu:
        tMII.fMask = MIIM_SUBMENU
        tMII.cbSize = LenB(tMII)
        lR = GetMenuItemInfo(parentMenu.hMenu, parentMenu.ID, False, tMII)
        If lR = 0 Then
            'Debug.Print "Failed"
        End If

        hMenu = tMII.hSubMenu
        If (hMenu = 0) Then
            ' We don't have a sub menu for this item so we're
            ' going to have to add one:
            ' Debug.Print "Adding new sub-menu:"

            ' Create the new menu item and store it's handle so we can clear up
            ' again later:
            hMenu = CreatePopupMenu()
            If (hMenu = 0) Then
                'Debug.Print " *** Failed to create sub menu *** "
            Else
                'm_lSubMenuCount = m_lSubMenuCount + 1
                'ReDim Preserve m_hSubMenus(1 To m_lSubMenuCount) As tSubMenuItem
                'm_hSubMenus(m_lSubMenuCount).hMenu = hMenu
                'm_hSubMenus(m_lSubMenuCount).hSysMenuOwner = m_hLastMDIMenu

                ' Now set the parent item so it has a popup menu:
                tMII.fMask = MIIM_SUBMENU
                tMII.cbSize = LenB(tMII)
                tMII.hSubMenu = hMenu
                lR = SetMenuItemInfo(parentMenu.hMenu, parentMenu.ID, False, tMII)
                If (lR = 0) Then
                    'Debug.Print "Failed to modify menu to add the sub menu " & GetLastError()
                End If

                tMII.fMask = MIIM_ID
                Call GetMenuItemInfo(parentMenu.hMenu, parentMenu.ID, False, tMII)
                'Debug.Assert (tMII.wID = m_tMI(lPIndex).lID)
                'currMenu.lParentId = parentMenu.ID
            End If
        End If
    End If

    'If (hMenu <> 0) Then
    '    With tMIII
    '        .cbSize = Len(tMII)
    '        .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_DATA Or MIIM_STATE
    '        .wID = tMI.lID
    ''        .dwItemData = lIndex
    ' '   End With
    ' '   pSetTypeAndState tMIII, m_iMenuCount
    '    lR = InsertMenuItem(hMenu, -1, True, tMIII)
    '    If (lR = 0) Then
    '        Debug.Print "Failed to add new Menu item"
    '    Else
    '        ' Store the hMenu for this item:
    '        tMI.hMenu = hMenu
    '    End If
    '    If (bTopLevel) Then
    '        DrawMenuBar m_hWndParent
    '    End If
    'End If
End Function
Private Function pvFreeID() As Long
    If (m_LastID = 0) Then m_LastID = 99
    m_LastID = m_LastID + 1
    pvFreeID = m_LastID
End Function
Public Property Get Handle() As Long
    Handle = m_hMenu
End Property
Public Function BuildOn(ByVal hWnd As Long)
    If (m_IsCreated) Then
        Exit Function
    End If
    m_hWnd = hWnd
    'With m_SubClass
    '    Call .AddMsg(WM_MEASUREITEM, True)
    '    Call .AddMsg(WM_DRAWITEM, True)
    '    'Call .AddMsg(WM_ERASEBKGND, True)
    '    'Call .AddMsg(WM_INITMENUPOPUP, True)
    '    Call .SubClass(m_hWnd, Me)
    'End With
    Call pvSubClass
    Call SetMenu(m_hWnd, m_hMenu)
    Call DrawMenuBar(m_hWnd)
    m_IsCreated = True
End Function
Private Function pvExtractAccel(ByVal lStr As String) As String
    If (InStr(1, lStr, "|")) Then
        pvExtractAccel = Split(lStr, "|")(0)
    Else
        pvExtractAccel = lStr
    End If
End Function
Public Function AccelPressed(ByVal lLong As Long) As Boolean

End Function
Private Sub pvReplaceAllItems()
    Dim I As Long
    For I = 1 To m_CollMenus.Count
        Call ReplaceItem(I)
    Next
End Sub
Public Function KeyAccelPressed(ByVal lStr As String) As Boolean
    If (GetActiveWindow() = hWndKeyOwner) Then
        Dim cIndex As Long
        cIndex = IndexForKeyAccel(lStr)
        If (cIndex <> 0) Then
            If (ItemEnabled(cIndex)) Then
                If (Not (ItemSeparator(cIndex))) Then
                    RaiseEvent Click(cIndex)
                    KeyAccelPressed = True
                End If
            End If
        End If
    End If
End Function
Private Sub pvSubClass()
    With m_SubClass
        Call .AddMsg(WM_MEASUREITEM, True)
        Call .AddMsg(WM_DRAWITEM, True)
        Call .AddMsg(WM_ERASEBKGND, True)
        Call .AddMsg(WM_MENUSELECT, True)
        Call .AddMsg(WM_COMMAND, True)
        Call .AddMsg(WM_MENUCHAR, True)
        'Call .AddMsg(WM_INITMENUPOPUP, True)
        Call .SubClass(m_hWnd, Me)
    End With
End Sub
Private Function pvSetCaption(ByVal Index As Long, ByVal vNewValue As String)
    With m_CollMenus.Item(Index)
        .Caption = vNewValue
        Dim Fnd As Long
        Dim lLen As Long
        Dim Chk As String
        Fnd = InStr(1, vNewValue, "&")
        If (Fnd) Then
            lLen = Len(vNewValue)
            If (Fnd <> lLen) Then
                Chk = Mid$(vNewValue, Fnd + 1, 1)
                If (Chk <> "&") Then
                    .AccelChar = UCase$(Chk)
                End If
            End If
        End If
    End With
End Function
