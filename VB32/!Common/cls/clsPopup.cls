VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ItemClick(ByVal sItemKey As String)

Private Enum enumMenuItemInfoMembers
    MIIM_STATE = &H1
    MIIM_ID = &H2
    MIIM_SUBMENU = &H4
    MIIM_CHECKMARKS = &H8
    MIIM_TYPE = &H10
    MIIM_DATA = &H20
    MIIM_STRING = &H40
    MIIM_BITMAP = &H80
    MIIM_FTYPE = &H100
End Enum

Private Enum enumMenuFlags
    MF_APPEND = &H100&
    MF_BITMAP = &H4&
    MF_BYCOMMAND = &H0&
    MF_BYPOSITION = &H400&
    MF_CALLBACKS = &H8000000
    MF_CHANGE = &H80&
    MF_CHECKED = &H8&
    MF_CONV = &H40000000
    MF_DEFAULT = &H1000    ' #if(WINVER >= 0x0400)
    MF_DELETE = &H200&
    MF_DISABLED = &H2&
    MF_ENABLED = &H0&
    MF_END = &H80
    MF_ERRORS = &H10000000
    MF_GRAYED = &H1&
    MF_HELP = &H4000&
    MF_HILITE = &H80&
    MF_HSZ_INFO = &H1000000
    MF_INSERT = &H0&
    MF_LINKS = &H20000000
    MF_MASK = &HFF000000
    MF_MENUBARBREAK = &H20&
    MF_MENUBREAK = &H40&
    MF_MOUSESELECT = &H8000&
    MF_OWNERDRAW = &H100&
    MF_POPUP = &H10&
    MF_POSTMSGS = &H4000000
    MF_REMOVE = &H1000&
    MF_RIGHTJUSTIFY = &H4000   ' #if(WINVER >= 0x0400)
    MF_SENDMSGS = &H2000000
    MF_SEPARATOR = &H800&
    MF_STRING = &H0&
    MF_SYSMENU = &H2000&
    MF_UNCHECKED = &H0&
    MF_UNHILITE = &H0&
    MF_USECHECKBITMAPS = &H200&
End Enum

Public Enum enumMenuItemTypes
    MFT_STRING = MF_STRING
    MFT_BITMAP = MF_BITMAP
    MFT_MENUBARBREAK = MF_MENUBARBREAK
    MFT_MENUBREAK = MF_MENUBREAK
    MFT_OWNERDRAW = MF_OWNERDRAW
    MFT_RADIOCHECK = &H200
    MFT_SEPARATOR = MF_SEPARATOR
    MFT_RIGHTORDER = &H2000
    MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
End Enum

Public Enum enumMenuItemStates
    MFS_GRAYED = &H3
    MFS_DISABLED = MFS_GRAYED
    MFS_CHECKED = MF_CHECKED
    MFS_HILITE = MF_HILITE
    MFS_ENABLED = MF_ENABLED
    MFS_UNCHECKED = MF_UNCHECKED
    MFS_UNHILITE = MF_UNHILITE
    MFS_DEFAULT = MF_DEFAULT
    MFS_MASK = &H108B                   ' #if(WINVER >= 0x0500)
    MFS_HOTTRACKDRAWN = &H10000000      ' #if(WINVER >= 0x0500)
    MFS_CACHEDBMP = &H20000000          ' #if(WINVER >= 0x0500)
    MFS_BOTTOMGAPDROP = &H40000000      ' #if(WINVER >= 0x0500)
    MFS_TOPGAPDROP = &H80000000         ' #if(WINVER >= 0x0500)
    MFS_GAPDROP = &HC0000000            ' #if(WINVER >= 0x0500)
End Enum

Private Enum enumTrackPopupMenu
    TPM_CENTERALIGN = &H4
    TPM_LEFTALIGN = &H0
    TPM_RIGHTALIGN = &H8
    TPM_BOTTOMALIGN = &H20
    TPM_TOPALIGN = &H0
    TPM_VCENTERALIGN = &H10
    TPM_NONOTIFY = &H80
    TPM_RETURNCMD = &H100
    TPM_LEFTBUTTON = &H0
    TPM_RIGHTBUTTON = &H2
End Enum

'''Public Enum enumItemTypes
'''    eitChecked = MF_CHECKED
'''    eitAppend = MF_APPEND
'''    eitAlignLeft = TPM_LEFTALIGN
'''    eitAlignRight = TPM_RIGHTALIGN
'''    eitAlignCenter = TPM_CENTERALIGN
'''    eitDisabled = MF_DISABLED
'''    eitGrayed = MF_GRAYED
'''    eitSeparator = MF_SEPARATOR
'''    eitString = MF_STRING
'''End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MENUITEMINFO
    cbSize As Long
    fMask As enumMenuItemInfoMembers
    fType As enumMenuItemTypes
    fState As enumMenuItemStates
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos _
                                Lib "user32" ( _
                            lpPoint As POINTAPI) _
                        As Long

Private Declare Function InsertMenuItem _
                                Lib "user32.dll" _
                                Alias "InsertMenuItemA" ( _
                            ByVal hMenu As Long, _
                            ByVal uItem As Long, _
                            ByVal fByPosition As Long, _
                            lpmii As MENUITEMINFO) _
                        As Long

Private Declare Function CreatePopupMenu _
                                Lib "user32" () _
                        As Long
                        
Private Declare Function TrackPopupMenu _
                                Lib "user32" ( _
                            ByVal hMenu As Long, _
                            ByVal wFlags As enumTrackPopupMenu, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal nReserved As Long, _
                            ByVal hwnd As Long, _
                            lprc As RECT) _
                        As Long
                        
Private Declare Function AppendMenu _
                                Lib "user32" _
                                Alias "AppendMenuA" ( _
                            ByVal hMenu As Long, _
                            ByVal wFlags As Long, _
                            ByVal wIDNewItem As Long, _
                            ByVal lpNewItem As Any) _
                        As Long
                        
Private Declare Function DestroyMenu Lib "user32" ( _
                            ByVal hMenu As Long) _
                        As Long



Dim hMenu As Long

Private colHandles As Collection
Private lCurItem As Long


Public Function AddItem( _
                        ByVal sKey As String, _
                        ByVal sCaption As String, _
                        Optional ByVal eType As enumMenuItemTypes = MFT_STRING, _
                        Optional ByVal eState As enumMenuItemStates = MFS_ENABLED, _
                        Optional ByVal lItemData As Long)
    
    Dim newMenuItem As MENUITEMINFO
    
    lCurItem = lCurItem + 1
    
    With newMenuItem
        ' The size of this structure.
        .cbSize = Len(newMenuItem)
        ' Which elements of the structure to use.
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_DATA
        ' The type of item: a string.
        .fType = eType
        ' This item is currently enabled and is the default item.
        .fState = eState
        ' Set The ItemData
        .dwItemData = lItemData
        ' Assign this item an item identifier.
        .wID = lCurItem
        ' Display the following text for the item.
        .dwTypeData = sCaption
        ' We would set submenu to the handle of an existing
        ' popup to bind them together
        '.hSubMenu = ??
        .cch = Len(.dwTypeData)
    End With
    Dim lHandle As Long
    
    lHandle = InsertMenuItem(hMenu, lCurItem - 1, 1, newMenuItem)
    
    If lHandle <> 0 Then
        colHandles.Add sKey, "h" & lCurItem
    End If
    
End Function


Public Function PopupMenu(ByVal lHwnd As Long)   ', ByVal X As Single, ByVal Y As Single)
    
    Dim Pt As POINTAPI
    
    GetCursorPos Pt
    
    Dim lItem As Long
    Dim rec As RECT
    lItem = TrackPopupMenu(hMenu, TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTALIGN, Pt.x, Pt.y, 0, lHwnd, rec)
    If lItem <> 0 Then
        RaiseEvent ItemClick(colHandles(lItem))
    End If
    
End Function

Private Sub Class_Initialize()
    hMenu = CreatePopupMenu()
    Set colHandles = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set colHandles = Nothing
    DestroyMenu hMenu
End Sub
