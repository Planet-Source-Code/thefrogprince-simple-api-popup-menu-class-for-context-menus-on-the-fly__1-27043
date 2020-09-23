VERSION 5.00
Begin VB.Form frmPopupTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPopupTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents cPopup1 As clsPopup
Attribute cPopup1.VB_VarHelpID = -1
Dim WithEvents cPopup2 As clsPopup
Attribute cPopup2.VB_VarHelpID = -1



Private Sub cPopup1_ItemClick(ByVal sItemKey As String)
    MsgBox "You clicked menu item: " & sItemKey
End Sub

Private Sub Form_Load()
    Set cPopup1 = New clsPopup
    Set cPopup2 = New clsPopup
    
    cPopup1.AddItem "Item1", "This is the &first item"
    cPopup1.AddItem "Divider1", "", MFT_SEPARATOR
    cPopup1.AddItem "Item2", "This is the &second item", MFT_STRING, MFS_CHECKED
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        cPopup1.PopupMenu Me.hwnd   ', X, Y
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cPopup1 = Nothing
    Set cPopup2 = Nothing
    
End Sub
