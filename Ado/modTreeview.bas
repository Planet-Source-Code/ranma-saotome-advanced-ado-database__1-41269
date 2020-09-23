Attribute VB_Name = "modTreeview"
Option Explicit
'this module is for an advanced treeview
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
'
' Treeview Messages and styles
'
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
Public Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SELECTITEM As Long = (TV_FIRST + 11)
'
Public Const TVIF_STATE As Long = &H8
Public Const TVS_TRACKSELECT As Long = &H200&
Public Const TVS_FULLROWSELECT As Long = &H1000
Public Const TVIS_BOLD As Long = &H10
'
Public Const TVGN_ROOT As Long = &H0
Public Const TVGN_NEXT As Long = &H1
Public Const TVGN_CARET As Long = &H9
Public Const EM_LIMITTEXT = &HC5
Public Const WM_VSCROLL = &H115

'
' Treeview Item Structure
'
Public Type TVITEM
   mask As Long
   hItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Sub BoldTreeNode(nNode As Node)
'
' Make a tree node bold
'
' Many thanks to VBNet for this code
'

On Error GoTo vbErrorHandler

    Dim TVI As TVITEM
    Dim lRet As Long
    Dim hItemTV As Long
    Dim lhWnd As Long
    
    Set frmMain.TV.SelectedItem = nNode
    
    lhWnd = frmMain.TV.hwnd
    hItemTV = SendMessageLong(lhWnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    
    If hItemTV > 0 Then
        With TVI
            .hItem = hItemTV
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            lRet = SendMessageAny(lhWnd, TVM_GETITEM, 0&, TVI)
            .State = TVIS_BOLD
        End With
        lRet = SendMessageAny(lhWnd, TVM_SETITEM, 0&, TVI)
    End If
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source, , "frmCodeLib::BoldTreeNode"

End Sub




