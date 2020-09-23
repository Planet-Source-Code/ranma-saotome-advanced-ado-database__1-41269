Attribute VB_Name = "modDatabase"
Option Explicit

Public Sub Setcontrols(bval As Boolean)
'this enabled / disables the buttons and menus we have
'if you plan to use the toolbar you have to include it here
 Dim intFields As Integer
On Error Resume Next
With frmMain
    .mnucategory.Enabled = bval
    .mnudelcat.Enabled = bval
    .mnurefresh.Enabled = bval
    .mnuentry.Enabled = bval
    .mnuedit.Enabled = bval
    .mnuupdate.Enabled = Not bval
    .mnudelete.Enabled = bval
    .mnucancel.Enabled = Not bval
    .mnufirst.Enabled = bval
    .mnuprevious.Enabled = bval
    .mnunext.Enabled = bval
    .mnulast.Enabled = bval
    .mnuquit.Enabled = bval
    .TV.Enabled = bval
    .mnusearch.Visible = bval
    'toolbar usage example
    '.tbtools.Buttons("mybuttonkey").Visible = False
For intFields = 0 To 4
    .txtfields(intFields).Locked = bval
Next
End With
End Sub

Sub Main()
'dont run twice
If App.PrevInstance = True Then Exit Sub
'software protection
If Not App.CompanyName = "Sumari Arts" Then
MsgBox "File Corrupted.", vbCritical, App.Title
End
Exit Sub
End If
'hide application works under NT4.0 Win2k and XP (not tested under 9x)
App.TaskVisible = False
frmMain.Show
End Sub


Public Sub File2ListBox(sFile As String, oList As ListBox)
' sFile = Dir("C:\Random\*.txt")
'This would be the Sub to load the Info.txt File from the
'app.path into a listbox on a form
'got this from a VB forum
'Call File2Listbox(sfile,myForm.List1)
    Dim fnum As Integer
    Dim sTemp As String
    fnum = FreeFile()
    sFile = App.Path & "\Info.txt"
    
    oList.Clear
    Open sFile For Input As fnum


    While Not EOF(fnum)
        Line Input #fnum, sTemp
        oList.AddItem sTemp
    Wend
    Close fnum
End Sub
