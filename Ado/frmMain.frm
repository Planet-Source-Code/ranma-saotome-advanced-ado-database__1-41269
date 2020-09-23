VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   4800
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox T1 
      Height          =   3255
      Left            =   3840
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   720
      Width           =   4335
      Begin VB.TextBox txtfields 
         DataField       =   "six"
         Height          =   315
         Index           =   4
         Left            =   960
         TabIndex        =   10
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         DataField       =   "five"
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   9
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         DataField       =   "four"
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         DataField       =   "three"
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         DataField       =   "two"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "one"
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Field:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Field:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Field:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Field:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbtools 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   3120
      ScaleHeight     =   3255
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4545
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":019E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0274
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0334
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image imgSplitter 
      Height          =   3255
      Left            =   2760
      MousePointer    =   9  'Size W E
      ToolTipText     =   "Yes, you can reszie me! give it a try"
      Top             =   720
      Width           =   255
   End
   Begin VB.Menu mnumain 
      Caption         =   "&Database"
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnucategory 
         Caption         =   "Add new &category"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnudelcat 
         Caption         =   "Delete a cate&gory"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "&Refresh categorys"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuentry 
         Caption         =   "&Add new entry"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuedit 
         Caption         =   "&Edit existing entry"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "&Update entries"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnucancel 
         Caption         =   "Cancel u&pdate"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete current entry"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufirst 
         Caption         =   "move to first entry"
      End
      Begin VB.Menu mnuprevious 
         Caption         =   "move to previous entry"
      End
      Begin VB.Menu mnunext 
         Caption         =   "move to next entry"
      End
      Begin VB.Menu mnulast 
         Caption         =   "move to last entry"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu mnu23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquick 
         Caption         =   "&Quick search"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu20 
         Caption         =   "-"
      End
      Begin VB.Menu mnustitle 
         Caption         =   "Search by &title"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnutext 
         Caption         =   "Search &by text"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinfos 
         Caption         =   "&Datbase infos"
      End
      Begin VB.Menu mnu22 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim db As Connection
Private mbSplitting As Boolean
Private Const lVSplitLimit As Long = 1500

Private Sub Combo1_Click()
txtfields(0).Text = Combo1.Text
End Sub

Private Sub Form_Load()

On Error GoTo errhandler:

'recievieng some settings from a textfile
With Me
.Left = GetFromINI("settings", "Mainleft", App.Path & "/dtb.ini")
.Top = GetFromINI("settings", "Maintop", App.Path & "/dtb.ini")
.Width = GetFromINI("settings", "Mainwidth", App.Path & "/dtb.ini")
.Height = GetFromINI("settings", "mainheight", App.Path & "/dtb.ini")
.TV.Width = GetFromINI("settings", "Treeview", App.Path & "/dtb.ini")
End With

'lets ADO
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb"



'fill the treeview
Filltree
'some controls
Setcontrols True
'resize form (You can drag the treeview and stuff...)
'sorry i used the MS Tabbed Dialog control, but it was the easiest
Form_Resize

Dim oText As TextBox
  'Bind the text boxes to the data provider
For Each oText In frmMain.txtfields
    Set oText.DataSource = rs
Next
Set Combo1.DataSource = rs

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Resize()
 SizeControls TV.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'writing to textfile if possible
If Me.WindowState <> vbMinimized Then
Call WriteToINI("settings", "Mainleft", Me.Left, App.Path & "/dtb.ini")
Call WriteToINI("settings", "Maintop", Me.Top, App.Path & "/dtb.ini")
Call WriteToINI("settings", "Mainwidth", Me.Width, App.Path & "/dtb.ini")
Call WriteToINI("settings", "Mainheight", Me.Height, App.Path & "/dtb.ini")
Call WriteToINI("settings", "Treeview", TV.Width, App.Path & "/dtb.ini")
End If
Call WriteToINI("settings", "log", Date, App.Path & "/dtb.ini")
'shutdown rs
rs.Close
Set rs = Nothing
End Sub

Private Sub mnucancel_Click()
On Error GoTo errhandler

rs.CancelUpdate
Setcontrols True

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnucategory_Click()
On Error GoTo errhandler:
With FrmCatagory
    .Command2.Caption = "&Add new category"
    .Show vbModal
End With
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnudelcat_Click()
On Error GoTo errhandler:
With FrmCatagory
    .Command2.Caption = "&Delete category"
    .Show vbModal
End With
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub



Private Sub mnudelete_Click()
On Error Resume Next
If txtfields(0).Text = "" Then
Exit Sub
End If
If MsgBox("Delete Entry??", vbCritical + vbYesNo) = vbYes Then
rs.Delete
rs.MoveNext
rs.Update
Filltree
End If
End Sub

Private Sub mnuedit_Click()
On Error GoTo errhandler

If txtfields(0).Text = "" Then Exit Sub
Setcontrols False

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnuentry_Click()
On Error GoTo errhandler
MsgBox "Please choose a category!", vbInformation
Setcontrols False
rs.AddNew

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub mnufirst_Click()
  On Error GoTo GoFirstError

    rs.MoveFirst
    Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub mnuinfos_Click()
frmInfos.Show
End Sub

Private Sub mnulast_Click()
  On Error GoTo GoLastError
    rs.MoveLast
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub mnunext_Click()
  On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rs.MoveLast
  End If

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub mnuprevious_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    rs.MoveFirst
  End If

  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub mnuquick_Click()
'got this from PSC
'thanks
   Dim strFind As String
   Dim intfields As Integer
   Dim txtsearch As String
   
On Error GoTo FindError
  txtsearch = InputBox("Search for an entry", "enter your searchtext", "Treeview")
   If Trim(txtsearch) <> "" Then
     strFind = Trim(txtsearch)
     
     With rs
     .MoveFirst
       Do Until .EOF
         For intfields = 0 To 4
           If InStr(1, frmMain.txtfields(intfields), strFind, _
                    vbTextCompare) > 0 Then
              frmMain.txtfields(intfields).SelStart = _
                      InStr(1, frmMain.txtfields(intfields), _
                            strFind, vbTextCompare) - 1
              frmMain.txtfields(intfields).SelLength = Len(strFind)
            frmMain.txtfields(intfields).SetFocus
              Exit Sub
            End If
          Next
          .MoveNext
          DoEvents
        Loop
        MsgBox "No Match found in Database!", vbExclamation, "search..."
        .MoveFirst
      End With
     End If
     
     Exit Sub
     
FindError:
   
   MsgBox Err.Description
   Err.Clear
End Sub

Private Sub mnuquit_Click()
Unload Me
End Sub

Private Sub mnurefresh_Click()
Filltree
End Sub

Private Sub mnustitle_Click()
On Error Resume Next
'third search thingie
With frmSearch
    .Caption = "search by title"
    .Show
    .Text1.SelStart = 0
    .Text1.SelLength = Len(.Text1.Text)
End With
End Sub

Private Sub mnutext_Click()
On Error Resume Next
'third search thingie
With frmSearch
    .Caption = "search by text"
    .Show
    .Text1.SelStart = 0
    .Text1.SelLength = Len(.Text1.Text)
End With
End Sub

Private Sub mnuupdate_Click()
On Error GoTo errhandler
If Combo1.Text = "" Then Exit Sub

rs.Update
rs.MoveFirst
Setcontrols True
Filltree

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub tbtools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key

'Case "mybutton"

End Select
End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
'our first search routine to jump to the database position
On Error GoTo errhandler
    Select Case Node.Text
        Case "TestTree"
        'we do nothing
            Exit Sub
        Case Node.Key
        'we do nothing
            Exit Sub
        Case Else
        'lets search for the entrie
            Call Search(TV.SelectedItem.Text, rs, rs.Fields("two"))
           Exit Sub
    End Select
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub

Private Sub Filltree()
'Thanks for Redneck Software by Bob Davis (from PSC)
'but was written in DAO - so it was useless hehe and i added the bold treenode - looks nicer
'If you want to delete the pics from the treeview just delete the control, add a new one and put at least 5 pics in it - thats it!

Dim t As Integer
Dim i As Integer
Dim Ba As String
Dim MyNode As Node

     Close #1
     Open App.Path & "\Info.txt" For Output As #1
     Dim f As Integer
     f = 0

     '--- clear FrmMain Combo1 box
     frmMain.Combo1.Clear

     'set Imges the Image List for the Treeview Ctrl
     frmMain.TV.ImageList = frmMain.ImageList1

     '--- Open the DataBase
  Set rs = New Recordset
  rs.Open "select one from table1", db, adOpenStatic, adLockOptimistic
  
     '--- Clear the FrmMain Treeview
     frmMain.TV.Nodes.Clear

     '--- Put in the Main Node in the Treeview
     Set MyNode = frmMain.TV.Nodes.Add(, , "AB", "TestTree", 1)
     '---create a bold treenode!
     BoldTreeNode frmMain.TV.Nodes("AB")
     'Make sure Main Node is Visible
     MyNode.EnsureVisible
       
        '--- Put the Categories into a ListBox so I can Sort them
        '--- and I use it To put the Nodes in the Treeview
        On Error Resume Next
        '--- Move to the First Category
        rs.MoveFirst
           '--- Go through the Categories and add them to the ListBox
           For t = 1 To rs.RecordCount
             '--- add each Category to the Combo1 Box
             frmMain.Combo1.AddItem Trim(rs.Fields(0))
        '--- move to the Next Record
        rs.MoveNext
           Next t
      
    '--- Open the Table where we keep the Recipes
     '---I used unique datafield in the database for field number 2, you can change this if you want, but it helps alot when searching
     '---I also suggest you not allowing empty fields
  Set rs = New Recordset
  rs.Open "select one,two,three,four,five,six from table2 order by one", db, adOpenStatic, adLockOptimistic
                 
Dim oText As TextBox
  'Bind the text boxes to the data provider
For Each oText In frmMain.txtfields
    Set oText.DataSource = rs
Next
    Set Combo1.DataSource = rs
 
       'Go through the ListBox and Get each Category
       For i = 0 To frmMain.Combo1.ListCount - 1
           'Add the Category to the Treeview
           Set MyNode = frmMain.TV.Nodes.Add("AB", tvwChild, Trim(frmMain.Combo1.List(i)), Trim(frmMain.Combo1.List(i)), 2, 3)

           'Use this just in case there are not any records in the Record Set
           'this will happen when You Create a New Blank DataBase
           'Else Your Program will Error Out.
           On Error Resume Next

                'Move through the record set and get each record in a
                'category and put it in the Treeview
                rs.MoveFirst
                     For t = 1 To rs.RecordCount
                        '--- Find each category in the Recipes Table
                         If Trim(frmMain.Combo1.List(i)) = Trim(rs.Fields(0)) Then
                              f = f + 1
                              'Add the Record to the Treeview
                              Set MyNode = frmMain.TV.Nodes.Add(Trim(frmMain.Combo1.List(i)), tvwChild, , Trim(rs.Fields(1)), 4, 5)
                         Else
                         End If
                rs.MoveNext
                     Next t
         
                '--- this is where we put the DataBase Information in the Info.txt file
                Print #1, "Entries in Category " & frmMain.Combo1.List(i) & " = " & f
                             f = 0

                '--- This is supposed to work but It Don't that is
                '--- the reason I used ListBoxes to sort then in
                frmMain.TV.Nodes(frmMain.Combo1.List(i)).Sorted = True
      Next i

     '--- Make Sure all Tree Nodes are visible
     MyNode.EnsureVisible

     '--- Keep the Last tree node Sibling from Expanding
     frmMain.TV.Nodes(1).Child.LastSibling.Expanded = False

     '--- Select the first Node and Highlite it
     frmMain.TV.Nodes(1).Selected = True
     StatusBar1.SimpleText = "Database has currently " & rs.RecordCount & " entries."

     '--- Close the Info.txt file
     '---if you want you can just delete it but then it wont show in the info dialog
     Close #1
     'kill #1 (App.path & "\Info.txt")
End Sub


Private Sub SizeControls(ByVal X As Long)
    On Error Resume Next
'Whats missing here is the resizing of each textbox or the combobox
'but i leave this to you :-P

    Dim lHeightOffSet As Long
    
   lHeightOffSet = 0
   
    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    
   
    With imgSplitter
        .Left = X
        .Width = 150
        .ZOrder
    End With
    
    'scaling Treeview
    With TV
        .Move ScaleLeft, tbtools.Height, X, Me.ScaleHeight - (StatusBar1.Height + tbtools.Height + lHeightOffSet)
    End With
    
    'scaling the rest
    'if you dont have the Tabbed Dialog Control you can experiment with a pic or whatever ;-)
    With T1
        .Move X + 25, TV.Top, Me.ScaleWidth - (TV.Width + 50), TV.Height
    End With
   


    imgSplitter.Top = TV.Top
    imgSplitter.Height = TV.Height

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbSplitting = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    Dim sglPos As Single
    
    If mbSplitting Then
        sglPos = X + imgSplitter.Left
        If sglPos < lVSplitLimit Then
            picSplitter.Left = lVSplitLimit
        ElseIf sglPos > Me.Width - lVSplitLimit Then
            picSplitter.Left = Me.Width - lVSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Handle Splitter Movement - taken straight from the VB Template code
'
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbSplitting = False
End Sub

