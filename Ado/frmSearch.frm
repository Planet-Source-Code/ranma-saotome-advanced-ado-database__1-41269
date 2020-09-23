VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search "
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "frmSearch.frx":000C
      Left            =   120
      List            =   "frmSearch.frx":000E
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "frmSearch.frx":0010
      Left            =   120
      List            =   "frmSearch.frx":0012
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Text"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   2
      Text            =   " "
      Top             =   7200
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Click on any Item in  Box below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Input Search Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- this form has two list boexes one on top of the Other
'--- one has the Category Fields in It and The Other hase The Name Fields
'--- they both scrool together so we can get the Category and
'--- the Name Field at the Same Click
'---

Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim db As Connection

Private Sub cmdParse_Click()
'--- This is where we do the Search for the Name Field
'---

  Dim i As Integer
  Dim t As Integer
  Dim Ss As String
  Dim Ace As String
  
On Error Resume Next

  If Text1 <> "" Then
   '--- String to search for
   Ss = Text1
  Else
    MsgBox "Please enter something! ", vbOKOnly
    Text1.SetFocus
    Exit Sub
  End If
  
  
  List1.Clear
  List2.Clear
  
  If Me.Caption = "search by title" Then
   
      With rs
          .MoveFirst
              For t = 1 To rs.RecordCount
                 If .EOF Then Exit For
                    For i = 1 To Len(rs.Fields(1))
                        Ace = UCase(Trim(rs.Fields(1)))
                       If Mid$(Ace, i, Len(Ss)) = Trim(UCase(Text1)) Then
                          List1.AddItem Trim(rs.Fields(0))
                          List2.AddItem Trim(rs.Fields(1))
                       End If
                    Next i
          .MoveNext
              Next t
      End With
  
ElseIf Me.Caption = "search by text" Then
    
          With rs
          .MoveFirst
              For t = 1 To rs.RecordCount
                 If rs.EOF Then Exit For
                     For i = 1 To Len(rs.Fields(1))
                         Ace = UCase(Trim(rs.Fields(3)))
                         If Mid$(Ace, i, Len(Ss)) = Trim(UCase(Text1)) Then
                             List1.AddItem Trim(.Fields(0))
                             List2.AddItem Trim(.Fields(1))
                         End If
                     Next i
          .MoveNext
              Next t
      End With
End If


End Sub


Private Sub Form_Load()
On Error GoTo errhandler
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb"


Set rs = New Recordset
rs.Open "select one,two,three,four,five,six from table2", db, adOpenStatic, adLockOptimistic
Exit Sub

errhandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
Set rs = Nothing
End Sub

Private Sub List1_Scroll()
    '--- lets List1 and List2 scroll together
    Call SendMessage(List2.hwnd, LB_SETTOPINDEX, SendMessage(List1.hwnd, LB_GETTOPINDEX, 0, 0), 0)

End Sub


Private Sub List2_Scroll()
    '--- lets List1 and List2 scroll together
    Call SendMessage(List1.hwnd, LB_SETTOPINDEX, SendMessage(List2.hwnd, LB_GETTOPINDEX, 0, 0), 0)
End Sub


Private Sub List2_Click()
Dim itemlist1 As String
Dim itemlist2 As String
Dim cat As String
Dim rat As String

      List1.ListIndex = List2.ListIndex

      itemlist1 = List1.List(List1.ListIndex)
      '--- the category from list1
      cat = itemlist1
      itemlist2 = List2.List(List1.ListIndex)
      '--- the Recipe Name from List2
      rat = itemlist2

      '--- call function to load the Records in
      '--- FrmMain Text Boxes
      DoIt

End Sub
Private Sub List1_Click()
Dim itemlist1 As String
Dim itemlist2 As String
Dim cat As String
Dim rat As String
    
    List2.ListIndex = List1.ListIndex

    itemlist1 = List1.Text
    '--- the category from list1
    cat = itemlist1
    itemlist2 = List2.Text
    '--- the Recipe Name from List2
    rat = itemlist2

    '--- call function to load the Records in
    '--- FrmMain Text Boxes
    DoIt

End Sub

Private Function DoIt()
On Error Resume Next
Dim i As Integer
   
        With rs
            .MoveFirst
                For i = 1 To rs.RecordCount
                   If Trim(rs.Fields(0)) = Trim(List1.Text) And Trim(rs.Fields(1)) = Trim(List2.Text) Then
                        frmMain.txtfields(0) = Trim(.Fields(1))
                        frmMain.txtfields(1) = Trim(.Fields(2))
                        frmMain.txtfields(2) = Trim(.Fields(3))
                        frmMain.txtfields(3) = Trim(.Fields(4))
                   Else
            .MoveNext
                   End If
                Next i
        End With


End Function



Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
KeyAscii = 0
cmdParse_Click
End If
End Sub
