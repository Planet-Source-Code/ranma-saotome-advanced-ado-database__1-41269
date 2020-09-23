VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCatagory 
   Caption         =   "Category Menu"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView List1 
      Height          =   1695
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&My Value here"
      Height          =   320
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Present Categorys:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCatagory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Private Sub Command1_Click()
On Error Resume Next
rs.Close
Set rs = Nothing
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo errhandler:
Dim strCata As String

If Command2.Caption = "&Add new category" Then
        strCata = InputBox("Enter New Category Name", "Add New Category")
        If Trim(strCata) = "" Then Exit Sub
        rs.AddNew
        rs.Fields(0) = Trim(strCata)
        rs.Update
        Fillme
        Exit Sub
ElseIf Command2.Caption = "&Delete category" Then
            If MsgBox("Are you shure to delete the category " & List1.SelectedItem.Text & " permanently? ", vbQuestion + vbYesNo) = vbYes Then
            rs.Delete
            rs.Update
            Fillme
            Exit Sub
            End If
End If

Exit Sub
errhandler:
    MsgBox Err.Description
End Sub



Private Sub Form_Load()
On Error GoTo errhandler:
Dim db As Connection
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb"

Set rs = New Recordset
rs.Open "select one from table1", db, adOpenStatic, adLockOptimistic
 
List1.ColumnHeaders.Add Text:="Categorys", Width:=2200
Fillme


Exit Sub
errhandler:
    MsgBox Err.Description
End Sub


Private Sub Fillme()
List1.ListItems.Clear
rs.MoveFirst
 While Not rs.EOF
   List1.ListItems.Add , , rs("one")
   rs.MoveNext
 Wend
rs.MoveFirst
End Sub

Private Sub List1_Click()
On Error Resume Next
Call Search(List1.SelectedItem.Text, rs, rs.Fields("one"))
End Sub
