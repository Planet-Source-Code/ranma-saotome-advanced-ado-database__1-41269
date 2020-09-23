VERSION 5.00
Begin VB.Form frmInfos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Infos"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   2880
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label lblloc 
      Caption         =   "N/A"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblsize 
      Caption         =   "N/A"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdParse_Click()
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sfile As String
sfile = App.Path & "\Info.txt"
Call File2ListBox(sfile, List1)
DBInfo
lblloc.Caption = "Database location: " & App.Path
End Sub

Private Sub DBInfo()
Dim g As Long
On Error Resume Next

Open App.Path & "\Db1.mdb" For Binary As #1
g = LOF(1)
Close #1
lblsize.Caption = "Size of database : " & Format(g, "###,###,###,##0") & " k"

End Sub
