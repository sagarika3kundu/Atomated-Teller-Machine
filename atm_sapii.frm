VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SUBMIT"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6600
      TabIndex        =   3
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PASSWORD"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARD NO"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String


Private Sub Command4_Click()
sqlStr = "select * from user_detail where id=" & Text1.Text & "and password=" & Text2.Text
'sqlStr = "select * from user_detail"
rs.Open (sqlStr), conn, adOpenStatic, adLockReadOnly
'Print rs.Fields("id")

If (rs.RecordCount > 0) Then
Form2.Visible = True
Else
a = MsgBox("enter valid id or password", vbOKOnly, "atm")
End If
rs.Close

End Sub

Private Sub Form_Load()
'Set conn = New ADODB.Connection
'Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Documents\SAPI.mdb;Persist Security Info=False"
conn.Open
'sqlStr = "select * from user_detail"
'rs.Open (sqlStr), conn, adOpenDynamic, adLockOptimistic
End Sub
