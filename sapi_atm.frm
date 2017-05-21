VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CURRENT BALANCE"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String
Private Sub Form_Load()
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Documents\SAPI.mdb;Persist Security Info=False"
conn.Open
Dim a, b As Integer
a = Form4.Text1.Text
sqlStr = "select balance from user_detail where id=" & Form1.Text1.Text
rs.Open (sqlStr), conn, adOpenDynamic, adLockOptimistic
'Text1.Text = rs.Fields("balance")
b = Int(rs.Fields("balance"))
If (b >= a) Then
rs.Fields("balance") = b - a
rs.Update

Text1.Text = rs.Fields("balance")
a = MsgBox("amount withdrawn successfully", vbOKOnly)
Else
Text1.Text = rs.Fields("balance")
a = MsgBox("amount entered not available in balance", vbOKOnly)
End If
End Sub
