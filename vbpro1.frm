VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "YOUR CURRENT BALANCE"
      Height          =   975
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Documents\SAPI.mdb;Persist Security Info=False"
conn.Open

sqlStr = "select balance from user_detail where id=" & Form1.Text1.Text
rs.Open (sqlStr), conn, adOpenStatic, adLockReadOnly
Text1.Text = rs.Fields("balance")

End Sub

