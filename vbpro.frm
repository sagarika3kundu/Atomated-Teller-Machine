VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CONFIRM"
      Height          =   855
      Left            =   7680
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER AMOUNT"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   6720
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String

Private Sub Command2_Click()
Form5.Show
End Sub

