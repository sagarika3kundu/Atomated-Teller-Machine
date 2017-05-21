VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "WITHDRAW"
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK BALANCE"
      Height          =   1575
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Visible = True

End Sub

Private Sub Command2_Click()
Form4.Visible = True

End Sub

