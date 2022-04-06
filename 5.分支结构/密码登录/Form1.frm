VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "√‹¬Îµ«¬º"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command1 
      Caption         =   "»∑∂®"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "«Î ‰»Î√‹¬Î£∫"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "123" Then
  Form2.Show
  Form1.Hide
Else
  MsgBox "√‹¬Î¥ÌŒÛ£°«Î÷ÿ ‘£°"
End If
End Sub
