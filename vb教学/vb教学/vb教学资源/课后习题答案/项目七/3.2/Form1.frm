VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command1 
      Caption         =   "œ‘ æ"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim C(7) As Integer

Private Sub Command1_Click()
Dim i As Integer

A = Array(2, 8, 7, 6, 4, 28, 25, 30)
B = Array(79, 27, 32, 41, 57, 66, 78, 80)
For i = 0 To 7
    C(i) = A(i) + B(i)
    Print C(i)
Next i

End Sub
