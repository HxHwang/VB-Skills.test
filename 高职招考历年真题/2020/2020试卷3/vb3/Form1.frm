VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, a(1 To 10) As Integer
Dim sum As Double
sum = 0
For i = 1 To 10
a(i) = Int(Rnd * 101 + 0)
Print a(i);
If a(i) Mod 2 = 0 Then
sum = sum + a(i) ^ 2
End If
Next
Print
Print "偶数平方和为：" & sum
End Sub
