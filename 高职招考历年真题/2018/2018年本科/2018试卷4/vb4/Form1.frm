VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   8130
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印个数"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "打印10个随机数，判断个位上面为3的个数有几个"
      Height          =   1215
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub Command1_Click()
Dim a(10) As Integer
 For i = 1 To 10
   a(i) = Int(Rnd * 89 + 11)
   Print a(i);
 Next i
 Print
    For i = 1 To 10
     j = Mid(a(i), 2, 1)
     If j = 3 Then n = n + 1
    Next i
    Print "个数为3的个数为："; n

End Sub


