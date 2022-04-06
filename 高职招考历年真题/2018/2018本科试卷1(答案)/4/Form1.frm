VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   8820
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "最大值"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(6, 6), max
a(1, 1) = "姓名"
a(1, 6) = "最大值"
For i = 2 To 5
a(1, i) = "科目" & i - 1
a(i, 1) = "学生" & i - 1
Next i
For i = 2 To 5
Randomize
For j = 2 To 5
a(i, j) = Int(Rnd * 50 + 50)
Next j
Next i
For l = 2 To 5
For i = 3 To 5
If a(l, i) > max Then max = a(l, i)
Next i
a(l, 6) = max
Next l
For i = 1 To 5
For j = 1 To 6
Print a(i, j),
s = s + 1
If s Mod 6 = 0 Then Print
Next j
Next i

End Sub
