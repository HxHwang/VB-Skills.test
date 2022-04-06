VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   7545
   ClientTop       =   2400
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "统计"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i%, a(10 To 19) As Integer
Private Sub Command1_Click()
For i = 10 To 19
a(i) = Int(Rnd * 11 + 10)
Print a(i);
Next
End Sub
Private Sub Command2_Click()
Dim j, b(10 To 20), max As Integer
For i = 10 To 19
  For j = 10 To 20
  If a(i) = j Then b(j) = b(j) + 1
  Next
Next
max = b(10)
For j = 11 To 20
If b(j) > max Then max = b(j)
Next
Print
For j = 11 To 20
If b(j) = max Then Print j;
Next
Print "出现的次数最多，共出现" & max; "次"
End Sub
