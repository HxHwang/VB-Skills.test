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
   Begin VB.CommandButton Command2 
      Caption         =   "统计"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n%

Private Sub Command1_Click()
Dim i%, a%
n = 0
For i = 1 To 10
a = Int(Rnd * 3 + (-1))
Print a;
If a = 0 Then n = n + 1
Next
End Sub

Private Sub Command2_Click()
Print
Print "0出现" & n & "次"
End Sub
