VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB3"
   ClientHeight    =   3015
   ClientLeft      =   7905
   ClientTop       =   2415
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
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
Dim a(1 To 10) As Integer
Private Sub Command1_Click()
Dim i%
n = 0
For i = 1 To 10
a(i) = Int(Rnd * 3 + (-1))
Print a(i);
Next
End Sub

Private Sub Command2_Click()
Dim n1%, n2%, n3%
For i = 1 To 10
If a(i) = 0 Then n1 = n1 + 1
'If a(i) = -1 Then n2 = n2 + 1
'If a(i) = 1 Then n3 = n3 + 1
Next
Print
Print "0出现" & n1 & "次"
'Print "-1出现" & n2 & "次"
'Print "1出现" & n3 & "次"
End Sub
