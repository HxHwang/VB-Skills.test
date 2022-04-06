VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1 ' 下标的下界从1开始
Private Sub Command1_Click()

Dim i As Integer
Dim a(10) As Integer
Dim max As Integer

For i = 1 To 10 Step 1
    a(i) = Int(Rnd * 900 + 100) ' 范围：[100,999]
    Text1.Text = Text1.Text & a(i) & Space(1)
Next i

' 找最大值
max = a(1)
For i = 2 To 10 Step 1
    If a(i) > max Then
        max = a(i)
    End If
Next i
Label1.Caption = "十个数中最大的数是：" & max



End Sub
