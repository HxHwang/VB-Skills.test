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
      Caption         =   "显示"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1 ' 下标的下界从1开始
Private Sub Command1_Click()

Text1.Text = ""
Dim i As Integer
Dim a(10) As Integer
Dim ji As Integer, ou As Integer

Text1.Text = Space(1) ' 加个空格，与题目保持一致
For i = 1 To 10 Step 1
    a(i) = Int(Rnd * 90 + 10) ' 范围[10,99]
    Text1.Text = Text1.Text & a(i) & Space(1)
    ' 统计奇数和偶数的个数
    If a(i) Mod 2 = 0 Then
        ou = ou + 1
    Else
        ji = ji + 1
    End If
Next i
Label1.Caption = "奇数的个数是：" & ji & "偶数的个数是：" & ou

End Sub
