VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Cls ' 输出的同时，清空上一次的结果
Dim i As Integer, j As Integer
Dim n As Integer
n = Val(Text1.Text)

If n <= 0 Or n > 20 Then
    MsgBox "请输入1-20的数字"
    Exit Sub
End If


For i = 1 To n Step 1
    For j = 1 To n Step 1
        ' 斜对角线的值为0
        If i + j = n + 1 Then
            Print 0;
        Else
            Print 1;
        End If
    Next j
    Print
Next i


End Sub
