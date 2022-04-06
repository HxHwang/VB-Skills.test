VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5505
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
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

Dim i As Integer, j As Integer
Dim n As Integer
n = Val(Text1.Text)
For i = 1 To n Step 1
    For j = 1 To n Step 1
        ' 输出矩阵的下半部分
        If i >= j Then
            Print 1;
        Else
            Print 0;
        End If
    Next j
    Print
Next i


End Sub
