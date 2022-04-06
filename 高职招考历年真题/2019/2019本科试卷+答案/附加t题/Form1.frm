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
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim i As Integer, j As Integer
Dim n As Integer
n = Val(Text1.Text)
For i = 1 To n Step 1
    For j = 1 To n Step 1
        ' 实现最外尾全部是1
        If i = 1 Or i = n Or j = 1 Or j = n Or i = j Or i + j = n + 1 Then
            Print 1;
        Else
            Print 0;
        End If
    Next j
    Print
Next i
'总结：
' i = j 正对角线
' i + j = n + 1 斜对角线
' i = 1 or i = n or j = 1 or j = n 最外围一圈



End Sub

