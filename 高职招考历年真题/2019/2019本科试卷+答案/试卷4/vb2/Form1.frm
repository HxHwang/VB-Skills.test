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
      Caption         =   "最大公约数"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim i As Integer
Dim a As Integer, b As Integer
Dim max As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
max = 0
' 对任意变量循环
For i = 1 To a Step 1
    ' 判断两个整数，能同时被一个数整除
    If a Mod i = 0 And b Mod i = 0 Then
        ' 找到最大公约数
        If i > max Then
            max = i
        End If
    End If
Next i
Label1.Caption = max

End Sub
