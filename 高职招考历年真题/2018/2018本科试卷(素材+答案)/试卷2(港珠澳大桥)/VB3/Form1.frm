VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "开始打印图形"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "请输入1--9之间的整数："
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    n = Val(Text1.Text)
    For i = n To 1 Step -1
        For j = 1 To 10 - i
            Print " ";
        Next j
        For j = 1 To 2 * i - 1
            Print "*";
        Next j
        Print
    Next i
End Sub
