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
      Caption         =   "确定(&k)"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "正负号(&s)"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "绝对值(&a)"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个数"
      Height          =   495
      Left            =   480
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
If Option1.Value = True Then
    Label1 = Abs(Text1)
Else
    Select Case Val(Text1)
        Case Is > 0
            Label1 = "+"
        Case Is = 0
            Label1 = "0没有正负号"
        Case Is < 0
            Label1 = "-"
    End Select
End If

End Sub

