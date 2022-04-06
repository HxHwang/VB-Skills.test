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
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "通话时长（分钟）"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim phoneLength As Single
phoneLength = CSng(Text1.Text)

If phoneLength > 0 And phoneLength <= 3 Then
    res = 0.5
Else
    res = 0.5 + (phoneLength - 3) * 0.15
    ' 总话费封顶10元
    If res > 10 Then
        res = 10
    End If
End If
MsgBox "话费：" & res & ""

End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
