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
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "邮件重量(克)"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim n As Integer
Dim res As Integer
n = Val(Text1.Text)

Select Case n
    Case Is <= 1000
        res = 8
    Case Is > 1000
        res = 8 + ((n - 1000) / 500# + 1) * 4
End Select
MsgBox "邮资：" & res & "元"
' 1800 = 1000 + 500 + 300

End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
