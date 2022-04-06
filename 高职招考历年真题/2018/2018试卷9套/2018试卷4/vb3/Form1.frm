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
      Caption         =   "计算个税"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "输入工资"
      Height          =   495
      Left            =   720
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

Dim salary As Double
Dim res As Double
salary = Val(Text1.Text)

Select Case salary
    Case Is <= 5000
        res = 0
    Case 5000 To 10000
        res = (salary - 5000) * 0.05
    Case Is > 10000
        res = (salary - 10000) * 0.05 + (salary - 10000) * 0.1
End Select
' 15000 = 5000 + 5000 + 5000
'         不扣税  5%      10%
MsgBox "应缴税：" & res & "元"


End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
