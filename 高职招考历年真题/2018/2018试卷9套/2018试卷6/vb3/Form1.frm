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
      Caption         =   "判断"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "输入距离"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim length As Single
Dim buxing As Single
Dim qixing As Single
length = CSng(Text1.Text)
buxing = length / 1.2
qixing = length / 3 + 27 + 23

If buxing < qixing Then
    MsgBox "步行快"
ElseIf buxing > qixing Then
    MsgBox "骑行快"
Else
    MsgBox "一样快"
End If

End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
