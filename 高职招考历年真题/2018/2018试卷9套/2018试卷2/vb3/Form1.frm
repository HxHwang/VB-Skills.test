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
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "数学成绩"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "语文成绩"
      Height          =   495
      Left            =   720
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

Dim a As Integer, b As Integer

a = Val(Text1.Text)
b = Val(Text2.Text)

' 判断范围是否满足
If a < 0 Or a > 100 Or b < 0 Or b > 100 Then
    MsgBox "请输入0~100范围内的整数"
    Exit Sub
End If

If a >= 90 And b >= 90 Then
    res = "获得单项奖学金"
ElseIf a = 100 Or b = 100 Then
    res = "获得单项奖学金"
Else
    res = "没有奖学金！"
End If

MsgBox res




End Sub
