VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&k)"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "正负号(&b)"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "绝对值(&a)"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个数"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
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
n = Val(Text1.Text)
If Option1.Value = True Then
    Label2.Caption = Abs(n)
Else
    If Sgn(n) > 0 Then
        Label2.Caption = "+"
    ElseIf Sgn(n) < 0 Then
        Label2.Caption = "-"
    Else
        Label2.Caption = "0为无效值"
    End If
End If




End Sub

