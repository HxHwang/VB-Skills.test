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
      Caption         =   "转换(&k)"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "小写(&L)"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "大写(&U)"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入字符"
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As String
str = Text1
If Option1 = True Then
    Text1 = UCase(str)
Else
    Text1 = LCase(str)
End If
End Sub
