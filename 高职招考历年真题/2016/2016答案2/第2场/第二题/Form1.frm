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
      Caption         =   "显示"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'随机数公式 Int(Rnd*(n-m+1)+m)
'要求：生成2位数 11-99 的数字
Dim i As Integer, str As String, num As Integer
Dim jishu As Integer, oushu As Integer
str = Space(1) & ""
For i = 1 To 10
    num = Int(Rnd * 89 + 11)
    str = str & num & Space(1)
    If num Mod 2 = 0 Then
        oushu = oushu + 1
    Else
        jishu = jishu + 1
    End If
Next i
    Text1 = str
    Label1 = "奇数的个数是：" & jishu & "偶数的个数是：" & oushu
End Sub
