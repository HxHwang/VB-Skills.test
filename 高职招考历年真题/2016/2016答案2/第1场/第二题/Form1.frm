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
      Caption         =   "统计小写个数"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   1335
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "输入字符串"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As String, i As Integer, s1 As String, small As Integer
str = Text1
For i = 1 To Len(str)
     s1 = s1 & Mid(str, i, 1) & Space(1)
     Select Case Mid(str, i, 1)
        Case "a" To "z"
            small = small + 1
     End Select
Next i
Label2 = "间隔输出:" & s1
Label3 = "小写字母个数是：" & small
End Sub
