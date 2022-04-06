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
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1815
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

Dim i As Integer
Dim str As String
Dim zifu As String * 1
Dim res As String
Dim count As Integer
str = Text1.Text
count = 0
For i = 1 To Len(str) Step 1
    zifu = Mid(str, i, 1)
    res = res & zifu & Space(1)
    If zifu >= "a" And zifu <= "z" Then
        count = count + 1
    End If
Next i
Label2.Caption = "间隔输出：" & res
Label3.Caption = "小写字母个数是：" & count


End Sub
