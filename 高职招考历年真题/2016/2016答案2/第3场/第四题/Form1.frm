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
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "统计并输出"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "初始化"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Private Sub Command1_Click()
Dim zero As String
zero = Command1.Caption
str = str & zero
End Sub

Private Sub Command2_Click()
Dim one As String
one = Command2.Caption
str = str & one
End Sub

Private Sub Command3_Click()
str = ""
Label1 = ""
Label2 = ""
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim char As String, zeroCount As Integer, oneCount As Integer
'for循环每一个字符，统计0的个数和1的个数
For i = 1 To Len(str)
    '获取二进制的每个字符
    char = Mid(str, i, 1)
    '判断所获得的字符是zero还是one
    If char = Command1.Caption Then
        zeroCount = zeroCount + 1
    Else
        oneCount = oneCount + 1
    End If
Next i
Label1 = "这个二进制数是：" & str
Label2 = "这个数中零的个数是：" & zeroCount & "," & "1的个数是：" & oneCount
End Sub

Private Sub Command5_Click()
End
End Sub
