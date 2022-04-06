VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5685
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "打印算法2"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印算法1"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Private Sub Command1_Click()
Dim num As Integer, i As Integer, j As Integer
num = Val(Text1)
For i = num To 1 Step -1
    For j = 1 To num
        Select Case j
            Case i: Print "0" & Space(2);
            Case Else: Print "1" & Space(2);
        End Select
    Next j
    Print
Next i
End Sub

Private Sub Command2_Click()
Dim a() As Integer, n As Integer
n = Val(Text1.Text)
ReDim a(n, n) As Integer
For i = 1 To n
    For j = 1 To n
        If i + j = n + 1 Then
            a(i, j) = "0"
        Else
            a(i, j) = "1"
        End If
    Next j
Next i

For i = 1 To n
    For j = 1 To n
       Print a(i, j);
    Next j
    Print
Next i
End Sub
