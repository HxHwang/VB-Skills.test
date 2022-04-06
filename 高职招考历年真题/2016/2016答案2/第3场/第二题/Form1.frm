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
      Caption         =   "计算"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, num As Integer, str As String
Dim max As Integer
str = ""
max = 0
For i = 1 To 10
    num = Int(Rnd * 900 + 100)
    If num > max Then
        max = num
    End If
    str = str & num & Space(1)
Next i
Text1 = str
Label1 = "十个数中最大的数是：" & max
End Sub

