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
      Caption         =   "求奇数平方和"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Private Sub Command1_Click()

Dim a(10) As Integer
Dim i As Integer
Dim sum As Integer
sum = 0

For i = 1 To 10 Step 1
    ' 范围：[10,50]
    a(i) = Int(Rnd * 41 + 10)
    Print a(i);
Next i
Print

' 计算奇数平方和
For i = 1 To 10 Step 1
    If a(i) Mod 2 = 1 Then
        sum = sum + a(i) * a(i)
    End If
Next i
Print "奇数平方和：" & sum


End Sub
