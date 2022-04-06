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
      Caption         =   "求个数"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
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
Dim count As Integer
count = 0

For i = 1 To 10 Step 1
    a(i) = Int(Rnd * 301 + 300)
    Print a(i);
Next i
Print ' 换行

' 计算个数
For i = 1 To 10 Step 1
    If a(i) Mod 13 = 0 Then
        count = count + 1
    End If
Next i
Print "能被13整除数的个数：" & count

End Sub
