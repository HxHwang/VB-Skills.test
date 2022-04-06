VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求式数"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Single, ge As Integer, shi As Integer, bai As Integer, qian As Integer
Dim res As String
For i = 1000 To 9999
    ge = i Mod 10
    shi = i \ 10 Mod 10
    bai = i \ 100 Mod 10
    qian = i \ 1000
    res = ge & shi & bai & qian
    If i * 4 = Val(res) Then
        Print i
    End If
Next i
End Sub
'取出每个数的个位、十位、百位、千位
'利用拼凑法 拼出个结果！
