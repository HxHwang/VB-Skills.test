VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ˮ�ɻ���"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾˮ�ɻ���"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, ge As Integer, shi As Integer, bai As Integer
Print "ˮ�ɻ����У�"
For i = 100 To 999
    ge = i Mod 10
    shi = i \ 10 Mod 10
    bai = i \ 100
    If (ge * ge * ge) + (shi * shi * shi) + (bai * bai * bai) = i Then
        Print i
    End If
Next i

End Sub
