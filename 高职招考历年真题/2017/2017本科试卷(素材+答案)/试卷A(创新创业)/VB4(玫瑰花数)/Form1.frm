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
      Caption         =   "求玫瑰花数"
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ge As Integer, shi As Integer, bai As Integer, qian As Integer
Dim i As Integer
For i = 1000 To 9999
    ge = i Mod 10
    shi = i \ 10 Mod 10
    bai = i \ 100 Mod 10
    qian = i \ 1000
    If ge * ge * ge * ge + shi * shi * shi * shi + bai * bai * bai * bai + qian * qian * qian * qian = i Then
        Print i
    End If
Next i
End Sub
