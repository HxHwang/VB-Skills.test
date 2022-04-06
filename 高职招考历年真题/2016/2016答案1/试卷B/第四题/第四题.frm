VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   11985
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印零巧数"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 1000 To 9999
    ge = i Mod 10
    shi = i \ 10 Mod 10
    bai = i \ 100 Mod 10
    qian = i \ 1000
    If bai = 0 Then
        If i = (qian * 100 + shi * 10 + ge) * 9 Then Print i
    End If
Next i
End Sub
