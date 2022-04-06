VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   11895
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印式数"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
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
    If i * 4 = ge * 1000 + shi * 100 + bai * 10 + qian Then Print i
Next i
End Sub

