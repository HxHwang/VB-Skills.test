VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "������ƽ����"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For i = 1 To 10
        x = Int(Rnd * 41 + 10)
        Print x;
        If x Mod 2 = 1 Then s = s + x * x
    Next i
    Print
    Print "����ƽ���ͣ�"; s
End Sub
