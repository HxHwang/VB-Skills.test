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
      Caption         =   "�ж�"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    t1 = (x - 3) * 2 + 10
    t2 = (x - 3) * 2.5 + 8
    If x <= 3 Then t1 = 10: t2 = 8
    If t1 > t2 Then
        MsgBox "����һ����"
    ElseIf t1 < t2 Then
        MsgBox "����������"
    Else
        MsgBox "һ����"
    End If
End Sub
