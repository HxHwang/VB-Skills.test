VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ж�ż��"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "���۵ȼ�"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ż��"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim a As Integer
a = InputBox("����������", "�ж�ż��")
If a Mod 2 = 0 Then
    Print a
End If
End Sub

Private Sub Command2_Click()
Dim a As interger
a = InputBox("������ɼ�", "���۵ȼ�")

    Select Case a
        Case 81 To 100
            Print "����"
        Case 60 To 80
            Print "�ϸ�"
        Case 0 To 59
            Print "������"
        Case Else
     MsgBox "�������", , "����"
        End Select

End Sub
