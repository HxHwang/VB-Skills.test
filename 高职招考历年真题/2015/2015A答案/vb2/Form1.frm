VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ж�����"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�ȼ�����"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ����"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub Command1_Click()
Cls
x = Val(InputBox("����������", "�ж�����", 0))
If x >= 0 Then Print x
End Sub

Private Sub Command2_Click()
Cls
Select Case x
Case 1 To 4
Print "D"
Case 5 To 10
Print "C"
Case 11 To 14
Print "B"
Case Else
Print "A"
End Select
End Sub
