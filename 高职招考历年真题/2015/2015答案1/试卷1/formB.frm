VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ж�����"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�ȼ�����"
      Height          =   735
      Left            =   3435
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ����"
      Height          =   735
      Left            =   1155
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim a%

a = Val(InputBox("����������"))
If a > 0 Then Print a

End Sub

Private Sub Command2_Click()

'����һ�� �Ƽ���
Dim a As Integer
a = Val(InputBox("����������", "�ȼ�����"))
Select Case a
Case 1 To 4
    Print "D"
Case 5 To 10
    Print "C"
Case 11 To 14
    Print "B"
Case Else            ' �� Case Is > 14
    Print "A"
End Select

'������

'If a > 14 Then
'Print "A"
'ElseIf a >= 11 Then
'Print "B"
'ElseIf a >= 5 Then
'Print "C"
'ElseIf a >= 1 Then
'Print "D"
'End If

End Sub
