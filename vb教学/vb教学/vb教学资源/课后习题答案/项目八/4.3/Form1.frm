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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Form1.AutoRedraw = True
    If Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Then
        Print "���µ������ּ�"
    ElseIf Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "z" Then
        Print "���µ�����ĸ��"
    End If
End Sub

