VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "����_GB2312"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   6255
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Dim a As Double
Dim r As Single
Dim i As Integer
a = 12
r = 0.01
i = 0
Do While a < 20     '���˿������ڵ���20��ʱ����ѭ��
a = a * (1 + r)
i = i + 1
Loop
Print i; "����й��˿ڴﵽ20��"

End Sub

