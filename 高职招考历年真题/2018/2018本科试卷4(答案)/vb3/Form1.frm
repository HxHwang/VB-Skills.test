VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7710
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "��ֶκ���"
      Height          =   1575
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����x"
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = Val(Text1)
If x < 0 Then
   x = x + 3
   b = MsgBox("���x��ֵΪ��" & x, 0, "vb3")
ElseIf x > 0 And x < 10 Then
   x = x / 2
    b = MsgBox("���x��ֵΪ��" & x, 0, "vb3")
ElseIf x <= 10 Then
   x = Sqr(x) - 3
    b = MsgBox("���x��ֵΪ��" & x, 0, "vb3")
End If
End Sub
