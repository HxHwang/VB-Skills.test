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
      Caption         =   "ת��"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "����ٷ��Ƴɼ�"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    If x >= 90 Then
        MsgBox "����"
    ElseIf x >= 80 Then
        MsgBox "����"
    ElseIf x >= 70 Then
        MsgBox "�е�"
    ElseIf x >= 60 Then
        MsgBox "����"
    Else
        MsgBox "������"
    End If
End Sub
