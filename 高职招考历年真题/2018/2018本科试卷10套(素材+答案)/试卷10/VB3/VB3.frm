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
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "����y"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "����x"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    y = Val(Text2.Text)
    If x > 0 And y > 0 Then
        MsgBox "�ڵ�һ����"
    ElseIf x < 0 And y > 0 Then
        MsgBox "�ڵڶ�����"
    ElseIf x < 0 And y < 0 Then
        MsgBox "�ڵ�������"
    Else
        MsgBox "�ڵ�������"
    End If
End Sub
