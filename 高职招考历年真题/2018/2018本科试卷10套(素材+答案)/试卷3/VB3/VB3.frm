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
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "��ѧ�ɼ�"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "���ĳɼ�"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    If a >= 90 And b >= 90 Then
        MsgBox "������㽱ѧ��"
    ElseIf a = 100 Or b = 100 Then
        MsgBox "��õ��ѧ��"
    Else
        MsgBox "û�н�ѧ��"
    End If
End Sub
