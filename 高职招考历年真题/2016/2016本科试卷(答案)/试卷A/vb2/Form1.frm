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
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�жϽ��"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim z As Single
z = Val(Text1.Text)
If z >= 60 Then
Label1.Caption = "���أ����ܲ���!"
Else
Label1.Caption = "�ϸ񣬿��Բ���!"
End If
End Sub

Private Sub Command2_Click()
End
End Sub
