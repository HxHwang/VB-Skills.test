VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "У�����"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "ABC" Then
        MsgBox "��ȷ"
    Else
        MsgBox "����"
    End If
End Sub
