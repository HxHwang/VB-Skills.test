VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   975
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As String
str = Text1.Text
If str = "ABC" Then
    MsgBox "��ȷ"
Else
    MsgBox "����"
End If

End Sub
