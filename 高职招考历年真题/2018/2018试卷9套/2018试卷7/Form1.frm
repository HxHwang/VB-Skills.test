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
      Caption         =   "�ж�"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim juli As Integer
Dim a As Integer
Dim b As Integer
juli = Val(Text1.Text)
' (������ - 3 / ÿ1����) * 2
a = 10 + (juli - 3) * 2
b = 8 + (juli - 3) * 2.5

If a > b Then
    MsgBox "����һ��"
ElseIf a < b Then
    MsgBox "��������"
Else
    MsgBox "һ����"
End If


End Sub
