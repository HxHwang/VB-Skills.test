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
      Caption         =   "ת��"
      Height          =   495
      Left            =   1560
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
      Caption         =   "����ٷ��Ƴɼ�"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim n As Integer
Dim res As String
n = Val(Text1.Text)

Select Case n
    Case Is >= 90
        res = "����"
    Case 80 To 90
        res = "����"
    Case 70 To 80
        res = "�е�"
    Case 60 To 70
        res = "����"
    Case 0 To 60
        res = "������"
    Case Else
        res = "������0~100������"
End Select

' ���������Ϣ�����ʽ����ʾ����
a = MsgBox(res, , "VB3")


End Sub
