VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�жϸ���"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ż����"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
num = Val(InputBox("����������", "�жϸ���", "-40"))
If num < 0 Then
    Print num
Else
    MsgBox "���Ǹ��������������룡"
End If

End Sub

Private Sub Command2_Click()
Dim m As Integer, n As Integer
Dim i As Integer, sum As Integer
sum = 0
m = Val(Text1.Text)
n = Val(Text2.Text)
If m < n Then
    For i = m To n
        If i Mod 2 = 0 Then
            sum = sum + i
        End If
    Next i
Else
    MsgBox "m>n������������"
End If
Print sum
End Sub
