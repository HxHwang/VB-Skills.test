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
   Begin VB.CommandButton Command5 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ͳ�Ʋ����"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʼ��"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Private Sub Command1_Click()

str = str & Command1.Caption

End Sub

Private Sub Command2_Click()

str = str & Command2.Caption

End Sub

Private Sub Command3_Click()
str = ""
Label1.Caption = ""
End Sub

Private Sub Command4_Click()

Dim i As Integer
Dim zero As Integer, one As Integer
Dim zifu As String * 1
'���ַ���ͳ�Ʋ����
For i = 1 To Len(str) Step 1
    zifu = Mid(str, i, 1)
    If zifu = "0" Then
        zero = zero + 1
    Else
        one = one + 1
    End If
Next i

' ������
Label1.Caption = "������������ǣ�" & str & vbCrLf & "���������ĸ����ǣ�" & zero & ",1�ĸ����ǣ�" & one

End Sub

Private Sub Command5_Click()
    End
End Sub

