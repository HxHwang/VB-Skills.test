VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ж��Ƿ�Ϊ������"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   7050
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�ж��Ƿ�Ϊ������"
      Height          =   180
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "�������֤"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


a = Text1

If Len(a) <> 18 Then

MsgBox ("�������������֤���룡")

Text1.Text = " "

Else
d = Mid(a, 1, 2)
  If d = 35 Then
   Print "�Ǹ�����"
  Else
   Print "���Ǹ�����"
  End If

End If
End Sub
