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
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������ĸ2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������ĸ1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Private Sub Command1_Click()
Dim x As String
x = InputBox("��������ĸ", "������ĸ��", "A")
str = str & x & Space(1)
End Sub

Private Sub Command2_Click()
Dim y As String
y = InputBox("��������ĸ", "������ĸ��", "A")
str = str & y & Space(1)
End Sub

Private Sub Command3_Click()
Label1 = "��ĸ1����ĸ2��������ĸ�У�" & str
End Sub

Private Sub Command4_Click()
End
End Sub

