VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ݴ���"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3345
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
num = Int(Rnd * 900 + 100)
Text1.Text = num

End Sub

Private Sub Command2_Click()
Dim str As String, i As Integer, n As Integer, a As Integer
Dim str2 As String
str = Text1.Text
n = Len(str)
For i = 1 To Len(str)
    a = Mid(str, n - i + 1, 1)
    str2 = str2 & a
Next i
Text2.Text = str2
End Sub
