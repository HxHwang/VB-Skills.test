VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����ݵľ���ֵ"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "����_GB2312"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4335
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "|a|"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Integer    '�������
    a = Val(Text1.Text)     '���ı����������ֵ��������a
    If a < 0 Then   '��aΪ����ʱȡ���෴��
    a = -a
    End If
    Text2.Text = Str$(a)
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub
