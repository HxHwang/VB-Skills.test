VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Բ�ܳ������������"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "����_GB2312"
      Size            =   15
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4380
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�ܳ���"
      Height          =   300
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�뾶R��"
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '�������
    Dim R As Double
    Dim L As Double
    Dim S As Double
    '���峣��
    Const PI = 3.1416
    '��ȡ�뾶R��ֵ
    R = Text1.Text
    '����Բ�ܳ������
    L = 2 * PI * R
    S = PI * R * R
    '���Բ�ܳ��������ֵ
    Text2.Text = L
    Text3.Text = S
End Sub
