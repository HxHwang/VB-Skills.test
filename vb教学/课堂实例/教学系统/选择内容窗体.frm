VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "��ѡ���ѧ����"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   Picture         =   "ѡ�����ݴ���.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11190
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   6960
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����γ�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   5040
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Form2.WindowState = 2
With Combo1
     .AddItem "��5�� ���ӱ�������Ӧ��"
     .Text = "��5�� ���ӱ�������Ӧ��"
End With
With Combo2
     .AddItem "5.4 ����ʹ�������"
     .Text = "5.4 ����ʹ�������"
End With
End Sub

