VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "��ý���ѧ�μ�"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form3"
   Picture         =   "��ѧ.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ѧ�μ�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ʵ����ϰ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ:����·��ġ���ѧ�μ�����       ť�ڸ�������ʾ��ѧ�μ�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3960
      TabIndex        =   4
      Top             =   3840
      Width           =   7215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   13215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.ShockwaveFlash1.Movie = App.Path & "\123.swf"
Command1.Enabled = False
Label2.Caption = ""
End Sub

Private Sub Command2_Click()
Form3.Hide
form5.Show
End Sub

Private Sub Command3_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Form_Load()
Command1.Enabled = True
Form3.WindowState = 2
End Sub

