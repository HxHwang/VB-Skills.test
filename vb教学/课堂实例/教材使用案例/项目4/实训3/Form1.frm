VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6030
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������Ф"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ʮ����Ф"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim year As Integer, a As Integer
Private Sub Command1_Click()
year = Val(Text1.Text)
a = year Mod 12
Select Case a
Case 1
   Text2.Text = "��"
Case 2
   Text2.Text = "��"
Case 3
   Text2.Text = "��"
Case 4
   Text2.Text = "��"
Case 5
   Text2.Text = "ţ"
Case 6
   Text2.Text = "��"
Case 7
   Text2.Text = "��"
Case 8
   Text2.Text = "��"
Case 9
   Text2.Text = "��"
Case 10
   Text2.Text = "��"
Case 11
   Text2.Text = "��"
Case 0
   Text2.Text = "��"
End Select


End Sub

