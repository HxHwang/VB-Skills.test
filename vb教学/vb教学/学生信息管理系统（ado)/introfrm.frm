VERSION 5.00
Begin VB.Form introfrm 
   Caption         =   "��ӭ����"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6510
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��¼ϵͳ"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ѧ����Ϣ����ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "introfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
login.Show
End Sub
