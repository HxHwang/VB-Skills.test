VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check2 
      Caption         =   "��С����ϵͳ����"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "������ʾϵͳ����ͼ��"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ҳ"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�±�ǩҳ"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��ҳ"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "ϵͳ����"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "����ʱ"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
