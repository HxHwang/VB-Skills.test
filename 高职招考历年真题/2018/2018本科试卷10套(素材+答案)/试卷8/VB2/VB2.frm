VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ʱ��Ӧ��"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      Caption         =   "����߽��Զ�ֹͣ"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ֹͣ"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   2040
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�����ƶ�"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�����ƶ�"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "�����ƶ�"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
