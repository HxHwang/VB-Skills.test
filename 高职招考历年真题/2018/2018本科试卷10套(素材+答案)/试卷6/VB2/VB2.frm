VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��װ����"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check2 
      Caption         =   "��װ������"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�Զ���װ"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���..."
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "VB2.frx":0000
      Left            =   480
      List            =   "VB2.frx":0016
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D:\"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "��װλ��"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ����Ҫ��װ�ĳ���"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
