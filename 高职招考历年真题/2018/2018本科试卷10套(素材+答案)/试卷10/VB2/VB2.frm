VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ͼ��ɹ�"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5280
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1740
      ItemData        =   "VB2.frx":0000
      Left            =   480
      List            =   "VB2.frx":0010
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "ͼ���嵥"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
