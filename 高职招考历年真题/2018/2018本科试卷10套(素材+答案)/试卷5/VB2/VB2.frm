VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "选课"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "关闭"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   1320
      ItemData        =   "VB2.frx":0000
      Left            =   2760
      List            =   "VB2.frx":000A
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1320
      ItemData        =   "VB2.frx":001A
      Left            =   360
      List            =   "VB2.frx":002A
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "已选课程"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "选择您的课程"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
