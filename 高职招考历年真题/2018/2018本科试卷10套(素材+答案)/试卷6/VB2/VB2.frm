VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "安装程序"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check2 
      Caption         =   "安装后重启"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动安装"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览..."
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
      Caption         =   "安装位置"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "请选择你要安装的程序"
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
