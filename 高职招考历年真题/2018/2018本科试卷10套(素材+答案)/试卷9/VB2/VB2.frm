VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "离开、忙碌时自动回复"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "VB2.frx":0000
      Left            =   1320
      List            =   "VB2.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "启动时自动登录"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "开机自动登录"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "状态"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "登录"
      Height          =   495
      Left            =   480
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
