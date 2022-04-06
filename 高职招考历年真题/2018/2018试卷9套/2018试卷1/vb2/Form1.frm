VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check2 
      Caption         =   "最小化到系统托盘"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "总是显示系统托盘图标"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "新标签页"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "主页"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "主页"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "系统托盘"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "启动时"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
