VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "整点报时"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1640
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "24小时制"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Text            =   "11:10:25"
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "时钟显示"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
