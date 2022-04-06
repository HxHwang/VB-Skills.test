VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问卷调查"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "VB2.frx":0000
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "运动"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "女"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "男"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "运动心得"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "日常活动"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "您的性别"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
