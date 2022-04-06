VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "问卷调查"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "运动"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "女"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "男"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "运动心得"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "日常活动"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "您的性别"
      Height          =   495
      Left            =   480
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
