VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QQµÇÂ¼"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "µÇÂ¼"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¼Ç×¡ÃÜÂë"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "×Ô¶¯µÇÂ¼"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "VB2.frx":0000
      Left            =   1440
      List            =   "VB2.frx":000A
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "ÃÜÂë"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "QQºÅÂë"
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
