VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   7095
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "×Ô¶¯µÇÂ½"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "µÇÂ½"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "132465"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   2760
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÃÜÂë"
      Height          =   180
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "qqÕËºÅ"
      Height          =   180
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
