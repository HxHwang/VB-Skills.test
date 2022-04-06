VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4335
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton C2 
      Cancel          =   -1  'True
      Caption         =   "否"
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   800
   End
   Begin VB.CommandButton C1 
      Caption         =   "是"
      Default         =   -1  'True
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   800
   End
   Begin VB.Label L1 
      Caption         =   "请确认"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
