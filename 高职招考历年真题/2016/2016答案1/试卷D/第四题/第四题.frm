VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示最大公约数"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "输入两个整数："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "答案："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
i = Text1
Do Until Text1 Mod i = 0 And Text2 Mod i = 0
    i = i - 1
Loop
Label3 = i
End Sub
