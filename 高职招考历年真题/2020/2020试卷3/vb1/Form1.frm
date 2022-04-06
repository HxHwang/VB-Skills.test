VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB1"
   ClientHeight    =   3015
   ClientLeft      =   4590
   ClientTop       =   1800
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "‘›Õ£"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ø™ º"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " ‰»Î£∫"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 200
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Text1 = Text1 + 2
End Sub
