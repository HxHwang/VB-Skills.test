VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6195
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "关闭"
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加法"
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "和　数"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "加　数"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "被加数"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub Command3_Click()
End

End Sub

