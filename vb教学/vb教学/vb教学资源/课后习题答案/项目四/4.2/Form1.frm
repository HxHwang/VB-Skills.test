VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7005
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示结果"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "请输入x的值"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "函 数 计 算 器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Long, y As Long

Private Sub Command1_Click()
x = Val(Text1.Text)
Select Case x
Case Is < 1
   y = x
Case 1 To 10
   y = 2 * x - 1
Case Else
   y = 3 * x - 11
End Select
Text2.Text = Str$(y)
End Sub

