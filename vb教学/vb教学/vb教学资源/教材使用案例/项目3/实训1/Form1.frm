VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "大小写转换器"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5700
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "小写"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "大写"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S As String
Private Sub Command1_Click()
S = Text1.Text
Text1.Text = UCase(S)
End Sub

Private Sub Command2_Click()
S = Text1.Text
Text1.Text = LCase(S)
End Sub
