VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "输入日期"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Mid(Text1.Text, 6, 2) < 7 Then
        Print "这是上半年"
    Else
        Print "这是下半年"
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = "2018-04-01"
End Sub
