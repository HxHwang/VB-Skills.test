VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "第5个字符"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "字符长度"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String

Private Sub Text1_Change()
s = Text1
End Sub

Private Sub Command1_Click()
Text2 = Len(s)
End Sub

Private Sub Command2_Click()
Text2 = Mid(s, 5, 1)
End Sub
