VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6210
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cmd2 
      Caption         =   "存盘"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "计算"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton Opt2 
      Caption         =   "200-400之间素数"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "100-200之间素数"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd1_Click()

Dim i As Integer
Dim temp As Long
temp = 0

If Opt2.Value Then
   For i = 200 To 400
       If isprime(i) Then
          temp = temp + i
       End If
   Next
Else
   For i = 100 To 200
      If isprime(i) Then
         temp = temp + i
      End If
   Next
End If
Text1.Text = temp


End Sub

Private Sub Cmd2_Click()
putdata "\out.txt", Text1.Text
End Sub
