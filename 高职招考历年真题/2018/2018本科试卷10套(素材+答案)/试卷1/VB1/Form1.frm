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
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "输入字符串"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Len(Text1.Text) >= 5 Then
        Print "这是一个长字符串"
    Else
        Print "这是一个短字符串"
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
End Sub
