VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, j As Integer, n As Integer
n = Text1
For i = 1 To n
    For j = 1 To n
        Select Case j
            Case i
                Print 0 & Space(1);
            Case Else
                Print 1 & Space(1);
        End Select
    Next j
        Print ""
Next i
End Sub

Private Sub Command2_Click()
Cls
End Sub
