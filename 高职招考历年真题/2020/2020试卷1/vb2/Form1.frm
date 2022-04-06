VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB2"
   ClientHeight    =   3015
   ClientLeft      =   8250
   ClientTop       =   1905
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "整数："
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a%
a = Val(Text1)
If a >= 0 And a <= 100 Then
MsgBox a & "为合理数", , "VB2"
Else
MsgBox a & "不是合理数", , "VB2"
End If

End Sub
