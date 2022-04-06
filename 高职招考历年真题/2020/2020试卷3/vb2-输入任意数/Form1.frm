VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB2"
   ClientHeight    =   3015
   ClientLeft      =   4125
   ClientTop       =   1890
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "输入："
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, k, f As Integer

n = Val(Text1.Text)
k = Len(n)
Print k
f = 0 '是
For i = 1 To k
If Mid(n, i, 1) > 7 Then
f = 1 '不是8进制
Exit For
End If
Next i
If f = 0 Then
MsgBox n & "是八进制数", , "VB2"
Else
MsgBox n & "不是八进制数", , "VB2"
End If
End Sub
