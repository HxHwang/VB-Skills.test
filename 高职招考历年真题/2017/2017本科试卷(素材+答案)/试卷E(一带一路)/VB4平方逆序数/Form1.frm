VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "平方逆序数"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 10 To 100
j = Val(StrReverse(i))
If i < j Then
    If i ^ 2 = StrReverse(j ^ 2) Then
        Print i & "^2 =" & i ^ 2 & Space(5) & j & "^2 =" & j ^ 2 & Space(5)
    End If
End If
Next
End Sub

