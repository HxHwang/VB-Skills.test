VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求和"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i%, sum%
sum = 0
For i = 10 To 99
If i Mod 5 = 0 And i Mod 7 = 0 Then
Print i;
sum = sum + i
End If
Next

Print
Print "两位数中能被5和7整除的数之和:" & sum

End Sub
