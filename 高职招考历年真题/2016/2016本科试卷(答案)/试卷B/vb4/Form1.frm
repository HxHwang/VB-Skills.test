VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   8760
   ClientTop       =   5220
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "¡„«… ˝"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim i As Integer
For i = 1000 To 9999
g = i Mod 10
s = i \ 10 Mod 10
b = i \ 100 Mod 10
q = i \ 1000
If b = 0 And (q * 100 + s * 10 + g) * 9 = i Then Print i
Next i

End Sub

