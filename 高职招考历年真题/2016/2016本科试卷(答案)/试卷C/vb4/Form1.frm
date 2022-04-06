VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   8775
   ClientTop       =   4905
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "ÇóÊ½Êý"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim i As Single
For i = 1000 To 9999
g = i Mod 10
s = i \ 10 Mod 10
b = i \ 100 Mod 10
q = i \ 1000
If i * 4 = g * 1000 + s * 100 + b * 10 + q Then Print i
Next i
End Sub
