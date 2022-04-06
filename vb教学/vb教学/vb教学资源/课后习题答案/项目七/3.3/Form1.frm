VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6705
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command1 
      Caption         =   "≤È’“"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Text1.Text
b = 10 * (a Mod 10) + a \ 10
For i = 1 To 99
j = 10 * (i Mod 10) + i \ 10
If a + i = b + j Then
 Print i
End If
Next i


End Sub
