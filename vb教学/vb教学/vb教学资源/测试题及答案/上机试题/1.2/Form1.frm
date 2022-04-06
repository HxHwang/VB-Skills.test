VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6465
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.HScrollBar HS1 
      Height          =   855
      Left            =   480
      Max             =   2000
      Min             =   1000
      TabIndex        =   1
      Top             =   1080
      Value           =   1000
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2520
      Width           =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HS1_Change()
Text1.Width = HS1.Value
End Sub
