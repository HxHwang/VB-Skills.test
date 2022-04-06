VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6435
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2760
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Line (1000, 1000)-(1800, 1600), , B

End Sub

Private Sub Command2_Click()
    Form1.Line (1000, 1000)-(1800, 1000)
    Form1.Line (1800, 1000)-(1800, 1600)
    Form1.Line (1800, 1600)-(1000, 1600)
    Form1.Line (1000, 1600)-(1000, 1000)
End Sub

Private Sub Command3_Click()
    Shape1.Top = 1000
    Shape1.Left = 1000
    Shape1.Width = 800
    Shape1.Height = 600
    Shape1.Visible = True
End Sub
