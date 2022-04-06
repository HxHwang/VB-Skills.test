VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10860
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "ASCIIÂë"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ð¡Ð´×Ö·û"
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = LCase(Text1.Text)
End Sub

Private Sub Command2_Click()
Text2.Text = Asc(Text1.Text)
End Sub
