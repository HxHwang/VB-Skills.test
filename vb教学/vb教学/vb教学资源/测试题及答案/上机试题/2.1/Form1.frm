VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   7980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   855
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   5100
      ItemData        =   "Form1.frx":0000
      Left            =   4800
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   5100
      ItemData        =   "Form1.frx":0004
      Left            =   360
      List            =   "Form1.frx":0017
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex

End Sub

Private Sub Command2_Click()
List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command3_Click()
For i = List1.ListCount - 1 To 0 Step -1
List2.AddItem List1.List(i)
List1.RemoveItem i
Next i

  
End Sub

Private Sub Command4_Click()
For i = List2.ListCount - 1 To 0 Step -1
List1.AddItem List2.List(i)
List2.RemoveItem i
Next i
End Sub
