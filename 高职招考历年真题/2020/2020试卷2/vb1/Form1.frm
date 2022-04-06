VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1110
      ItemData        =   "Form1.frx":0000
      Left            =   600
      List            =   "Form1.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "项目："
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then List1.RemoveItem (i)
Next i
End Sub
