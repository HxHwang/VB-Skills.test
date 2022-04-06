VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB1"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   1440
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":000D
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "项目："
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.ListIndex = -1 Then
MsgBox "请先选择要删除的项", , "提示选择"
Else
Combo1.RemoveItem Combo1.ListIndex
End If

End Sub
