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
   Begin VB.CheckBox Check1 
      Caption         =   "全选"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   900
      ItemData        =   "Form1.frx":0000
      Left            =   600
      List            =   "Form1.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "项目："
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i%
Private Sub Check1_Click()
If Check1.Value = 1 Then
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
Else
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End If
End Sub
Private Sub Command1_Click()
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then
List1.RemoveItem i
End If
Next
End Sub
