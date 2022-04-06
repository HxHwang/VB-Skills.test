VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "选课"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "Form1.frx":0004
      Left            =   360
      List            =   "Form1.frx":0020
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "已选课程列表"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "课程列表"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim flag As Boolean
flag = True
For i = 0 To List2.ListCount - 1
   If List1.Text = List2.List(i) Then
        flag = False
   End If
Next i
If flag Then
    List2.AddItem List1.Text
Else
    MsgBox "已选过这门课程!"
End If
End Sub
