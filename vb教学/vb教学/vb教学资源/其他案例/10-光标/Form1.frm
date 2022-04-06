VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7935
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "确定"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2940
      ItemData        =   "Form1.frx":0000
      Left            =   4440
      List            =   "Form1.frx":0034
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2940
      ItemData        =   "Form1.frx":0078
      Left            =   720
      List            =   "Form1.frx":0088
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1935
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1875
      ScaleWidth      =   1275
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
Private Sub Command2_Click()
List1.Visible = True
List2.Visible = True
End Sub

Private Sub Command3_Click()

Select Case List1.Text
 Case Picture1.Name
  Picture1.MousePointer = List2.ListIndex
 Case Text1.Name
  Text1.MousePointer = List2.ListIndex
 Case Command1.Name
  Command1.MousePointer = List2.ListIndex
 Case Frame1.Name
  Frame1.MousePointer = List2.ListIndex
End Select
List1.Visible = False
List2.Visible = False

End Sub
