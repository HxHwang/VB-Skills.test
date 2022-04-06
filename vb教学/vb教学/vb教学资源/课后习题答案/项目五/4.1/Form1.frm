VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4005
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "确定"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1320
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0016
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   1260
      Left            =   2040
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub CmdCancel_Click()
    List1.AddItem (a)
    Combo1.RemoveItem (Combo1.ListCount - 1)
     CmdCancel.Enabled = False
End Sub

Private Sub CmdOk_Click()
    a = List1.Text
    Combo1.AddItem (a)
    List1.RemoveItem (List1.ListIndex)
    CmdCancel.Enabled = True
End Sub
