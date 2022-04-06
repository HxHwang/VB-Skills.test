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
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
      Begin VB.CheckBox Check2 
         Caption         =   "斜体"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "宋体"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "黑体"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Text            =   "新起点，新追求"
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text1.FontBold = True
    Else
        Text1.FontBold = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Text1.FontItalic = True
    Else
        Text1.FontItalic = False
    End If
End Sub

Private Sub Option1_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "宋体"
End Sub
