VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6570
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "字型"
      Height          =   2175
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "斜体"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "粗体"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "大小"
      Height          =   2175
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
      Begin VB.OptionButton Option6 
         Caption         =   "24点"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "20点"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "16点"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
      Begin VB.OptionButton Option3 
         Caption         =   "楷体"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "黑体"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宋体"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5415
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

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Text1.FontUnderline = True
Else
  Text1.FontUnderline = False
End If
End Sub

Private Sub Option1_Click()

Text1.FontName = "宋体"


End Sub

Private Sub Option2_Click()
Text1.FontName = "黑体"
End Sub

Private Sub Option3_Click()
Text1.FontName = "楷体_GB2312"
End Sub

Private Sub Option4_Click()
Text1.FontSize = 16
End Sub

Private Sub Option5_Click()
Text1.FontSize = 20
End Sub

Private Sub Option6_Click()
Text1.FontSize = 24
End Sub
