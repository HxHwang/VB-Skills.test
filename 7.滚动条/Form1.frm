VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   7065
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "字体颜色"
      Height          =   2295
      Left            =   840
      TabIndex        =   13
      Top             =   4200
      Width           =   5415
      Begin VB.TextBox Text5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   4080
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   4080
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4080
         TabIndex        =   17
         Top             =   345
         Width           =   615
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   720
         Max             =   255
         TabIndex        =   16
         Top             =   1320
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   720
         Max             =   255
         TabIndex        =   15
         Top             =   840
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   720
         Max             =   255
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "调色板颜色"
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "G"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "R"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
   End
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

Private Sub HScroll1_Change()
Text2.Text = HScroll1.Value
Text1.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Text3.Text = HScroll2.Value
Text1.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Text4.Text = HScroll3.Value
Text1.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
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

