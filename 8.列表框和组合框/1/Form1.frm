VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7890
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "颜色"
      Height          =   2895
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   3360
         TabIndex        =   16
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   3360
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   3360
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.HScrollBar hsbblue 
         Height          =   255
         LargeChange     =   5
         Left            =   1320
         Max             =   255
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.HScrollBar hsbgreen 
         Height          =   255
         LargeChange     =   5
         Left            =   1320
         Max             =   255
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.HScrollBar hsbred 
         Height          =   255
         LargeChange     =   5
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2400
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "调色板颜色"
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "蓝"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "绿"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "红"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   720
      List            =   "Form1.frx":0019
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0047
      Left            =   720
      List            =   "Form1.frx":0057
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   960
      ItemData        =   "Form1.frx":0073
      Left            =   720
      List            =   "Form1.frx":0083
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "字号"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "字形"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "中文字体"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sst As String

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0
Text1.FontName = "宋体"
Case 1
Text1.FontName = "黑体"
Case 2
Text1.FontName = "楷体_GB2312"
Case 3
Text1.FontName = "仿宋_GB2312"
End Select
End Sub



Private Sub Combo2_Click()

Select Case Combo2.ListIndex
Case 0
Text1.FontSize = 26
Case 1
Text1.FontSize = 24
Case 2
Text1.FontSize = 22
Case 3
Text1.FontSize = 20
Case 4
Text1.FontSize = 18
Case 5
Text1.FontSize = 16
Case 6
Text1.FontSize = 14
End Select
End Sub

Private Sub Command1_Click()
sst = Text1.SelText
Text1.SelText = ""
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
sst = Text1.SelText
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
Text1.SelText = sst
Command1.Enabled = True
Command2.Enabled = True
End Sub



Private Sub hsbblue_Change()
Text1.ForeColor = RGB(hsbred, hsbgreen, hsbblue)
Label8.BackColor = RGB(hsbred, hsbgreen, hsbblue)
Text4.Text = hsbblue.Value
End Sub

Private Sub hsbgreen_Change()
Text1.ForeColor = RGB(hsbred, hsbgreen, hsbblue)
Label8.BackColor = RGB(hsbred, hsbgreen, hsbblue)
Text3.Text = hsbgreen.Value
End Sub

Private Sub hsbred_Change()
Text1.ForeColor = RGB(hsbred, hsbgreen, hsbblue)
Label8.BackColor = RGB(hsbred, hsbgreen, hsbblue)
Text2.Text = hsbred.Value
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
Case 0
Text1.FontBold = False
Text1.FontItalic = False
Case 1
Text1.FontBold = True
Text1.FontItalic = False
Case 2
Text1.FontBold = False
Text1.FontItalic = True
Case 3
Text1.FontBold = True
Text1.FontItalic = True
End Select
End Sub

