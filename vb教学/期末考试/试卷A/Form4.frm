VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "个人简介表"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   4500
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "年龄"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   4335
      Begin VB.OptionButton Option3 
         Caption         =   "20"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "19"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "18"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "17"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "16"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "15"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "女"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   900
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "男"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "提交"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "兴趣爱好"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "购物"
         Height          =   375
         Index           =   8
         Left            =   3120
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "绘画"
         Height          =   375
         Index           =   7
         Left            =   1560
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "篮球"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "看书"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "上网"
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "演讲"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "游泳"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "朗诵"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "羽毛球"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "性别："
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "姓名："
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String, b As String, c As String, d As String

Private Sub Check1_Click(Index As Integer)
d = ""
For Index = 0 To 8
        If Check1(Index).Value = 1 Then
        d = d & Check1(Index).Caption & "  "
        End If
    Next Index
End Sub

Private Sub Command1_Click()
a = Text1.Text
Form4.Hide
Form5.Show
Form5.Label2.Caption = "姓名：" & a
Form5.Label3.Caption = "性别：" & b
Form5.Label4.Caption = "年龄：" & c
Form5.Label5.Caption = "兴趣爱好：" & d
End Sub

Private Sub Command2_Click()
Text1.Text = ""
'Combo1.Text = 16
Option1.Value = False
Option2.Value = False
For Index = 0 To 5
  Option3(Index).Value = False
Next Index
For Index = 0 To 8
  Check1(Index).Value = 0
Next Index
End Sub


Private Sub Option1_Click()
If Option1.Value = True Then
b = "男"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
b = "女"
End If
End Sub

Private Sub Option3_Click(Index As Integer)

        If Option3(Index).Value = True Then
        c = Option3(Index).Caption
        End If
   
End Sub

