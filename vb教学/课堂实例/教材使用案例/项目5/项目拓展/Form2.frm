VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "学生基本信息"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   4485
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbYear 
      Height          =   300
      Left            =   3240
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "性别"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optFemale 
         Caption         =   "女"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optMale 
         Caption         =   "男"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox cmbDepart 
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "张晓民"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "入学时间"
      Height          =   180
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "专业"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'为"专业"组合框添加项目，并设置默认项
cmbDepart.AddItem "计算机"
cmbDepart.AddItem "会计"
cmbDepart.AddItem "市场营销"
cmbDepart.AddItem "管理"
cmbDepart.ListIndex = 0
'为"入学时间"组合框添加项目，并设置默认项
cmbYear.AddItem "2001年9月"
cmbYear.AddItem "2002年9月"
cmbYear.AddItem "2003年9月"
cmbYear.ListIndex = 2
'设置默认性别
optMale.Value = True
End Sub

Private Sub OKButton_Click()
'定义一个用于存储性别的字符串
Dim man As String
'根据所选的性别，将性别赋给所定义的字符串
If optMale Then
man = "男"
Else
man = "女"
End If
'将所输入的学生基本信息显示到主窗体上
Form1.Print txtName.Text + "   " + man + "    " + cmbYear.Text + _
"    " + cmbDepart.Text
'隐藏对话框
Form2.Hide
End Sub


