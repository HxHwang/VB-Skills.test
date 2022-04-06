VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3990
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdInput 
      Caption         =   "输入学生信息"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdInput_Click()
'显示"学生基本信息"对话框
Form2.Show 0
End Sub

Private Sub Form_Load()
'在窗体上显示"姓名    性别  入学时间      专业"
Form1.Print "姓名    性别  入学时间      专业"
End Sub

