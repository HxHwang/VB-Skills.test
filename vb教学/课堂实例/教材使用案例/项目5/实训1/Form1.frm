VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   3225
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'向组合框中添加列表项
With Combo1
.AddItem "北京"
.AddItem "上海"
.AddItem "湖北武汉"
.AddItem "湖南长沙"
.AddItem "四川成都"
.AddItem "广东广州"
.ListIndex = 0
End With
End Sub


Private Sub Combo1_Click()
'在文本框中分行显示所选中的列表项
Text1.Text = Text1.Text + Combo1.Text + Chr(13) + Chr(10)
End Sub

