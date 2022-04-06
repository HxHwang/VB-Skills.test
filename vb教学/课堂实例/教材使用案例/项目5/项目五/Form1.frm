VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5385
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox LstShow 
      Height          =   4560
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "查找"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "删除信息"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "录入信息"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ListBox LstAddr 
      Height          =   600
      ItemData        =   "Form1.frx":0000
      Left            =   960
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtNumb 
      Height          =   375
      Left            =   960
      MaxLength       =   7
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox CboMan 
      Height          =   300
      ItemData        =   "Form1.frx":0004
      Left            =   960
      List            =   "Form1.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "藉贯："
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "性别："
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "学号："
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sindex As Integer
Dim infr As String
Dim Sname As String, Snumb As String, man As String, addr As String

Private Sub Form_Load()
With LstAddr
.AddItem "北京"
.AddItem "上海"
.AddItem "湖北武汉"
.AddItem "湖南长沙"
.AddItem "四川成都"
.AddItem "广东广州"
End With
CboMan.ListIndex = 0
LstAddr.ListIndex = 0
i = 0
sindex = 0
End Sub

Private Sub cboMan_Click()
man = CboMan.Text
End Sub

Private Sub cmdDel_Click()
LstShow.RemoveItem sindex
i = i - 1
End Sub

Private Sub cmdInput_Click()
infr = Sname + "   " + man + "   " + addr
LstShow.AddItem infr, i
i = i + 1
End Sub

Private Sub cmdQuit_Click()
Unload Form1
End Sub

Private Sub lstAddr_Click()
addr = LstAddr.Text
End Sub

Private Sub lstShow_Click()
sindex = LstShow.ListIndex
End Sub

Private Sub txtName_Change()
Sname = TxtName.Text
End Sub

Private Sub txtNumb_Change()
Snumb = TxtNumb.Text
ls = Len(Sname)
Select Case ls
    Case 2
        Sname = Sname + "     "
    Case 3
        Sname = Sname + "    "
    Case 4
        Sname = Sname
End Select
End Sub

Private Sub txtNumb_LostFocus()
    If Len(Snumb) < 7 Then
        MsgBox "学号必须为7位数", vbOKOnly + vbCritical, "错误"
        TxtNumb.SetFocus
    End If
End Sub


Private Sub cmdFind_Click()
    Dim mystr As String
    Dim mybt As Integer
    Dim j As Integer
    Dim fs As Integer
    fs = 0
    '设置转支起点
step:
    mystr = InputBox("请输入所要查找学生学号", "查找对话框")
    If mystr = "" Then
      mybt = MsgBox("未输入学号，是否重新输入学号?", _
            vbOKCancel + vbQuestion, "确认输入")
        If mybt = 1 Then
            GoTo step
        Else
            Form1.Show
            Exit Sub
        End If
    End If
    For j = 0 To LstShow.ListCount - 1
      If mystr = Left(LstShow.List(j), 7) Then
         fs = 1
        Exit For
      End If
    Next
    If fs = 0 Then
       MsgBox "没有该学生的信息", vbOKOnly + vbCritical, _
       "错误"
    Else
       LstShow.ListIndex = j
    End If
End Sub

