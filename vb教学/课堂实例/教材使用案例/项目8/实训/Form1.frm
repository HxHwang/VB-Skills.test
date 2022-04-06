VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "动态创建文件并向文件中输入内容"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4950
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "创建文件并输入数据"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Label Label1 
         Caption         =   "请输入保存文件的位置及文件名"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type stu
  stname As String * 10
  num As String
  age As Integer
  addr As String
  End Type
Private Sub Command1_Click()
   CommonDialog1.Filter = "txt(*.txt)|*.txt|doc(*.doc)|*.doc"     '保存文件类型
   CommonDialog1.ShowSave                                         '保存对话框
   Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Then                    '文件名不能为空
        MsgBox "文件名不能为空"
    Else
        Open Text1.Text For Output As #1    '在对应的位置新建文件
        MsgBox "创建文件成功，请按照提示输入学生信息！"
        Static stud() As stu                          '定义静态数组
        n = InputBox("请输入学生个数")             '利用输入函数输入数据
        ReDim stud(n) As stu
        For i = 1 To n
            stud(i).stname = InputBox("请输入姓名:")
            stud(i).num = InputBox("请输入年级:")
            stud(i).age = InputBox("请输入年龄:")
            stud(i).addr = InputBox("请输入地址:")
            Write #1, stud(i).stname, stud(i).num, stud(i).age, stud(i).addr
        Next i
        Close #1
        MsgBox "输入完毕！"
        End
    End If
End Sub

Private Sub Form_DblClick()
 End
End Sub

