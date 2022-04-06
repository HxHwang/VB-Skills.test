VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "学科选修"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5145
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2850
      Width           =   420
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   300
      Left            =   2295
      TabIndex        =   4
      Top             =   2400
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   315
      Left            =   2310
      TabIndex        =   3
      Top             =   1935
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   300
      Left            =   2325
      TabIndex        =   2
      Top             =   1485
      Width           =   420
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2850
      TabIndex        =   1
      Top             =   1185
      Width           =   1845
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   225
      TabIndex        =   0
      Top             =   1185
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   "已选学科："
      Height          =   255
      Left            =   2835
      TabIndex        =   8
      Top             =   855
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "所有学科："
      Height          =   225
      Left            =   225
      TabIndex        =   7
      Top             =   870
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "学科选修"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   165
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    Dim m As Integer, n As Integer
    m = List1.ListCount
    Rem 利用循环将列表框List1的列表项全部移动到List2
    For n = 0 To m - 1
        List2.AddItem (List1.List(0))   '移动列表框顶端列表项
        List1.RemoveItem (0)
    Next n
    List2.Selected(0) = True            '设置被选择列表项
End Sub

Private Sub Command4_Click()
    Dim m As Integer, n As Integer
    m = List2.ListCount
    Rem 利用循环将列表框List2的列表项全部移动到List1
    For n = 0 To m - 1
        List1.AddItem (List2.List(0))   '移动列表框顶端列表项
        List2.RemoveItem (0)
    Next n
    List1.Selected(0) = True            '设置被选择列表项
End Sub

Private Sub Form_Load()
    Rem 初始化列表框List1
    List1.AddItem ("语文")
    List1.AddItem ("高数")
    List1.AddItem ("英语")
    List1.AddItem ("计算机基础")
    List1.AddItem ("计算机网络")
    List1.AddItem ("图形图像")
    List1.AddItem ("多媒体")
    List1.AddItem ("电子基础")
    List1.AddItem ("C程序设计")
    List1.AddItem ("C++程序设计")
    List1.AddItem ("VB程序设计")
    List1.AddItem ("数据库基础")
    List1.AddItem ("数据结构")
    List1.AddItem ("会计原理")
    List1.AddItem ("马列")
    List1.AddItem ("邓选")
    List1.Selected(0) = True                '设置被选择项
End Sub


Private Sub List1_DblClick()                '当列表框内某列表项被双击时
    Dim n As Integer
    n = List1.ListIndex                     '记录当前列表项索引值
    If List1.ListCount > 0 And n >= 0 Then
        List2.AddItem (List1.Text)          '将当前列表项添加到List2
        List1.RemoveItem (n)                '从列表框中删除当前列表项
        If List1.ListCount > n Then List1.ListIndex = n    '重设被选择的列表项
    End If
End Sub
Private Sub Command1_Click()                '当“>”键被单击时
    Dim n As Integer
    n = List1.ListIndex                     '记录当前列表项索引值
    If List1.ListCount > 0 And n >= 0 Then
        List2.AddItem (List1.Text)          '将当前列表项添加到List2
        List1.RemoveItem (n)                '从列表框中删除当前列表项
        If List1.ListCount > n Then List1.ListIndex = n    '重设被选择的列表项
    End If
End Sub
Private Sub List2_DblClick()                '当列表框内某列表项被双击时
    Dim n As Integer
    n = List2.ListIndex                     '记录当前列表项索引值
    If List1.ListCount > 0 And n >= 0 Then
        List1.AddItem (List2.Text)          '将当前列表项添加到List1
        List2.RemoveItem (n)                '从列表框中删除当前列表项
        If List2.ListCount > n Then List2.ListIndex = n    '重设被选择的列表项
    End If
End Sub
Private Sub Command3_Click()                '当“>”键被单击时
    Dim n As Integer
    n = List2.ListIndex                     '记录当前列表项索引值
    If List1.ListCount > 0 And n >= 0 Then
        List1.AddItem (List2.Text)          '将当前列表项添加到List1
        List2.RemoveItem (n)                '从列表框中删除当前列表项
        If List2.ListCount > n Then List2.ListIndex = n   '重设被选择的列表项
    End If
End Sub
