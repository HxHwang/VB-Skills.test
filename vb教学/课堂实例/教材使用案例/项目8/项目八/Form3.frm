VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form3"
   ScaleHeight     =   4230
   ScaleWidth      =   3885
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdBack 
      Caption         =   "返回"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一记录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "上一记录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增成绩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox TxtScore 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TxtNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "成绩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "学号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   420
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type stu
  sNum  As String * 10
  sName As String * 10
  Score As String * 4
End Type

Dim gstu As stu
Dim recordlen As Integer
Dim currentrecord As Integer
Dim lastrecord As Integer
Public Sub ShowCurrent()
    '显示当前记录
    Get #1, currentrecord, gstu
    TxtNum.text = gstu.sNum
    TxtName.text = gstu.sName
    TxtScore.text = gstu.Score
End Sub

Public Sub SaveCurrent()
    '保存当前记录
    gstu.sNum = TxtNum.text
    gstu.sName = TxtName.text
    gstu.Score = TxtScore.text
    Put #1, currentrecord, gstu

End Sub

Private Sub Form_Load()
    recordlen = Len(gstu)
    If fName <> "" Then
        Open fName For Random As #1 Len = recordlen
        currentrecord = 1
        lastrecord = FileLen(fName) / recordlen
        If lastrecord = 0 Then
            lastrecord = 1
        End If
        ShowCurrent
    End If
End Sub

Private Sub cmdPrevious_Click()
    '如果当前记录已为第1个记录，也不能再显示
    If currentrecord = 1 Then
        Beep
        MsgBox "已到文件顶部！", vbOKOnly + vbExclamation, "错误"
    Else
    '如果当前不是第1个记录，则先保存当前记录
        '然后再显示当前记录
        SaveCurrent
        '将当前记录移到上一个记录
        currentrecord = currentrecord - 1
        '显示当前记录
        ShowCurrent
    End If
    TxtNum.SetFocus
End Sub


Private Sub cmdNext_Click()
    '如果当记录为最后的记录，则不能再显示
    If currentrecord = lastrecord Then
        Beep
        MsgBox "已显示完全部成绩！", vbOKOnly + vbExclamation, "错误"
    Else
    '如果当前记录不是最后记录，则先保存当前记录
        '然后再显示当前记录
        SaveCurrent
        '当前记录移到下一个记录
        currentrecord = currentrecord + 1
        '显示当前记录
        ShowCurrent
    End If
    TxtNum.SetFocus

End Sub

Private Sub cmdAdd_Click()
    '将所输入的记录保存到文件的最后记录
    SaveCurrent
    '在文件的最后增加1个空白记录，并保存
    lastrecord = lastrecord + 1
    currentrecord = lastrecord
    '保存后，将文本框中的内容清除
    TxtNum.text = ""
    TxtName.text = ""
    TxtScore.text = ""
    TxtNum.SetFocus

End Sub


Private Sub cmdFind_Click()
    Dim nsearch As String
    Dim found As Boolean
    Dim recnum As Long
    Dim fstu As stu
    '输入要查找的学生的学号
    nsearch = InputBox("请输入要查找的学生的学号：", "查找")
    If nsearch = "" Then
        Exit Sub
    End If
    found = False
    '从文件的第一个记录开始找起
    '直到找到某个记录中的学号字段和所输入的学号一致为止
    For recnum = 1 To lastrecord
        Get #1, recnum, fstu
        If nsearch = Trim(fstu.sNum) Then
            found = True
            Exit For
        End If
    Next
    '如果找到了，就显示该记录
    If found = True Then
        SaveCurrent
        currentrecord = recnum
        ShowCurrent
    '否则提示用户未找到该学生
    Else
        MsgBox "无学号为" + nsearch + "的成绩"
    End If
End Sub

Private Sub cmdback_Click()
    Unload Form3
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #1
End Sub
