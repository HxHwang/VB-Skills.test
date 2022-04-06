VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5355
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   3330
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   3450
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "编辑"
      Begin VB.Menu MnuCut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheck 
         Caption         =   "查看"
      End
      Begin VB.Menu MnuModi 
         Caption         =   "修改"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sourfile As String '用于保存源文件
Dim DestFile As String '用于保存目标文件
Dim SureCopy As Integer '用于控制是否单击复制或剪切菜单
Dim SureDell As Boolean '用于控制是否单击删除文件
Dim sfn As String '用于保存被选中文件的文件名

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub Form_Load()
    Drive1.Drive = "F"
    SureCopy = 0
    SureDell = False
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    '如果已经选择复制或剪切命令，则将当前路径作为目标路径
    If SureCopy = 1 Then
        If Right(Dir1.Path, 1) <> "\" Then
            DestFile = Dir1.Path + "\" + sfn
        Else
            DestFile = Dir1.Path + sfn
        End If
     '如果没有，将当前路径作为源路径
    Else
        If Right(Dir1.Path, 1) <> "\" Then
            SourPath = Dir1.Path + "\"
        Else
            SourPath = Dir1.Path
        End If
    End If
End Sub

Private Sub File1_DblClick()

    '得到编辑文件的详细路径
    fName = Dir1.Path + "\" + File1.FileName
    '打开文件编辑查看窗体
    Form2.Show
    Form1.Hide

End Sub

Private Sub File1_Click()
    sfn = File1.FileName
    Form1.Caption = Dir1.Path + "\" + File1.FileName
    '选中源文件
    If Right(Dir1.Path, 1) <> "\" Then
        Sourfile = Dir1.Path + "\" + File1.FileName
    Else
        Sourfile = Dir1.Path + File1.FileName
    End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
    '选中文件后，如果按下"D"键，则询问是否删除文件
    If KeyAscii = 100 Then
        SureDel = MsgBox("确实要删除文件吗？", vbYesNo + vbQuestion, "删除文件")
        '如果单击"是"按钮，则删除选中文件；如果单击"否"按钮，则不删除文件
        Select Case SureDel
        Case vbYes
            '删除文件
            Kill (Sourfile)
            '更新文件列表
            File1.Refresh
        Case vbNo
            Exit Sub
        End Select
    End If
End Sub

Private Sub MnuCheck_Click()
    '得到编辑文件的详细路径
    fName = Dir1.Path + "\" + File1.FileName
    '打开文件编辑查看窗体
    Form3.Show
    Form1.Hide
End Sub

Private Sub mnuCopy_Click()
    '选择复制命令后，以后路径将作为目标路径
    If sfn = "" Then
        MsgBox "未选中文件", vbOKOnly + vbCritical, "错误"
        SureCopy = 0
     Else
        SureCopy = 1
        SureDell = False
    End If
End Sub

Private Sub mnuCut_Click()
    '选择剪切命令后，以后路径将作为目标路径，同时删除被选中的文件
    If sfn = "" Then
        MsgBox "未选中文件", vbOKOnly + vbCritical, "错误"
         SureCopy = 0
        SureDell = False
    Else
        SureCopy = 1
        SureDell = True
    End If

End Sub

Private Sub MnuModi_Click()
    '得到编辑文件的详细路径
    fName = Dir1.Path + "\" + File1.FileName
    '打开文件编辑查看窗体
    Form2.Show
    Form1.Hide
End Sub

Private Sub mnuPaste_Click()
    '如果文件名已经存在，则询问是否覆盖文件
    If SureCopy = 1 Then
       If Dir(DestFile) <> "" Then
            intfile = MsgBox("文件" + DestFile + "已经存在，是否覆盖？", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "覆盖文件")
            Select Case intfile
            '覆盖文件
              Case vbYes
                FileCopy Sourfile, DestFile
              Case vbNo
                Exit Sub
              End Select
        Else
        '复制文件
         FileCopy Sourfile, DestFile
        End If
    End If
    '如果选择的是剪切命令，则删除源文件
    If SureDell = True Then
        Kill (Sourfile)
    End If
    File1.Refresh
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '在文件列表框中单击鼠标右键，弹出【编辑】菜单的快捷菜单
    If Button = 2 Then
        PopupMenu MnuEdit
    End If
End Sub

