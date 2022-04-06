VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7995
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5040
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      Height          =   3510
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Menu mnuedit 
      Caption         =   "编辑"
      Begin VB.Menu Mnucopy 
         Caption         =   "复制"
      End
      Begin VB.Menu Mnucut 
         Caption         =   "剪切"
      End
      Begin VB.Menu Mnupaste 
         Caption         =   "粘贴"
      End
      Begin VB.Menu Mnudel 
         Caption         =   "删除"
      End
      Begin VB.Menu Mnuchakan 
         Caption         =   "查看"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As Integer

Dim copy As Boolean
Dim del As Boolean
Dim mfile As String
Dim wname As String
Private Sub Dir1_Change()
File1.Path = Dir1.Path
If Right(Dir1.Path, 1) = "\" Then
   mfile = Dir1.Path & wname
Else
   mfile = Dir1.Path & "\" & wname
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
wname = File1.FileName
If Right(Dir1.Path, 1) = "\" Then
   yfile = Dir1.Path & wname
Else
   yfile = Dir1.Path & "\" & wname
End If

End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
 Call Mnudel_Click
End If
End Sub

Private Sub Form_Load()
Drive1.Drive = "D"
copy = False
cut = False
End Sub

Private Sub Mnuchakan_Click()
Form2.Show
Form1.Hide
yfile = Dir1.Path & File1.FileName
End Sub

Private Sub Mnucopy_Click()
If yfile = "" Then
  MsgBox "未选中文件！", vbOKOnly + vbCritical, "警告"
Else
  copy = True
  del = False
End If
End Sub

Private Sub Mnucut_Click()
If yfile = "" Then
  MsgBox "未选中文件！", vbOKOnly + vbCritical, "警告"
Else
  copy = True
  del = True
End If
End Sub

Private Sub Mnudel_Click()
answer = MsgBox("确定要删除该文件吗？", vbOKCancel + vbCritical, "删除")
If answer = vbOK Then
    Kill yfile
    File1.Refresh
End If
End Sub

Private Sub Mnupaste_Click()
If copy = True Then
   FileCopy yfile, mfile
   File1.Refresh
End If
End Sub
