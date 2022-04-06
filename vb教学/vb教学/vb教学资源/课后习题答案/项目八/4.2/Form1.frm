VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "动态读取文件中的内容"
   ClientHeight    =   4605
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtText 
      BackColor       =   &H00C0FFFF&
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu MunFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu Munopen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu MunSave 
         Caption         =   "保存"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu Munquit 
         Caption         =   "退出(&Q)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fName As String
Dim textbuff As String
    
Sub EditOpen()
    If fName <> "" Then
        '打开顺序文件
        Open fName For Input As #1
        '读取顺序文件中的内容，并将它显示到文本框中
        Do While Not EOF(1)
            Line Input #1, text
            textbuff = textbuff + text + Chr(13) + Chr(10)
            TxtText.text = textbuff
        Loop
        Close #1
    End If
End Sub



Private Sub Munopen_Click()
    Dim text As String

    '显示"打开"对话框
    CommonDialog1.ShowOpen
    fName = CommonDialog1.FileName
    Call EditOpen
End Sub

Private Sub Munquit_Click()
 End
End Sub

Private Sub MunSave_Click()
    Dim fName As String
    Dim text As String

    '显示"另存为"对话框
    CommonDialog1.ShowSave
    fName = CommonDialog1.FileName
    If fName <> "" Then
         '打开顺序文件
         Open fName For Output As #1
         '将文本框中的内容写入文件
         Print #1, TxtText.text
         '关闭文件
         Close #1
    End If
End Sub
