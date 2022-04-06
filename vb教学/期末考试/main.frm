VERSION 5.00
Begin VB.Form main 
   Caption         =   "期末考试"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9210
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu Mnukaoti 
      Caption         =   "考题"
      Begin VB.Menu Mnunum1 
         Caption         =   "第一题"
      End
      Begin VB.Menu Mnunum2 
         Caption         =   "第二题"
      End
      Begin VB.Menu Mnunum3 
         Caption         =   "第三题"
      End
      Begin VB.Menu Mnunum4 
         Caption         =   "第四题"
      End
      Begin VB.Menu Mnunum5 
         Caption         =   "第五题"
      End
   End
   Begin VB.Menu Mnuexit 
      Caption         =   "退出"
      Begin VB.Menu Mnutuichu 
         Caption         =   "退出程序"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mnunum1_Click()
Form1.Show
End Sub

Private Sub Mnunum2_Click()
Form2.Show
End Sub

Private Sub Mnunum3_Click()
Form3.Show
End Sub

Private Sub Mnunum4_Click()
Form4.Show
End Sub

Private Sub Mnunum5_Click()
Form6.Show
End Sub

Private Sub Mnutuichu_Click()
Dim answer As Integer
answer = MsgBox("确定要退出当前程序吗？", vbOKCancel + vbQuestion, "退出")
If answer = vbOK Then
   End
Else
  Exit Sub
End If
End Sub
