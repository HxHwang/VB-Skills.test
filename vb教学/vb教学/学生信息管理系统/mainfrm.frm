VERSION 5.00
Begin VB.Form mainfrm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "学生信息管理系统"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8550
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu filemenu 
      Caption         =   "文件"
      Begin VB.Menu addusermenu 
         Caption         =   "添加用户"
      End
      Begin VB.Menu exitmenu 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu xuejimenu 
      Caption         =   "学籍管理"
      Begin VB.Menu personinfomenu 
         Caption         =   "查看学生信息"
      End
      Begin VB.Menu addinfomenu 
         Caption         =   "添加学生信息"
      End
      Begin VB.Menu editinfomenu 
         Caption         =   "修改学生信息"
      End
      Begin VB.Menu findinfomenu 
         Caption         =   "查找学生信息"
      End
   End
   Begin VB.Menu chengjimenu 
      Caption         =   "成绩管理"
      Begin VB.Menu addscoremenu 
         Caption         =   "添加成绩"
      End
      Begin VB.Menu editmenu 
         Caption         =   "修改成绩"
      End
      Begin VB.Menu findscoremenu 
         Caption         =   "查询成绩"
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub addinfomenu_Click()
addstudentfrm.Show
mainfrm.Hide
End Sub


Private Sub addusermenu_Click()
adduser.Show
mainfrm.Hide
End Sub

Private Sub editinfomenu_Click()
editfrm.Show
mainfrm.Hide
End Sub

Private Sub exitmenu_Click()
Dim ans As String
ans = MsgBox("确定退出系统？", vbYesNo + vbInformation, "退出")
If ans = vbYes Then
   End
End If
End Sub

Private Sub findinfomenu_Click()
findfrm.Show
mainfrm.Hide
End Sub

Private Sub personinfomenu_Click()
infofrm.Show
mainfrm.Hide

End Sub
