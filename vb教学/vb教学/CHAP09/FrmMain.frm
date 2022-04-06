VERSION 5.00
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "学生档案管理系统"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6405
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Menu System 
      Caption         =   "系统"
      Index           =   1
      Begin VB.Menu Add_User 
         Caption         =   "添加用户"
      End
      Begin VB.Menu Change_PWD 
         Caption         =   "修改密码"
      End
      Begin VB.Menu System_EXIT 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu Stuff_Info 
      Caption         =   "学生基本信息"
      Index           =   2
      Begin VB.Menu Add_Stuff 
         Caption         =   "添加学生信息"
      End
      Begin VB.Menu Change_Stuff 
         Caption         =   "修改学生信息"
      End
      Begin VB.Menu Check_Stuff 
         Caption         =   "查询学生信息"
      End
      Begin VB.Menu Del_Stuff 
         Caption         =   "删除学生信息"
      End
   End
   Begin VB.Menu Stuff_Checkin 
      Caption         =   "学生出勤信息"
      Index           =   3
      Begin VB.Menu Add_Checkin 
         Caption         =   "添加出勤信息"
         Begin VB.Menu AddAttendance 
            Caption         =   "添加上下学信息"
         End
         Begin VB.Menu AddOtherKQ 
            Caption         =   "添加其他出勤信息"
         End
      End
      Begin VB.Menu Change_Checkin 
         Caption         =   "修改出勤信息"
         Begin VB.Menu ChangeAttendance 
            Caption         =   "修改上下学信息"
         End
         Begin VB.Menu ChangeOtherKQ 
            Caption         =   "修改其他出勤信息"
         End
      End
      Begin VB.Menu Check_Checkin 
         Caption         =   "查询出勤信息"
      End
      Begin VB.Menu Del_Checkin 
         Caption         =   "删除出勤信息"
         Begin VB.Menu delInOut 
            Caption         =   "删除上下学信息"
         End
         Begin VB.Menu delOtherKQ 
            Caption         =   "删除其他出勤信息"
         End
      End
      Begin VB.Menu SetTime 
         Caption         =   "设置上下学时间"
      End
   End
   Begin VB.Menu Stuff_Alteration 
      Caption         =   "学生调动信息"
      Index           =   4
      Begin VB.Menu Add_Alter 
         Caption         =   "添加调动信息"
      End
      Begin VB.Menu Chage_Alter 
         Caption         =   "修改调动信息"
      End
      Begin VB.Menu Check_Alter 
         Caption         =   "查询调动信息"
      End
      Begin VB.Menu Del_Alter 
         Caption         =   "删除调动信息"
      End
   End
   Begin VB.Menu System_Help 
      Caption         =   "帮助"
      Index           =   6
      Begin VB.Menu About 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SQL As String

Private Sub About_Click()                         '关于窗体
    frmAbout.Show
    frmAbout.ZOrder 0
End Sub

Private Sub Add_Alter_Click()                     '添加调动信息
    flag = 1
    frmAlteration.Caption = "添加学生调动信息"
    frmAlteration.Show
    frmAlteration.ZOrder 0
End Sub

Private Sub Add_Stuff_Click()                      '添加员工信息
    flag = 1
    frmStuff_info.Show
    frmStuff_info.ZOrder 0
End Sub

Private Sub Add_User_Click()                       '添加用户
    Dim fAdd As New frmAddUser
    fAdd.Show
    fAdd.ZOrder 0
End Sub

Private Sub AddAttendance_Click()                  '添加上下学信息
    flag = 1
    FrmAttendance.Show
    FrmAttendance.ZOrder 0
End Sub

Private Sub AddOtherKQ_Click()                     '添加其他考勤信息
    flag = 1
    frmOtherKQ.Show
    frmOtherKQ.ZOrder 0
End Sub

Private Sub Chage_Alter_Click()                     '修改调动信息
    frmAlterationResult.Show
    frmAlterationResult.ZOrder 0
End Sub

Private Sub Change_PWD_Click()                      '修改密码
    Dim fChangePWD As New frmChangePWD
    fChangePWD.Show
End Sub

Private Sub Change_Stuff_Click()                    '修改员工信息
    frmCheckStuff.topic = "选择修改条件"
    frmCheckStuff.Caption = "修改学生基本信息"
    SQL = "select * from StuffInfo order by SID"
    frmResult.createList (SQL)
    frmResult.Show
    frmResult.ZOrder 0
    frmCheckStuff.Show
    frmCheckStuff.ZOrder 0
End Sub

Private Sub ChangeAttendance_Click()                 '修改上下学信息
    frmAResult.Show
    'frmAResult.ZOrder 0
End Sub

Private Sub changeOtherKQ_Click()                    '修改其他考勤信息
    frmOKQResult.Show
    frmOKQResult.ZOrder 0
End Sub

Private Sub Check_Alter_Click()                       '查询调动信息
    frmCheckAlter.Show
    frmCheckAlter.ZOrder 0
End Sub

Private Sub Check_Checkin_Click()                    '查询其他考勤信息
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub Check_Stuff_Click()                       '查询员工信息
    SQL = "select * from StuffInfo"
    frmResult.createList (SQL)
    frmResult.Show
    frmCheckStuff.Show
    frmResult.ZOrder 1
    frmCheckStuff.ZOrder 0
End Sub

Private Sub Del_Alter_Click()                          '删除调动信息
    frmAlterationResult.Show
    frmAlterationResult.ZOrder 0
End Sub

Private Sub Del_Stuff_Click()                          '删除员工信息
    frmCheckStuff.topic = "选择删除条件"
    frmCheckStuff.Caption = "删除学生基本信息"
    SQL = "select * from StuffInfo"
    frmResult.createList (SQL)
    frmResult.Show
    frmCheckStuff.Show
    frmResult.ZOrder 1
    frmCheckStuff.ZOrder 0
End Sub

Private Sub delInOut_Click()                            '删除上下学信息
    Dim SQL As String
    SQL = "select * from AttendanceInfo order by ID desc"
    Call frmAResult.ListTopic
    Call frmAResult.ShowData(SQL)
    frmAResult.Show
    frmAResult.ZOrder 0
End Sub

Private Sub delOtherKQ_Click()                            '删除其他考勤信息
    frmOKQResult.Show
    frmOKQResult.ZOrder 0
End Sub

Private Sub SetTime_Click()                               '设置上下学时间
    frmSetTime.Show
    frmSetTime.ZOrder 0
End Sub

Private Sub System_EXIT_Click()
    Unload Me
    Exit Sub
End Sub
