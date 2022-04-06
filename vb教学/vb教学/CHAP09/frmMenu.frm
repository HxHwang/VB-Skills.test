VERSION 5.00
Begin VB.Form popmenu 
   Caption         =   "菜单"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Menu popmenu 
      Caption         =   "popmenu"
      Visible         =   0   'False
      Begin VB.Menu add 
         Caption         =   "添加学生基本信息"
      End
      Begin VB.Menu Change 
         Caption         =   "修改学生基本信息"
      End
      Begin VB.Menu Check 
         Caption         =   "查询学生基本信息"
      End
      Begin VB.Menu Del 
         Caption         =   "删除学生基本信息"
      End
   End
   Begin VB.Menu popmenu1 
      Caption         =   "popmenu1"
      Visible         =   0   'False
      Begin VB.Menu addInOut 
         Caption         =   "添加上下学信息"
      End
      Begin VB.Menu checkKQ1 
         Caption         =   "查询考勤信息"
      End
      Begin VB.Menu delKQ1 
         Caption         =   "删除上下学信息"
      End
   End
   Begin VB.Menu popmenu2 
      Caption         =   "popmenu2"
      Visible         =   0   'False
      Begin VB.Menu addOtherKQ 
         Caption         =   "添加其他考勤信息"
      End
      Begin VB.Menu checkKQ2 
         Caption         =   "查询考勤信息"
      End
      Begin VB.Menu delKQ2 
         Caption         =   "删除其他考勤信息"
      End
   End
   Begin VB.Menu popmenu3 
      Caption         =   "popmenu3"
      Visible         =   0   'False
      Begin VB.Menu addAlteration 
         Caption         =   "添加调动信息"
      End
      Begin VB.Menu ChangeAlter 
         Caption         =   "修改调动信息"
      End
      Begin VB.Menu Checkalter 
         Caption         =   "查询调动信息"
      End
      Begin VB.Menu DelAlter 
         Caption         =   "删除调动信息"
      End
   End
End
Attribute VB_Name = "popmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public str1 As String

Private Sub add_Click()                             '添加学生信息
    flag = 1
    frmStuff_info.Show vbModal
    frmStuff_info.ZOrder 0
    frmResult.ZOrder 1
End Sub

Private Sub addAlteration_Click()                   '添加调动信息
    flag = 1
    frmAlteration.Show
    frmAlteration.ZOrder 0
End Sub

Private Sub addInOut_Click()                        '添加上下学信息
    flag = 1
    FrmAttendance.Show
    FrmAttendance.ZOrder 0
End Sub

Private Sub AddOtherKQ_Click()                       '添加其他考勤信息
    flag = 1
    frmOtherKQ.Show
    frmOtherKQ.ZOrder 0
End Sub

Private Sub change_Click()  '修改学生信息
    flag = 2
    If frmResult.rsGrid.Rows > 1 Then
    
        gSQL = "select * from StuffInfo where SID='" & Trim(frmResult.rsGrid.TextMatrix( _
                frmResult.rsGrid.Row, 0)) & "'"
        frmStuff_info.Show
        frmStuff_info.ZOrder 0
    Else
     MsgBox "目前没有学生信息,请先添加学生信息！", vbOKOnly + vbExclamation, "警告!"
     flag = 1
     frmStuff_info.Show
    End If
End Sub

Private Sub aa_Click()                           '修改其他考勤信息
    flag = 2
    If frmOKQResult.LRecordList.Rows > 1 Then
        kqsql2 = "select * from StuffInfo where SID='" & Trim(frmOKQResult.LRecordList.TextMatrix( _
                frmOKQResult.LRecordList.Row, 0)) & "'"
        frmOtherKQ.Show
        frmOtherKQ.ZOrder 0
    Else
     MsgBox "目前没有学生信息,请先添加学生信息！", vbOKOnly + vbExclamation, "警告!"
     flag = 1
     frmOtherKQ.Show
    End If
End Sub


Private Sub tt_Click()                           '修改上下学信息
    flag = 2
    'FrmAttendance.Caption = "修改学生上下学信息"
    If frmAResult.recordlist.Rows > 1 Then
        kqsql = "select * from AttendanceInfo where ID='" & Trim(frmAResult.recordlist.TextMatrix( _
                frmAResult.recordlist.Row, 0)) & "'"
        FrmAttendance.Show
        FrmAttendance.ZOrder 0
    Else
     MsgBox "目前没有上下学信息,请先添加信息！", vbOKOnly + vbExclamation, "警告!"
     flag = 1
     FrmAttendance.Show
    End If
End Sub
Private Sub ChangeAlter_Click()                      '修改调动信息
    Dim rs As New ADODB.Recordset
    flag = 2
    frmAlteration.Caption = "修改学生调动信息"
    If frmAlterationResult.DataGrid1.Row < 0 Then
        MsgBox "目前没有记录！", vbOKOnly + vbExclamation, "提示！"
        flag = 1
        frmAlteration.Show
        frmAlteration.ZOrder 0
    Else
       str1 = "select * from AlterationInfo where ID=" & Trim( _
            frmAlterationResult.DataGrid1.Columns(0))
        frmAlteration.ID = Trim(frmAlterationResult.DataGrid1.Columns(0))
        Set rs = TransactSQL(str1)
        If rs.EOF = False Then
        With rs
            frmAlteration.AID = rs(1)
            frmAlteration.AName = rs(2)
            frmAlteration.AOldDept = rs(3)
            frmAlteration.ANewDept = rs(4)
            frmAlteration.AOldPosition = rs(5)
            frmAlteration.ANewPosition = rs(6)
            frmAlteration.AOutTime = rs(7)
            frmAlteration.AInTime = rs(8)
            frmAlteration.ARemark = rs(9)
        End With
            rs.Close
        End If
        frmAlteration.Show
        frmAlteration.ZOrder 0
    End If
End Sub

Private Sub check_Click()                            '查询学生信息
    frmCheckStuff.Show
End Sub

Private Sub checkKQ1_Click()                         '查询考勤信息
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub checkKQ2_Click()                         '查询考勤信息
    frmCheckKQ.Show
    frmCheckKQ.ZOrder 0
End Sub

Private Sub del_Click()                              '删除学生信息
    Dim SQL As String
    If frmResult.rsGrid.Rows = 1 Then
        MsgBox "目前没有学生信息,请先添加学生信息！", vbOKOnly + vbExclamation, "警告!"
        flag = 1
        frmStuff_info.Show
        frmStuff_info.ZOrder 0
    Else
        SQL = "delete from StuffInfo where SID='" & Trim(frmResult.rsGrid.TextMatrix( _
                frmResult.rsGrid.Row, 0)) & "'"
        If MsgBox("真的要删除这条记录么？", vbOKCancel + vbExclamation, "提示！") = vbOK _
        Then
        TransactSQL (SQL)
        MsgBox "学生信息记录已经删除！", vbOKOnly + vbExclamation, "警告!"
        Unload Me
        SQL = "select * from StuffInfo"
       
        frmResult.createList (SQL)
        Unload frmResult
        frmResult.Show
        End If
    End If
End Sub

Private Sub delKQ1_Click()                           '删除上下学信息
     Dim SQL As String
    If frmAResult.recordlist.Rows = 1 Then
        MsgBox "目前没有上下学信息！", vbOKOnly + vbExclamation, "警告！"
        flag = 1
        FrmAttendance.Show
        FrmAttendance.ZOrder 0
    Else
            SQL = "delete from AttendanceInfo where ID=" & Trim(frmAResult.recordlist.TextMatrix( _
                frmAResult.recordlist.Row, 0))
        If MsgBox("真的要删除这条记录么？", vbOKCancel + vbExclamation, "提示！") = vbOK _
        Then
        TransactSQL (SQL)
        MsgBox "记录已经删除！", vbOKOnly + vbExclamation, "警告!"
        Unload Me
        SQL = "select * from AttendanceInfo"
        'frmAResult.ListTopic
        frmAResult.ShowData (SQL)
        
        Unload frmAResult
        frmAResult.Show
         Else
         Unload frmAResult
         End If
    End If
    
End Sub

Private Sub DelAlter_Click()                         '删除调动信息
    Dim SQL As String
    If frmAlterationResult.DataGrid1.Row < 0 Then
        MsgBox "目前没有记录！", vbOKOnly + vbExclamation, "提示！"
        flag = 1
        frmAlteration.Show
        frmAlteration.ZOrder 0
    Else
        SQL = "delete from AlterationInfo where ID=" & Trim( _
            frmAlterationResult.DataGrid1.Columns(0).CellText( _
            frmAlterationResult.DataGrid1.Bookmark))
        If MsgBox("真的要删除这条记录么？", vbOKCancel) = vbOK Then
            TransactSQL (SQL)
            MsgBox "记录已经删除！", vbOKOnly + vbExclamation, "提示！"
            SQL = "select * from AlterationInfo order by ID"
            frmAlterationResult.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Person.mdb"
            frmAlterationResult.Adodc1.RecordSource = SQL
            If SQL <> "" Then
                frmAlterationResult.Adodc1.Refresh
            End If
            Set frmAlterationResult.DataGrid1.DataSource = frmAlterationResult.Adodc1.Recordset
            frmAlterationResult.DataGrid1.Refresh
            frmAlterationResult.Show
            frmAlterationResult.ZOrder 0
        End If
    End If
End Sub

Private Sub delKQ2_Click()                            '删除其他考勤信息
    Dim SQL As String
    Select Case frmOKQResult.SSTab.Caption
    Case "学生请假信息列表"
        If frmOKQResult.LRecordList.Rows = 1 Then
            MsgBox "目前没有请假信息！", vbOKOnly + vbExclamation, "警告！"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from LeaveInfo where LID="
            SQL = SQL & Trim(frmOKQResult.LRecordList.TextMatrix( _
                                    frmOKQResult.LRecordList.Row, 0))
            If MsgBox("真的要删除这条记录么？", vbOKCancel + vbExclamation, "提示！") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "记录已经删除！", vbOKOnly + vbExclamation, "警告!"
                Unload Me
                SQL = "select * from LeaveInfo"
                Call frmOKQResult.LeaveTopic
                Call frmOKQResult.ShowLRecord(SQL)
                Unload frmOKQResult
                frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "学生请假信息列表"
            End If
          End If
    Case "学生补课信息列表"
        If frmOKQResult.ORecordList.Rows = 1 Then
            MsgBox "目前没有补课信息！", vbOKOnly + vbExclamation, "警告！"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from OvertimeInfo where OID="
            SQL = SQL & Trim(frmOKQResult.ORecordList.TextMatrix( _
                                    frmOKQResult.ORecordList.Row, 0))
            If MsgBox("真的要删除这条记录么？", vbOKCancel + vbExclamation, "提示！") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "记录已经删除！", vbOKOnly + vbExclamation, "警告!"
                Unload Me
                SQL = "select * from OvertimeInfo"
                Call frmOKQResult.OverTimeTopic
                Call frmOKQResult.ShowORecord(SQL)
                 Unload frmOKQResult
                frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "学生补课信息列表"
            End If
        End If
    Case "学生旷课信息列表"
        If frmOKQResult.ERecordList.Rows = 1 Then
            MsgBox "目前没有旷课信息！", vbOKOnly + vbExclamation, "警告！"
            flag = 1
            frmOtherKQ.Show
            frmOtherKQ.ZOrder 0
        Else
            SQL = "delete from ErrandInfo where EID="
            SQL = SQL & Trim(frmOKQResult.ERecordList.TextMatrix( _
                                    frmOKQResult.ERecordList.Row, 0))
            If MsgBox("真的要删除这条记录么？", vbOKCancel + vbExclamation, "提示！") = vbOK _
            Then
                TransactSQL (SQL)
                MsgBox "记录已经删除！", vbOKOnly + vbExclamation, "警告!"
                Unload Me
                SQL = "select * from ErrandInfo"
                Call frmOKQResult.ErrandTopic
                Call frmOKQResult.ShowERecord(SQL)
                 Unload frmOKQResult
                'frmOKQResult.Show
                 frmOKQResult.SSTab.Caption = "学生旷课信息列"
            End If
        End If
    End Select
End Sub



