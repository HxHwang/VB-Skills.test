VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAResult 
   Caption         =   "学生出勤结果列表"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid recordlist 
      Height          =   5775
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   16777215
      FillStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "学生出勤信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmAResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ListTopic()
    Dim i As Integer
    With recordlist                                   '设置表头
        .TextMatrix(0, 0) = "记录编号"
        .TextMatrix(0, 1) = "学生编号"
        .TextMatrix(0, 2) = "学生姓名"
        .TextMatrix(0, 3) = "出勤日期"
        .TextMatrix(0, 4) = "进出标志"
        .TextMatrix(0, 5) = "上学时间"
        .TextMatrix(0, 6) = "下学时间"
        .TextMatrix(0, 7) = "迟到次数"
        .TextMatrix(0, 8) = "早退次数"
        For i = 0 To 8                             '设置所有表格对齐方式
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 8                             '设置每列宽度
            .ColWidth(i) = 1500
        Next i
    End With
End Sub

Public Sub ShowData(query As String)
    Dim rsAttendance As New ADODB.Recordset
    Set rsAttendance = TransactSQL(query)
    If rsAttendance.EOF = False Then
    With recordlist
        .Rows = 1
        While Not rsAttendance.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsAttendance(0)
            .TextMatrix(.Rows - 1, 1) = rsAttendance(1)
            .TextMatrix(.Rows - 1, 2) = rsAttendance(2)
            .TextMatrix(.Rows - 1, 3) = rsAttendance(3)
            .TextMatrix(.Rows - 1, 4) = rsAttendance(4)
            If IsNull(rsAttendance(5)) Then
            .TextMatrix(.Rows - 1, 5) = ""
            Else
            .TextMatrix(.Rows - 1, 5) = rsAttendance(5)
            End If
            If IsNull(rsAttendance(6)) Then
            .TextMatrix(.Rows - 1, 6) = ""
            Else
            .TextMatrix(.Rows - 1, 6) = rsAttendance(6)
            End If
            .TextMatrix(.Rows - 1, 7) = rsAttendance(7)
            .TextMatrix(.Rows - 1, 8) = rsAttendance(8)
            rsAttendance.MoveNext
        Wend
        rsAttendance.Close
    End With
    End If
End Sub

Private Sub Form_Load()
    Dim SQL As String
    SQL = "select * from AttendanceInfo order by ID desc"
    Call ListTopic
    Call ShowData(SQL)
End Sub

Private Sub recordlist_DblClick()
    flag = 2
    If frmAResult.recordlist.Rows > 1 Then
        kqsql = "select * from AttendanceInfo where ID=" & Trim( _
        frmAResult.recordlist.TextMatrix(frmAResult.recordlist.Row, 0))
        FrmAttendance.Show
        FrmAttendance.ZOrder 0
        ArecordID = Trim(frmAResult.recordlist.TextMatrix(frmAResult.recordlist.Row, 0))
    Else
     MsgBox "没有出勤信息！", vbOKOnly + vbExclamation, "警告!"
     flag = 1
     FrmAttendance.Show
    End If
End Sub

Private Sub recordlist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        PopupMenu popmenu.popmenu1
    End If
End Sub
