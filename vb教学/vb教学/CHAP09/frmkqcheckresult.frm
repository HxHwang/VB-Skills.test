VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkqcheckresult 
   Caption         =   "出勤查询结果列表"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   673.882
   ScaleMode       =   0  'User
   ScaleWidth      =   1001.192
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Erecordlist 
      Height          =   1275
      Left            =   0
      TabIndex        =   7
      Top             =   5160
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   2249
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Orecordlist 
      Height          =   1200
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   2117
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Lrecordlist 
      Height          =   1155
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   2037
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Arecordlist 
      Height          =   1185
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   2090
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "旷课查询结果列表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "补课查询结果列表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "请假查询结果列表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "出勤查询结果列表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmkqcheckresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()

End Sub
Public Sub ATopic()
    Dim i As Integer
    With Arecordlist                               '设置表头
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

Public Sub ShowAResult(query As String)
    Dim rsAttendance As New ADODB.Recordset
    Set rsAttendance = TransactSQL(query)
    If rsAttendance.EOF = False Then
    With Arecordlist
        .Rows = 1
        While Not rsAttendance.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsAttendance(0)
            .TextMatrix(.Rows - 1, 1) = rsAttendance(1)
            .TextMatrix(.Rows - 1, 2) = rsAttendance(2)
            .TextMatrix(.Rows - 1, 3) = rsAttendance(3)
            .TextMatrix(.Rows - 1, 4) = rsAttendance(4)
            If IsNull(rsAttendance(5)) = True Then
            .TextMatrix(.Rows - 1, 5) = ""
            Else
            .TextMatrix(.Rows - 1, 5) = rsAttendance(5)
            End If
            If IsNull(rsAttendance(6)) = True Then
            .TextMatrix(.Rows - 1, 6) = ""
            Else
            .TextMatrix(.Rows - 1, 6) = rsAttendance(6)
            End If
            .TextMatrix(.Rows - 1, 7) = rsAttendance(7)
            .TextMatrix(.Rows - 1, 8) = rsAttendance(8)
            rsAttendance.MoveNext
        Wend
    End With
    rsAttendance.Close
    End If
End Sub
Public Sub LTopic()
    Dim i As Integer
    With Lrecordlist                                '设置请假信息列表表头
        .TextMatrix(0, 0) = "记录编号"
        .TextMatrix(0, 1) = "学生编号"
        .TextMatrix(0, 2) = "病假天数"
        .TextMatrix(0, 3) = "事假天数"
        .TextMatrix(0, 4) = "开始时间"
        For i = 0 To 4                             '设置对齐方式
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '设置列宽
            .ColWidth(i) = 1500
        Next i
    End With
End Sub

Public Sub ShowLResult(query As String)            '显示请假信息
    Dim rsLeave As New ADODB.Recordset
    Set rsLeave = TransactSQL(query)
    If rsLeave.EOF = False Then
    With Lrecordlist
        .Rows = 1
        While Not rsLeave.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsLeave(0)
            .TextMatrix(.Rows - 1, 1) = rsLeave(1)
            .TextMatrix(.Rows - 1, 2) = rsLeave(2)
            .TextMatrix(.Rows - 1, 3) = rsLeave(3)
            .TextMatrix(.Rows - 1, 4) = rsLeave(4)
            rsLeave.MoveNext
        Wend
        rsLeave.Close
    End With
    End If
End Sub

Public Sub OTopic()
    Dim i As Integer
    With Orecordlist                                '设置补课信息列表表头
        .TextMatrix(0, 0) = "记录编号"
        .TextMatrix(0, 1) = "学生编号"
        .TextMatrix(0, 2) = "特殊补课天数"
        .TextMatrix(0, 3) = "正常补课天数"
        .TextMatrix(0, 4) = "补课时间"
        For i = 0 To 4                             '设置对齐方式
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '设置列宽
            .ColWidth(i) = 1800
        Next i
    End With
End Sub

Public Sub ShowOResult(query As String)            '显示加班信息
    Dim rsOvertime As New ADODB.Recordset
    Set rsOvertime = TransactSQL(query)
    If rsOvertime.EOF = False Then
    With Orecordlist
        .Rows = 1
        While Not rsOvertime.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsOvertime(0)
            .TextMatrix(.Rows - 1, 1) = rsOvertime(1)
            .TextMatrix(.Rows - 1, 2) = rsOvertime(2)
            .TextMatrix(.Rows - 1, 3) = rsOvertime(3)
            .TextMatrix(.Rows - 1, 4) = rsOvertime(4)
            rsOvertime.MoveNext
        Wend
        rsOvertime.Close
    End With
    End If
End Sub

Public Sub ETopic()
    Dim i As Integer
    With Erecordlist                                '设置旷课信息列表表头
        .TextMatrix(0, 0) = "记录编号"
        .TextMatrix(0, 1) = "学生编号"
        .TextMatrix(0, 2) = "旷课天数"
        .TextMatrix(0, 3) = "旷课目的"
        .TextMatrix(0, 4) = "旷课开始时间"
        For i = 0 To 4                             '设置对齐方式
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '设置列宽
            .ColWidth(i) = 2000
        Next i
    End With
End Sub

Public Sub ShowEReslut(query As String)             '显示旷课信息
    Dim rsErrand As New ADODB.Recordset
    Set rsErrand = TransactSQL(query)
    If rsErrand.EOF = False Then
    With Erecordlist
        .Rows = 1
        While Not rsErrand.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsErrand(0)
            .TextMatrix(.Rows - 1, 1) = rsErrand(1)
            .TextMatrix(.Rows - 1, 2) = rsErrand(2)
            .TextMatrix(.Rows - 1, 3) = rsErrand(3)
            .TextMatrix(.Rows - 1, 4) = rsErrand(4)
            rsErrand.MoveNext
        Wend
        rsErrand.Close
    End With
    End If
End Sub

