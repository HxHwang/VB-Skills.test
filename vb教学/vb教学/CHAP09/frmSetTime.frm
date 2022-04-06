VERSION 5.00
Begin VB.Form frmSetTime 
   Caption         =   "设置上下学时间"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6525
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      Begin VB.TextBox EndTime 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox BeginTime 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "下学时间："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "上学时间："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "设置上下学时间"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "delete from TimeSetting"
    TransactSQL (sql)
    If IsDate(Me.BeginTime) = False Or Me.BeginTime = "" Then
        MsgBox "请正确地输入时间！", vbOKOnly + vbExclamation, "警告！"
        Me.BeginTime.SetFocus
    Else
        If IsDate(Me.EndTime) = False Or Me.EndTime = "" Then
            MsgBox "请正确地输入时间！", vbOKOnly + vbExclamation, "警告！"
            Me.EndTime.SetFocus
    Else
        sql = "select * from TimeSetting"
        Set rs = TransactSQL(sql)
        rs.AddNew                               '设置时间
            rs.Fields(0) = Me.BeginTime
            rs.Fields(1) = Me.EndTime
            rs.Update
            rs.Close
        MsgBox "时间已经设置！", vbOKOnly + vbExclamation, "设置结果！"
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from TimeSetting"
    Set rs = TransactSQL(sql)
    If rs.EOF = True Then
        Me.BeginTime = ""
        Me.EndTime = ""
    Else
        Me.BeginTime = rs(0)
        Me.EndTime = rs(1)
    End If
    rs.Close
End Sub

