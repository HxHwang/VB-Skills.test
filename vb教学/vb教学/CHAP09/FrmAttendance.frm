VERSION 5.00
Begin VB.Form FrmAttendance 
   Caption         =   "ѧ��������Ϣ"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9630
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "ѧ��������Ϣ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   12
      Top             =   2760
      Width           =   8655
      Begin VB.Frame Frame3 
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   8295
         Begin VB.TextBox InTime 
            BeginProperty Font 
               Name            =   "����_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox OutTime 
            BeginProperty Font 
               Name            =   "����_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton InFlag 
            Caption         =   "��ѧʱ�䣺"
            BeginProperty Font 
               Name            =   "����_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton OutFlag 
            Caption         =   "��ѧʱ�䣺"
            BeginProperty Font 
               Name            =   "����_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4320
            TabIndex        =   5
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.TextBox NowDate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "��ǰ���ڣ�"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧ��������Ϣ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   8655
      Begin VB.ComboBox ASID 
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox ASName 
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ѧ��������"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ѧ����ţ�"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label topic 
      Alignment       =   2  'Center
      Caption         =   "ѧ������ѧ��Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "FrmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ilate As Integer                                  '�ٵ�����
Private iearly As Integer                                 '���˴���
Private aflag As String                                   '�����־
Private addflag As Boolean                                '��ӱ�־
Private firstID As String                                 '��һ��ѧ�����

Private Sub ASID_KeyDown(KeyCode As Integer, Shift As Integer)
    TabToEnter KeyCode
End Sub

Private Sub ASID_LostFocus()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select SName from StuffInfo where SID='" & Me.ASID.Text & "'"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        Me.ASName = rs(0)                           '��ʼ��ѧ������
    Else
        MsgBox "ѧ�����������󣬻���û�����ѧ����", vbOKOnly + vbExclamation, "���棡"
        Me.ASID = ""
        Me.ASID.SetFocus
        Me.ASID.ListIndex = 0
    End If
    rs.Close
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CheckRecord()                           '�ж��Ƿ���ڼ�¼
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo where AStuffID='" & Me.ASID.Text & "'"
    SQL = SQL & " and AFlag='" & aflag & "' and ADate=#" & Me.NowDate & "#"
        Set rs = TransactSQL(SQL)
        If rs.EOF = False Then
            MsgBox "�Ѿ�����������¼��", vbOKOnly + vbExclamation, "���棡"
            addflag = True
        Else
            addflag = False
        End If
        rs.Close
End Sub

Private Sub in_add()                                '�����ѧ��¼
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo"
    Set rs = TransactSQL(SQL)
    rs.AddNew
    rs.Fields(1) = Me.ASID
    rs.Fields(2) = Me.ASName
    rs.Fields(3) = Me.NowDate
    rs.Fields(4) = aflag
    rs.Fields(5) = Me.InTime
    rs.Fields(7) = ilate
    rs.Update
    rs.Close
End Sub

Private Sub out_add()                               '��ӷ�ѧ��¼
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select * from AttendanceInfo"
    Set rs = TransactSQL(SQL)
    rs.AddNew
    rs.Fields(1) = Me.ASID
    rs.Fields(2) = Me.ASName
    rs.Fields(3) = Me.NowDate
    rs.Fields(4) = aflag
    rs.Fields(6) = Me.OutTime
    rs.Fields(8) = iearly
    rs.Update
    rs.Close
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim sql2 As String
    Dim rs As New ADODB.Recordset
    Dim rsTime As New ADODB.Recordset
    sql2 = "select * from AttendanceInfo order by ID desc"
    SQL = "select * from TimeSetting"
    Set rsTime = TransactSQL(SQL)
    If flag = 1 Then
    ilate = 0
    iearly = 0
    If Me.InFlag = False And Me.OutFlag = False Then
        MsgBox "��ѡ������ѧ��", vbOKOnly + vbExclamation, "���棡"
    Else
    If Me.InFlag = True Then                         '�����ѧ��¼
        aflag = "��"
        If Me.InTime = "" Or IsDate(Me.InTime) = False Then
            MsgBox "��������ȷ��ʱ�䣡", vbOKOnly + vbExclamation, "����!"
            Me.InTime = ""
            Me.InTime.SetFocus
        Else
            If DateDiff("s", Me.InTime, rsTime(0)) < 0 Then
                ilate = 1
            End If
            Call CheckRecord
            If addflag = False Then
                Call in_add
                MsgBox "�Ѿ������ѧ��¼��", vbOKOnly + vbExclamation, "��ӽ����"
                Call init
                Me.InFlag = False
            Else
                Call init
                Me.InFlag = False
            End If
        End If
    End If
    If Me.OutFlag = True Then                        '��ӷ�ѧ��¼
        aflag = "��"
        If Me.OutTime = "" Or IsDate(Me.OutTime) = False Then
            MsgBox "��������ȷ��ʱ�䣡", vbOKOnly + vbExclamation, "����!"
            Me.OutTime = ""
            Me.OutTime.SetFocus
        Else
            If DateDiff("s", Me.OutTime, rsTime(1)) > 0 Then
                iearly = 1
            End If
            Call CheckRecord
            If addflag = False Then
                Call out_add
                MsgBox "�Ѿ���ӷ�ѧ��¼��", vbOKOnly + vbExclamation, "��ӽ����"
                Call init
                Me.OutFlag = False
            Else
                Call init
                Me.OutFlag = False
            End If
        End If
    End If
    End If
        Call frmAResult.ListTopic
        Call frmAResult.ShowData(sql2)
        frmAResult.Show
        frmAResult.ZOrder 0
        Me.ZOrder 0
    Else                                             '�޸ļ�¼
        If MsgBox("ȷ���޸ı��Ϊ" & Me.ASID & "��ѧ����Ϣ?", vbOKCancel, "��ʾ��") _
                                                                = vbOK Then
            If Me.InFlag = True Then
                If DateDiff("s", Me.InTime, rsTime(0)) < 0 Then
                    ilate = 1
                End If
                SQL = "update AttendanceInfo set AInTime=#" & Me.InTime & "#,"
                SQL = SQL & "ALate=" & ilate & " where ID=" & ArecordID
                TransactSQL (SQL)                     '�޸���ѧ��¼
                Call frmAResult.ListTopic
                Call frmAResult.ShowData(sql2)
                frmAResult.Show
                MsgBox "��Ϣ�Ѿ��޸ģ�", vbOKOnly + vbExclamation, "�޸Ľ����"
                Unload Me
            ElseIf Me.OutFlag = True Then
                If DateDiff("s", Me.OutTime, rsTime(1)) > 0 Then
                    iearly = 1
                End If
                SQL = "update AttendanceInfo set AOutTime=#" & Me.OutTime & "#,"
                SQL = SQL & "AEarly=" & iearly & " where ID=" & ArecordID
                TransactSQL (SQL)                     '�޸ķ�ѧ��¼
                Call frmAResult.ListTopic
                Call frmAResult.ShowData(sql2)
                frmAResult.Show
                MsgBox "��Ϣ�Ѿ��޸ģ�", vbOKOnly + vbExclamation, "�޸Ľ����"
                Unload Me
            End If
        Else
        Unload Me
        End If
    End If
    rsTime.Close
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    If flag = 1 Then
    SQL = "select SID from StuffInfo order by SID"
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        rs.MoveFirst
        firstID = rs(0)
    While Not rs.EOF
        Me.ASID.AddItem rs(0)                     '��ʼ��ѧ�����
        rs.MoveNext
    Wend
        rs.Close
    Else
        MsgBox "Ŀǰû��ѧ����", vbOKOnly + vbExclamation, "���棡"
    End If
    Me.NowDate = Date
    Me.ASID.ListIndex = 0
    SQL = "select SName from StuffInfo where SID='" & firstID & "'"
    Set rs = TransactSQL(SQL)
    Me.ASName = rs(0)                             '��ʼ��ѧ������
    rs.Close
    Me.OutTime = ""
    Me.InTime = ""
   ElseIf flag = 2 Then
       
        Set rs = TransactSQL(kqsql)
         'If rs.EOF = False And rs.BOF Then

        If rs.EOF = False Then
         rs.MoveFirst
        firstID = rs(0)
        With rs
            Me.ASID = rs(1)
            Me.ASName = rs(2)
            Me.NowDate = rs(3)
            If IsNull(rs(5)) = True Then
            Me.InTime = ""
            Me.OutFlag = True
            Else
            Me.InTime = rs(5)
            End If
            If IsNull(rs(6)) = True Then
            Me.OutTime = ""
            Me.InFlag = True
            Else
            Me.OutTime = rs(6)
            End If
        End With
        rs.Close
        End If
    End If
    
End Sub
Private Sub init()                                '��ʼ��
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    SQL = "select SName from StuffInfo where SID='" & firstID & "'"
    Set rs = TransactSQL(SQL)
    Me.ASID.ListIndex = 0
    Me.ASName = rs(0)
    Me.InTime = ""
    Me.OutTime = ""
End Sub


