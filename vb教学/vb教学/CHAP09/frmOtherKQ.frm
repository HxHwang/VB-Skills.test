VERSION 5.00
Begin VB.Form frmOtherKQ 
   Caption         =   "���ѧ��������Ϣ"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   8940
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      Caption         =   "��ʼʱ����Ϣ"
      Height          =   855
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   8415
      Begin VB.TextBox FromDay 
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
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "��ʼʱ�䣺"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
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
      Left            =   5640
      TabIndex        =   10
      Top             =   6720
      Width           =   1815
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
      Left            =   1440
      TabIndex        =   9
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "ѧ��������Ϣ"
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   8415
      Begin VB.TextBox EDays 
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
         Left            =   6000
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox EPurpose 
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lab 
         Alignment       =   1  'Right Justify
         Caption         =   "����������"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         Caption         =   "����Ŀ�ģ�"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ѧ��������Ϣ"
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   8415
      Begin VB.TextBox SOverDays 
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
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox COverDays 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "���ⲹ��������"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "��������������"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ѧ�������Ϣ"
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   8415
      Begin VB.TextBox ILeave 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox PLeave 
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
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "���٣�"
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
         Left            =   4800
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "�¼٣�"
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
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧ��������Ϣ"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   8415
      Begin VB.ComboBox ASID 
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   2175
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
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ѧ������������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2280
      TabIndex        =   23
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmOtherKQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstID As String                               'ѧ��ID

Private Sub ASID_KeyDown(KeyCode As Integer, Shift As Integer)
    TabToEnter KeyCode
End Sub

Private Sub ASID_LostFocus()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select SName from StuffInfo where SID='" & Me.ASID.Text & "'"
    Set rs = TransactSQL(sql)
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

Private Sub cmdOK_Click()
    Dim sql As String
    Dim rsTime As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim ipleave As Integer                           '�����¼�����
    Dim iileave As Integer                           '���벡������
    Dim COverDays As Integer                         '������������
    Dim SOverDays As Integer                         '���ⲹ������
    ipleave = 0
    iileave = 0
    COverDays = 0
    SOverDays = 0
    If IsDate(Me.FromDay) = False Then
                MsgBox "������ȷ�Ŀ�ʼʱ�䣡", vbOKOnly + vbExclamation, "���棡"
                Me.FromDay = ""
                Me.FromDay.SetFocus
            End If
    If Me.PLeave <> "" Then
        If IsNumeric(Me.PLeave) = False Then
            MsgBox "������¼�������Ϊ������", vbOKOnly + vbExclamation, "���棡"
            Me.PLeave = ""
            Me.PLeave.SetFocus
        Else
            ipleave = Me.PLeave
        End If
    End If
    If Me.ILeave <> "" Then
        If IsNumeric(Me.ILeave) = False Then
            MsgBox "����Ĳ���������Ϊ������", vbOKOnly + vbExclamation, "���棡"
            Me.ILeave = ""
            Me.ILeave.SetFocus
        Else
            iileave = Me.ILeave
        End If
    End If
    If Me.COverDays <> "" Then
        If IsNumeric(Me.COverDays) = False Then
            MsgBox "������������Ϊ������", vbOKOnly + vbExclamation, "���棡"
            Me.COverDays = ""
            Me.COverDays.SetFocus
        Else
            COverDays = Me.COverDays
        End If
    End If
    If Me.SOverDays <> "" Then
        If IsNumeric(Me.SOverDays) = False Then
            MsgBox "���ⲹ������Ϊ������", vbOKOnly + vbExclamation, "���棡"
            Me.SOverDays = ""
            Me.SOverDays.SetFocus
        Else
            SOverDays = Me.SOverDays
        End If
    End If
    If Me.EDays <> "" Or Me.EPurpose <> "" Then
        If Me.EDays = "" Then
            MsgBox "���������������", vbOKOnly + vbExclamation, "���棡"
            Me.EDays = ""
            Me.EDays.SetFocus
        ElseIf IsNumeric(Me.EDays) = False Then
            MsgBox "��������Ϊ������", vbOKOnly + vbExclamation, "���棡"
            Me.EDays = ""
            Me.EDays.SetFocus
        ElseIf Me.EPurpose = "" Then
            MsgBox "���������Ŀ�ģ�", vbOKOnly + vbExclamation, "���棡"
            Me.EPurpose = ""
            Me.EPurpose.SetFocus
        End If
    End If
    If flag = 1 Then                                          '��Ӽ�¼
        If Me.PLeave = "" And Me.ILeave = "" And Me.EPurpose = "" _
                    And Me.EDays = "" And Me.COverDays = "" And Me.SOverDays = "" Then
        Else
            If Me.PLeave <> "" Or Me.ILeave <> "" Then
                sql = "select * from LeaveInfo"              '�����ټ�¼
                Set rs = TransactSQL(sql)
                rs.AddNew
                rs.Fields(1) = Me.ASID
                rs.Fields(2) = iileave
                rs.Fields(3) = ipleave
                rs.Fields(4) = Me.FromDay
                rs.Update
                rs.Close
                MsgBox "�Ѿ������ټ�¼��", vbOKOnly + vbExclamation, "��ӽ����"
                Call init
            ElseIf Me.COverDays <> "" _
                        Or Me.SOverDays <> "" Then            '��Ӳ�����Ϣ
                sql = "select * from OvertimeInfo"
                Set rs = TransactSQL(sql)
                rs.AddNew
                rs.Fields(1) = Me.ASID
                rs.Fields(2) = SOverDays
                rs.Fields(3) = COverDays
                rs.Fields(4) = Me.FromDay
                rs.Update
                rs.Close
                MsgBox "�Ѿ���Ӳ�����Ϣ��", vbOKOnly + vbExclamation, "��ӽ����"
                Call init
            ElseIf Me.EDays <> "" And Me.EPurpose <> "" Then '��ӿ��μ�¼
                sql = "select * from ErrandInfo"
                Set rs = TransactSQL(sql)
                rs.AddNew
                rs.Fields(1) = Me.ASID
                rs.Fields(2) = Me.EDays
                rs.Fields(3) = Me.EPurpose
                rs.Fields(4) = Me.FromDay
                rs.Update
                rs.Close
                MsgBox "�Ѿ���ӿ��μ�¼��", vbOKOnly + vbExclamation, "��ӽ����"
                Call init
            End If
        End If
        Select Case frmOKQResult.SSTab.Caption
        Case "ѧ�������Ϣ�б�"
            sql = "select * from LeaveInfo"
            Call frmOKQResult.LeaveTopic
            Call frmOKQResult.ShowLRecord(sql)
        Case "ѧ��������Ϣ�б�"
            sql = "select * from OvertimeInfo"
            Call frmOKQResult.OverTimeTopic
            Call frmOKQResult.ShowORecord(sql)
        Case "ѧ��������Ϣ�б�"
            sql = "select * from ErrandInfo"
            Call frmOKQResult.ErrandTopic
            Call frmOKQResult.ShowERecord(sql)
        End Select
        frmOKQResult.Show
        frmOKQResult.ZOrder 0
        Me.ZOrder 0
    Else
        If flag = 2 Then                                      '�޸������Ϣ
            If Me.PLeave <> "" And Me.ILeave <> "" Then
                If MsgBox("ȷ���޸ı��Ϊ" & Me.ASID & "ѧ���������Ϣ��", vbOKCancel) _
                                                                        = vbOK Then
                sql = "update LeaveInfo set LILL=" & ILeave
                sql = sql & ",LPrivate=" & PLeave & ",LFromDay=#" & Me.FromDay
                sql = sql & "# where LID=" & LrecordID
                TransactSQL (sql)
                MsgBox "��Ϣ�Ѿ��޸ģ�", vbOKOnly + vbExclamation, "�޸Ľ����"
                sql = "select * from LeaveInfo"
                Call frmOKQResult.LeaveTopic
                Call frmOKQResult.ShowLRecord(sql)
                frmOKQResult.Show
                frmOKQResult.ZOrder 0
                Unload Me
                End If
            End If
        ElseIf flag = 3 Then                                  '�޸Ĳ�����Ϣ
            If Me.COverDays <> "" And Me.SOverDays <> "" Then
                If MsgBox("ȷ���޸ı��Ϊ" & Me.ASID & "ѧ���Ĳ�����Ϣ��", vbOKCancel) _
                                                                        = vbOK Then
                sql = "update OvertimeInfo set OSpeciality=" & SOverDays
                sql = sql & ",OCommon=" & COverDays & ",OFromDay=#" & Me.FromDay & "#"
                sql = sql & " where OID=" & OrecordID
                TransactSQL (sql)
                sql = "select * from OvertimeInfo"
                Call frmOKQResult.OverTimeTopic
                Call frmOKQResult.ShowORecord(sql)
                frmOKQResult.Show
                frmOKQResult.ZOrder 0
                Unload Me
                End If
            End If
        Else
            If Me.EDays <> "" And Me.EPurpose <> "" Then      '�޸Ŀ�����Ϣ
                If MsgBox("ȷ���޸ı��Ϊ" & Me.ASID & "ѧ���Ŀ�����Ϣ��", vbOKCancel) _
                                                                        = vbOK Then
                sql = "update ErrandInfo set EErranddays=" & Me.EDays
                sql = sql & ",EPurpose='" & Me.EPurpose & "'"
                sql = sql & ",EFromday=#" & Me.FromDay & "#"
                sql = sql & " where EID=" & ErecordID
                TransactSQL (sql)
                sql = "select * from ErrandInfo"
                Call frmOKQResult.ErrandTopic
                Call frmOKQResult.ShowERecord(sql)
                frmOKQResult.Show
                frmOKQResult.ZOrder 0
                Unload Me
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim rsName As New ADODB.Recordset
    If flag = 1 Then
    sql = "select SID from StuffInfo order by SID"
    Set rs = TransactSQL(sql)
    If rs.EOF = False Then
        rs.MoveFirst
        firstID = rs(0)
    While Not rs.EOF
        Me.ASID.AddItem rs(0)                          '��ʼ��ѧ�����
        rs.MoveNext
    Wend
        rs.Close
    Else
        MsgBox "Ŀǰû��ѧ����", vbOKOnly + vbExclamation, "���棡"
    End If
    Me.ASID.ListIndex = 0
    sql = "select SName from StuffInfo where SID ='" & firstID & "'"
    Set rs = TransactSQL(sql)
    Me.ASName = rs(0)                                  '��ʼ��ѧ������
    Me.FromDay = Date
    rs.Close
    ElseIf flag = 2 Then                               '���������Ϣ
        Set rs = TransactSQL(kqsql2)
        If rs.EOF = False Then
        With rs
            Me.ASID = rs(1)
            sql = "select SName from StuffInfo where SID='" & rs(1) & "'"
            Set rsName = TransactSQL(kqsql2)
            Me.ASName = rsName(0)
            Me.FromDay = rs(4)
            Me.ILeave = rs(2)
            Me.PLeave = rs(3)
        End With
        End If
        rsName.Close
        rs.Close
    ElseIf flag = 3 Then                                '���벹����Ϣ
        Set rs = TransactSQL(kqsql2)
        If rs.EOF = False Then
        With rs
            Me.ASID = rs(1)
            sql = "select SName from StuffInfo where SID='" & rs(1) & "'"
            Set rsName = TransactSQL(sql)
            Me.ASName = rsName(0)
            Me.SOverDays = rs(2)
            Me.COverDays = rs(3)
            Me.FromDay = rs(4)
        End With
        End If
        rsName.Close
        rs.Close
    ElseIf flag = 4 Then                                 '���������Ϣ
        Set rs = TransactSQL(kqsql2)
        If rs.EOF = False Then
        With rs
            Me.ASID = rs(1)
            sql = "select SName from StuffInfo where SID='" & rs(1) & "'"
            Set rsName = TransactSQL(sql)
            Me.ASName = rsName(0)
            Me.EDays = rs(2)
            Me.EPurpose = rs(3)
            Me.FromDay = rs(4)
        End With
        End If
        rsName.Close
        rs.Close
    End If
End Sub

Private Sub init()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select SName from StuffInfo where SID='" & firstID & "'"
    Set rs = TransactSQL(sql)
    Me.ASID.ListIndex = 0
    Me.ASName = rs(0)
    Me.PLeave = ""
    Me.ILeave = ""
    Me.COverDays = ""
    Me.SOverDays = ""
    Me.EPurpose = ""
    Me.EDays = ""
End Sub

