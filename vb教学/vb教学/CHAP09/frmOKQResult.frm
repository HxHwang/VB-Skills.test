VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOKQResult 
   Caption         =   "ѧ������������Ϣ�б�"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   10215
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ѧ�������Ϣ�б�"
      TabPicture(0)   =   "frmOKQResult.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LRecordList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ѧ��������Ϣ�б�"
      TabPicture(1)   =   "frmOKQResult.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ORecordList"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ѧ��������Ϣ�б�"
      TabPicture(2)   =   "frmOKQResult.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ERecordList"
      Tab(2).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid LRecordList 
         Height          =   7215
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   11000
         _ExtentX        =   19394
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid ERecordList 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   4
         Top             =   360
         Width           =   11000
         _ExtentX        =   19394
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid ORecordList 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   3
         Top             =   360
         Width           =   11000
         _ExtentX        =   19394
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "����������Ϣ�б�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmOKQResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LeaveTopic()
    Dim i As Integer
    With LRecordList                                '���������Ϣ�б��ͷ
        .TextMatrix(0, 0) = "��¼���"
        .TextMatrix(0, 1) = "ѧ�����"
        .TextMatrix(0, 2) = "�¼�����"
        .TextMatrix(0, 3) = "��������"
        .TextMatrix(0, 4) = "��ʼʱ��"
        For i = 0 To 4                             '���ö��뷽ʽ
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '�����п�
            .ColWidth(i) = 1500
        Next i
    End With
End Sub

Public Sub ShowLRecord(query As String)            '��ʾ�����Ϣ
    Dim rsLeave As New ADODB.Recordset
    Set rsLeave = TransactSQL(query)
    If rsLeave.EOF = False Then
    With LRecordList
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

Public Sub OverTimeTopic()
    Dim i As Integer
    With ORecordList                                '���ò�����Ϣ�б��ͷ
        .TextMatrix(0, 0) = "��¼���"
        .TextMatrix(0, 1) = "ѧ�����"
        .TextMatrix(0, 2) = "���ⲹ������"
        .TextMatrix(0, 3) = "������������"
        .TextMatrix(0, 4) = "����ʱ��"
        For i = 0 To 4                             '���ö��뷽ʽ
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '�����п�
            .ColWidth(i) = 1800
        Next i
    End With
End Sub

Public Sub ShowORecord(query As String)            '��ʾ������Ϣ
    Dim rsOvertime As New ADODB.Recordset
    Set rsOvertime = TransactSQL(query)
    If rsOvertime.EOF = False Then
    With ORecordList
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

Public Sub ErrandTopic()
    Dim i As Integer
    With ERecordList                                '���ÿ�����Ϣ�б��ͷ
        .TextMatrix(0, 0) = "��¼���"
        .TextMatrix(0, 1) = "ѧ�����"
        .TextMatrix(0, 2) = "��������"
        .TextMatrix(0, 3) = "����Ŀ�ĵ�"
        .TextMatrix(0, 4) = "���ο�ʼʱ��"
        For i = 0 To 4                             '���ö��뷽ʽ
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 4                             '�����п�
            .ColWidth(i) = 2000
        Next i
    End With
End Sub

Public Sub ShowERecord(query As String)             '��ʾ������Ϣ
    Dim rsErrand As New ADODB.Recordset
    Set rsErrand = TransactSQL(query)
    If rsErrand.EOF = False Then
    With ERecordList
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

Private Sub ERecordList_DblClick()                'ѡ����μ�¼�޸�
    flag = 4
    If frmOKQResult.ERecordList.Rows > 1 Then
        kqsql2 = "select * from ErrandInfo where EID=" & Trim( _
        frmOKQResult.ERecordList.TextMatrix(frmOKQResult.ERecordList.Row, 0))
        frmOtherKQ.Show
        frmOtherKQ.ZOrder 0
    ErecordID = Trim(frmOKQResult.ERecordList.TextMatrix(frmOKQResult.ERecordList.Row, 0))
    Else
     MsgBox "û�п�����Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     frmOtherKQ.Show
    End If
End Sub

Private Sub Form_Load()
    Dim sql As String
    Select Case SSTab.Caption
    Case "ѧ�������Ϣ�б�"
        sql = "select * from LeaveInfo"
        Call LeaveTopic
        Call ShowLRecord(sql)
    Case "ѧ��������Ϣ�б�"
        sql = "select * from OvertimeInfo"
        Call OverTimeTopic
        Call ShowORecord(sql)
    Case "ѧ��������Ϣ�б�"
        sql = "select * from ErrandInfo"
        Call ErrandTopic
        Call ShowERecord(sql)
    End Select
End Sub

Private Sub LRecordList_DblClick()                'ѡ����ټ�¼�޸�
    flag = 2
    If frmOKQResult.LRecordList.Rows > 1 Then
        kqsql2 = "select * from LeaveInfo where LID=" & Trim( _
        frmOKQResult.LRecordList.TextMatrix(frmOKQResult.LRecordList.Row, 0))
        frmOtherKQ.Show
        frmOtherKQ.ZOrder 0
    LrecordID = Trim(frmOKQResult.LRecordList.TextMatrix(frmOKQResult.LRecordList.Row, 0))
    Else
     MsgBox "û�������Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     frmOtherKQ.Show
    End If
End Sub

Private Sub ORecordList_DblClick()                'ѡ�񲹿μ�¼�޸�
    flag = 3
    If frmOKQResult.ORecordList.Rows > 1 Then
        kqsql2 = "select * from OvertimeInfo where OID=" & Trim( _
        frmOKQResult.ORecordList.TextMatrix(frmOKQResult.ORecordList.Row, 0))
        frmOtherKQ.Show
        frmOtherKQ.ZOrder 0
    OrecordID = Trim(frmOKQResult.ORecordList.TextMatrix(frmOKQResult.ORecordList.Row, 0))
    Else
     MsgBox "û�в�����Ϣ��", vbOKOnly + vbExclamation, "����!"
     flag = 1
     frmOtherKQ.Show
    End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    Dim sql As String
    Select Case SSTab.Caption
    Case "ѧ�������Ϣ�б�"
        sql = "select * from LeaveInfo"
        Call LeaveTopic
        Call ShowLRecord(sql)
    Case "ѧ��������Ϣ�б�"
        sql = "select * from OvertimeInfo"
        Call OverTimeTopic
        Call ShowORecord(sql)
    Case "ѧ��������Ϣ�б�"
        sql = "select * from ErrandInfo"
        Call ErrandTopic
        Call ShowERecord(sql)
    End Select
End Sub

Private Sub Lrecordlist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        PopupMenu popmenu.popmenu2
    End If
End Sub

Private Sub Orecordlist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        PopupMenu popmenu.popmenu2
    End If
End Sub

Private Sub Erecordlist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        PopupMenu popmenu.popmenu2
    End If
End Sub
