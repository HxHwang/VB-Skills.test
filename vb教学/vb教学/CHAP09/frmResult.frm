VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResult 
   Caption         =   "ѧ��������Ϣ�б�"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   7275
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid rsGrid 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   12568
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   16777215
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
   Begin VB.Label Label1 
      Caption         =   "ѧ��������Ϣ�б�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3285
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim SQL As String
    SQL = "select * from StuffInfo order by SID"
    createList (SQL)
End Sub

Public Sub createList(SQL As String)
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim rsGird As MSFlexGrid
    With rsGrid                                   '���ñ�ͷ
        .TextMatrix(0, 0) = "ѧ�����"
        .TextMatrix(0, 1) = "ѧ������"
        .TextMatrix(0, 2) = "ѧ���Ա�"
        .TextMatrix(0, 3) = "ѧ������"
        .TextMatrix(0, 4) = "ѧ������"
        .TextMatrix(0, 5) = "ѧ������"
        .TextMatrix(0, 6) = "ѧ���꼶"
        .TextMatrix(0, 7) = "ѧ��רҵ"
        .TextMatrix(0, 8) = "��ͥסַ"
        .TextMatrix(0, 9) = "��������"
        .TextMatrix(0, 10) = "�绰����"
        .TextMatrix(0, 11) = "Email"
        .TextMatrix(0, 12) = "��ѧʱ��"
        .TextMatrix(0, 13) = "���뱾Уʱ��"
        .TextMatrix(0, 14) = "�༶"
        .TextMatrix(0, 15) = "��ʽ�Ͽ�ʱ��"
        .TextMatrix(0, 16) = "�༶ְ��"
        .TextMatrix(0, 17) = "��ע"
        For i = 0 To 17                             '�������б����뷽ʽ
            .ColAlignment(i) = 4
        Next i
        For i = 0 To 11                             '����ÿ�п��
            .ColWidth(i) = 1400
        Next i
        .ColWidth(12) = 2000
        .ColWidth(13) = 2000
        .ColWidth(14) = 1400
        .ColWidth(15) = 2000
        .ColWidth(16) = 1400
        .ColWidth(17) = 3000
    End With
    Set rs = TransactSQL(SQL)
    If rs.EOF = False Then
        With rsGrid                                 '��ʾ��Ϣ����
        .Rows = 1
        While Not rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rs(0)
            .TextMatrix(.Rows - 1, 1) = rs(1)
            .TextMatrix(.Rows - 1, 2) = rs(2)
            .TextMatrix(.Rows - 1, 3) = rs(3)
            .TextMatrix(.Rows - 1, 4) = rs(4)
            .TextMatrix(.Rows - 1, 5) = rs(5)
            .TextMatrix(.Rows - 1, 6) = rs(6)
            .TextMatrix(.Rows - 1, 7) = rs(7)
            .TextMatrix(.Rows - 1, 8) = rs(8)
            .TextMatrix(.Rows - 1, 9) = rs(9)
            .TextMatrix(.Rows - 1, 10) = rs(10)
            .TextMatrix(.Rows - 1, 11) = rs(11)
            .TextMatrix(.Rows - 1, 12) = rs(12)
            .TextMatrix(.Rows - 1, 13) = rs(13)
            .TextMatrix(.Rows - 1, 14) = rs(14)
            .TextMatrix(.Rows - 1, 15) = rs(15)
            .TextMatrix(.Rows - 1, 16) = rs(16)
            .TextMatrix(.Rows - 1, 17) = rs(17)
            rs.MoveNext
        Wend
        End With
    rs.Close
    End If
End Sub

Private Sub rsGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        PopupMenu popmenu.popmenu
    End If
End Sub

