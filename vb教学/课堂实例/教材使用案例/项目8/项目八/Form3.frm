VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form3"
   ScaleHeight     =   4230
   ScaleWidth      =   3885
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdBack 
      Caption         =   "����"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "��һ��¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "��һ��¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "�����ɼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox TxtScore 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox TxtNum 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�ɼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѧ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   420
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type stu
  sNum  As String * 10
  sName As String * 10
  Score As String * 4
End Type

Dim gstu As stu
Dim recordlen As Integer
Dim currentrecord As Integer
Dim lastrecord As Integer
Public Sub ShowCurrent()
    '��ʾ��ǰ��¼
    Get #1, currentrecord, gstu
    TxtNum.text = gstu.sNum
    TxtName.text = gstu.sName
    TxtScore.text = gstu.Score
End Sub

Public Sub SaveCurrent()
    '���浱ǰ��¼
    gstu.sNum = TxtNum.text
    gstu.sName = TxtName.text
    gstu.Score = TxtScore.text
    Put #1, currentrecord, gstu

End Sub

Private Sub Form_Load()
    recordlen = Len(gstu)
    If fName <> "" Then
        Open fName For Random As #1 Len = recordlen
        currentrecord = 1
        lastrecord = FileLen(fName) / recordlen
        If lastrecord = 0 Then
            lastrecord = 1
        End If
        ShowCurrent
    End If
End Sub

Private Sub cmdPrevious_Click()
    '�����ǰ��¼��Ϊ��1����¼��Ҳ��������ʾ
    If currentrecord = 1 Then
        Beep
        MsgBox "�ѵ��ļ�������", vbOKOnly + vbExclamation, "����"
    Else
    '�����ǰ���ǵ�1����¼�����ȱ��浱ǰ��¼
        'Ȼ������ʾ��ǰ��¼
        SaveCurrent
        '����ǰ��¼�Ƶ���һ����¼
        currentrecord = currentrecord - 1
        '��ʾ��ǰ��¼
        ShowCurrent
    End If
    TxtNum.SetFocus
End Sub


Private Sub cmdNext_Click()
    '�������¼Ϊ���ļ�¼����������ʾ
    If currentrecord = lastrecord Then
        Beep
        MsgBox "����ʾ��ȫ���ɼ���", vbOKOnly + vbExclamation, "����"
    Else
    '�����ǰ��¼��������¼�����ȱ��浱ǰ��¼
        'Ȼ������ʾ��ǰ��¼
        SaveCurrent
        '��ǰ��¼�Ƶ���һ����¼
        currentrecord = currentrecord + 1
        '��ʾ��ǰ��¼
        ShowCurrent
    End If
    TxtNum.SetFocus

End Sub

Private Sub cmdAdd_Click()
    '��������ļ�¼���浽�ļ�������¼
    SaveCurrent
    '���ļ����������1���հ׼�¼��������
    lastrecord = lastrecord + 1
    currentrecord = lastrecord
    '����󣬽��ı����е��������
    TxtNum.text = ""
    TxtName.text = ""
    TxtScore.text = ""
    TxtNum.SetFocus

End Sub


Private Sub cmdFind_Click()
    Dim nsearch As String
    Dim found As Boolean
    Dim recnum As Long
    Dim fstu As stu
    '����Ҫ���ҵ�ѧ����ѧ��
    nsearch = InputBox("������Ҫ���ҵ�ѧ����ѧ�ţ�", "����")
    If nsearch = "" Then
        Exit Sub
    End If
    found = False
    '���ļ��ĵ�һ����¼��ʼ����
    'ֱ���ҵ�ĳ����¼�е�ѧ���ֶκ��������ѧ��һ��Ϊֹ
    For recnum = 1 To lastrecord
        Get #1, recnum, fstu
        If nsearch = Trim(fstu.sNum) Then
            found = True
            Exit For
        End If
    Next
    '����ҵ��ˣ�����ʾ�ü�¼
    If found = True Then
        SaveCurrent
        currentrecord = recnum
        ShowCurrent
    '������ʾ�û�δ�ҵ���ѧ��
    Else
        MsgBox "��ѧ��Ϊ" + nsearch + "�ĳɼ�"
    End If
End Sub

Private Sub cmdback_Click()
    Unload Form3
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #1
End Sub
