VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5385
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox LstShow 
      Height          =   4560
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "����"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "ɾ����Ϣ"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "¼����Ϣ"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ListBox LstAddr 
      Height          =   600
      ItemData        =   "Form1.frx":0000
      Left            =   960
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtNumb 
      Height          =   375
      Left            =   960
      MaxLength       =   7
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox CboMan 
      Height          =   300
      ItemData        =   "Form1.frx":0004
      Left            =   960
      List            =   "Form1.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��᣺"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ѧ�ţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim sindex As Integer
Dim infr As String
Dim Sname As String, Snumb As String, man As String, addr As String

Private Sub Form_Load()
With LstAddr
.AddItem "����"
.AddItem "�Ϻ�"
.AddItem "�����人"
.AddItem "���ϳ�ɳ"
.AddItem "�Ĵ��ɶ�"
.AddItem "�㶫����"
End With
CboMan.ListIndex = 0
LstAddr.ListIndex = 0
i = 0
sindex = 0
End Sub

Private Sub cboMan_Click()
man = CboMan.Text
End Sub

Private Sub cmdDel_Click()
LstShow.RemoveItem sindex
i = i - 1
End Sub

Private Sub cmdInput_Click()
infr = Sname + "   " + man + "   " + addr
LstShow.AddItem infr, i
i = i + 1
End Sub

Private Sub cmdQuit_Click()
Unload Form1
End Sub

Private Sub lstAddr_Click()
addr = LstAddr.Text
End Sub

Private Sub lstShow_Click()
sindex = LstShow.ListIndex
End Sub

Private Sub txtName_Change()
Sname = TxtName.Text
End Sub

Private Sub txtNumb_Change()
Snumb = TxtNumb.Text
ls = Len(Sname)
Select Case ls
    Case 2
        Sname = Sname + "     "
    Case 3
        Sname = Sname + "    "
    Case 4
        Sname = Sname
End Select
End Sub

Private Sub txtNumb_LostFocus()
    If Len(Snumb) < 7 Then
        MsgBox "ѧ�ű���Ϊ7λ��", vbOKOnly + vbCritical, "����"
        TxtNumb.SetFocus
    End If
End Sub


Private Sub cmdFind_Click()
    Dim mystr As String
    Dim mybt As Integer
    Dim j As Integer
    Dim fs As Integer
    fs = 0
    '����ת֧���
step:
    mystr = InputBox("��������Ҫ����ѧ��ѧ��", "���ҶԻ���")
    If mystr = "" Then
      mybt = MsgBox("δ����ѧ�ţ��Ƿ���������ѧ��?", _
            vbOKCancel + vbQuestion, "ȷ������")
        If mybt = 1 Then
            GoTo step
        Else
            Form1.Show
            Exit Sub
        End If
    End If
    For j = 0 To LstShow.ListCount - 1
      If mystr = Left(LstShow.List(j), 7) Then
         fs = 1
        Exit For
      End If
    Next
    If fs = 0 Then
       MsgBox "û�и�ѧ������Ϣ", vbOKOnly + vbCritical, _
       "����"
    Else
       LstShow.ListIndex = j
    End If
End Sub

