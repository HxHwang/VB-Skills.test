VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ѧ��ѡ��"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5145
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2850
      Width           =   420
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   300
      Left            =   2295
      TabIndex        =   4
      Top             =   2400
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   315
      Left            =   2310
      TabIndex        =   3
      Top             =   1935
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   300
      Left            =   2325
      TabIndex        =   2
      Top             =   1485
      Width           =   420
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2850
      TabIndex        =   1
      Top             =   1185
      Width           =   1845
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   225
      TabIndex        =   0
      Top             =   1185
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   "��ѡѧ�ƣ�"
      Height          =   255
      Left            =   2835
      TabIndex        =   8
      Top             =   855
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "����ѧ�ƣ�"
      Height          =   225
      Left            =   225
      TabIndex        =   7
      Top             =   870
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "ѧ��ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   165
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    Dim m As Integer, n As Integer
    m = List1.ListCount
    Rem ����ѭ�����б��List1���б���ȫ���ƶ���List2
    For n = 0 To m - 1
        List2.AddItem (List1.List(0))   '�ƶ��б�򶥶��б���
        List1.RemoveItem (0)
    Next n
    List2.Selected(0) = True            '���ñ�ѡ���б���
End Sub

Private Sub Command4_Click()
    Dim m As Integer, n As Integer
    m = List2.ListCount
    Rem ����ѭ�����б��List2���б���ȫ���ƶ���List1
    For n = 0 To m - 1
        List1.AddItem (List2.List(0))   '�ƶ��б�򶥶��б���
        List2.RemoveItem (0)
    Next n
    List1.Selected(0) = True            '���ñ�ѡ���б���
End Sub

Private Sub Form_Load()
    Rem ��ʼ���б��List1
    List1.AddItem ("����")
    List1.AddItem ("����")
    List1.AddItem ("Ӣ��")
    List1.AddItem ("���������")
    List1.AddItem ("���������")
    List1.AddItem ("ͼ��ͼ��")
    List1.AddItem ("��ý��")
    List1.AddItem ("���ӻ���")
    List1.AddItem ("C�������")
    List1.AddItem ("C++�������")
    List1.AddItem ("VB�������")
    List1.AddItem ("���ݿ����")
    List1.AddItem ("���ݽṹ")
    List1.AddItem ("���ԭ��")
    List1.AddItem ("����")
    List1.AddItem ("��ѡ")
    List1.Selected(0) = True                '���ñ�ѡ����
End Sub


Private Sub List1_DblClick()                '���б����ĳ�б��˫��ʱ
    Dim n As Integer
    n = List1.ListIndex                     '��¼��ǰ�б�������ֵ
    If List1.ListCount > 0 And n >= 0 Then
        List2.AddItem (List1.Text)          '����ǰ�б�����ӵ�List2
        List1.RemoveItem (n)                '���б����ɾ����ǰ�б���
        If List1.ListCount > n Then List1.ListIndex = n    '���豻ѡ����б���
    End If
End Sub
Private Sub Command1_Click()                '����>����������ʱ
    Dim n As Integer
    n = List1.ListIndex                     '��¼��ǰ�б�������ֵ
    If List1.ListCount > 0 And n >= 0 Then
        List2.AddItem (List1.Text)          '����ǰ�б�����ӵ�List2
        List1.RemoveItem (n)                '���б����ɾ����ǰ�б���
        If List1.ListCount > n Then List1.ListIndex = n    '���豻ѡ����б���
    End If
End Sub
Private Sub List2_DblClick()                '���б����ĳ�б��˫��ʱ
    Dim n As Integer
    n = List2.ListIndex                     '��¼��ǰ�б�������ֵ
    If List1.ListCount > 0 And n >= 0 Then
        List1.AddItem (List2.Text)          '����ǰ�б�����ӵ�List1
        List2.RemoveItem (n)                '���б����ɾ����ǰ�б���
        If List2.ListCount > n Then List2.ListIndex = n    '���豻ѡ����б���
    End If
End Sub
Private Sub Command3_Click()                '����>����������ʱ
    Dim n As Integer
    n = List2.ListIndex                     '��¼��ǰ�б�������ֵ
    If List1.ListCount > 0 And n >= 0 Then
        List1.AddItem (List2.Text)          '����ǰ�б�����ӵ�List1
        List2.RemoveItem (n)                '���б����ɾ����ǰ�б���
        If List2.ListCount > n Then List2.ListIndex = n   '���豻ѡ����б���
    End If
End Sub
