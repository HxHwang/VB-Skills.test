VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   7050
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "�Ƽ�����"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "ȷ������"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   1080
   End
   Begin VB.HScrollBar HScMove 
      Height          =   375
      LargeChange     =   10
      Left            =   4680
      Max             =   100
      SmallChange     =   10
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtMobile 
      Height          =   1695
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame Frame6 
      Caption         =   "���ӹ���"
      Height          =   2295
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
      Begin VB.CheckBox chkFunc4 
         Caption         =   "���ߴ��䣨������"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkFunc3 
         Caption         =   "������Ƶ��Mp4)"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkFunc2 
         Caption         =   "�����֣�Mp3��"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkFunc1 
         Caption         =   "����"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "�۸�"
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton optPrice4 
         Caption         =   "1000����"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton optPrice3 
         Caption         =   "1000��2000"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optPrice2 
         Caption         =   "2000��3000"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optPrice1 
         Caption         =   "3000����"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "����"
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
      Begin VB.ComboBox cboVideo 
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":000D
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��Ļɫ��"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
      Begin VB.ComboBox cboView 
         Height          =   300
         ItemData        =   "Form1.frx":0029
         Left            =   120
         List            =   "Form1.frx":0033
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "���"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ʒ��"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.ListBox lstLab 
         Height          =   420
         ItemData        =   "Form1.frx":0043
         Left            =   120
         List            =   "Form1.frx":0056
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public price As Integer
Public mlab As String
Public mprice As String
Public mfunc As String
Public mobile As String



Private Sub Form_Load()
    Label1.Visible = False
    txtMobile.Visible = False
    Timer1.Enabled = False
    With cboType
        .AddItem "ֱ��"
        .AddItem "����"
        .AddItem "����"
        .AddItem "��ת��"
    End With
    cboType.ListIndex = 0
    cboView.ListIndex = 0
    optPrice4.Value = True
    cboVideo.ListIndex = 0
    mfunc = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Form1
    Unload Form2
End Sub
Private Sub chkFunc1_Click()
    mfunc = mfunc + ", " + chkFunc1.Caption
End Sub

Private Sub chkFunc2_Click()
    mfunc = mfunc + ", " + chkFunc1.Caption
End Sub

Private Sub chkFunc3_Click()
    mfunc = mfunc + ", " + chkFunc1.Caption
End Sub

Private Sub chkFunc4_Click()
    mfunc = mfunc + ", " + chkFunc1.Caption
End Sub
Private Sub cmdChoose_Click()
    cmdOk_Click
    If mlab = "" Then
       MsgBox "��û��ѡ��Ʒ�ƣ�", vbOKOnly + vbInformation, "��ʾ"
    Else
     cmdChoose.Caption = "���Եȡ�����"
      Timer1.Enabled = True
      Timer1.Interval = 500
    End If
End Sub

Private Sub cmdOk_Click()
    txtMobile.Visible = True
    Label1.Visible = True
    cmdChoose.Enabled = True
    HScMove.Value = 0
    Timer1.Enabled = False
    txtMobile.Text = "Ʒ�ƣ�" + lstLab.Text + Chr(13) + Chr(10) + _
    "��Ļɫ��:" + cboView.Text + Chr(13) + Chr(10) + _
                   "��ۣ�" + cboType.Text + Chr(13) + Chr(10) + _
    "�۸�Χ:" + mprice + Chr(13) + Chr(10) + _
                   "������" + cboVideo.Text + Chr(13) + Chr(10) + _
     "���ӹ��ܣ�" + mfunc
    End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub lstLab_Click()
    mlab = lstLab.Text
End Sub

Private Sub optPrice1_Click()
    mprice = optPrice1.Caption
    price = 3000
    chkFunc2.Enabled = True
    chkFunc3.Enabled = True
    chkFunc4.Enabled = True
End Sub

Private Sub optPrice2_Click()
    mprice = optPrice2.Caption
    price = 2000
    chkFunc2.Enabled = True
    chkFunc3.Enabled = True
    chkFunc4.Enabled = True
End Sub

Private Sub optPrice3_Click()
    mprice = optPrice3.Caption
    price = 1000
    chkFunc2.Enabled = True
    chkFunc3.Enabled = True
    chkFunc4.Enabled = True
End Sub

Private Sub optPrice4_Click()
    mprice = optPrice4.Caption
    price = 500
    chkFunc2.Enabled = False
    chkFunc3.Enabled = False
    chkFunc4.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If HScMove.Value = HScMove.Max Then
        mobile = ChooseMobile(price, mlab)
        Select Case mobile
            Case "��"
                MsgBox "�Բ���û��������Ļ��ͣ�", vbOKOnly + vbCritical, "ѡ������"
            Case "������"
                MsgBox "�Բ��������λ��û�н�����", vbOKOnly + vbCritical, "ѡ������"
            Case Else
                MsgBox "�Ƽ�����Ϊ��" + mobile, vbOKOnly + vbInformation, "�Ƽ�����"
        End Select
        Timer1.Interval = 0
        cmdChoose.Caption = "�Ƽ�����"
        HScMove.Value = 0
    Else
        HScMove.Value = HScMove.Value + HScMove.LargeChange
    End If
End Sub

