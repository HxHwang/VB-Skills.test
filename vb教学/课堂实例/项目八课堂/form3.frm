VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5550
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8265
   LinkTopic       =   "Form2"
   ScaleHeight     =   5550
   ScaleWidth      =   8265
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
   Begin VB.Menu Mnufile 
      Caption         =   "�ļ�"
      Begin VB.Menu Mnuopen 
         Caption         =   "��"
      End
      Begin VB.Menu Mnusave 
         Caption         =   "����"
      End
      Begin VB.Menu Mnusaveas 
         Caption         =   "���Ϊ"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim xiugai As Boolean
Private Sub openfile()
Text1.Text = ""
If yfile = "" Then
   MsgBox "δѡ���ļ���", vbOKOnly + vbCritical, "����"
   Exit Sub
Else
   Open yfile For Input As #1
   Do While EOF(1) = False
     Line Input #1, s
     Text1.Text = Text1.Text & s & vbCrLf
   Loop
   Close #1
   Form2.Caption = yfile
End If
End Sub
Private Sub Form_Load()
Call openfile
xiugai = False
End Sub

Private Sub Mnuopen_Click()
If yfile <> "" And xiugai = True Then
  Dim answer As Integer
  answer = MsgBox("�ļ��Ѿ��޸ģ��Ƿ񱣴棿", vbYesNo + vbQuestion, "��ʾ")
  If answer = vbYes Then
    Open yfile For Output As #1
    Print #1, Text1.Text
    Close #1
    MsgBox "�ļ��ѱ��棡", vbOKOnly + vbInformation, "��Ϣ"
    xiugai = False
  Else
    Exit Sub
  End If
Else
  CommonDialog1.Filter = "�ı��ĵ�|*.txt|�����ļ�|*.*"
  CommonDialog1.ShowOpen
  yfile = CommonDialog1.FileName
  Call openfile
End If
End Sub


Private Sub Mnusave_Click()
Open yfile For Output As #1
Print #1, Text1.Text
Close #1
MsgBox "�ļ��ѱ��棡", vbOKOnly + vbInformation, "��Ϣ"
End Sub

Private Sub Mnusaveas_Click()
CommonDialog1.Filter = "�ı��ĵ�|*.txt|�����ļ�|*.*"
CommonDialog1.ShowSave
yfile = CommonDialog1.FileName
Open yfile For Output As #1
Print #1, Text1.Text
Close #1
MsgBox "�ļ��ѱ��棡", vbOKOnly + vbInformation, "��Ϣ"
End Sub

Private Sub Text1_Change()
xiugai = True
End Sub
