VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7020
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   7020
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtText 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Menu MnuFile 
      Caption         =   "�ļ�"
      Begin VB.Menu MnuOpen 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "����"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub EditOpen()
    If fName <> "" Then
        '��˳���ļ�
        Open fName For Input As #1
        '��ȡ˳���ļ��е����ݣ���������ʾ���ı�����
        Do While Not EOF(1)
            Line Input #1, text
            textbuff = textbuff + text + Chr(13) + Chr(10)
            TxtText.text = textbuff
        Loop
        Close #1
    End If
End Sub

Private Sub Form_Load()
    Call EditOpen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub MnuOpen_Click()
    Dim text As String
    Dim textbuff As String
    '��ʾ"��"�Ի���
    CommonDialog1.ShowOpen
    fName = CommonDialog1.FileName
    Call EditOpen
End Sub

Private Sub MnuSave_Click()
    Dim fName As String
    Dim text As String
    Dim textbuff As String
    '��ʾ"���Ϊ"�Ի���
    CommonDialog1.ShowSave
    fName = CommonDialog1.FileName
    If fName <> "" Then
         '��˳���ļ�
         Open fName For Output As #1
         '���ı����е�����д���ļ�
         Print #1, TxtText.text
         '�ر��ļ�
         Close #1
    End If
End Sub

