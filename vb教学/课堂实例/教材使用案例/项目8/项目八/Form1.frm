VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5355
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdEnd 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   3330
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   3450
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "�༭"
      Begin VB.Menu MnuCut 
         Caption         =   "����"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "ճ��"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheck 
         Caption         =   "�鿴"
      End
      Begin VB.Menu MnuModi 
         Caption         =   "�޸�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sourfile As String '���ڱ���Դ�ļ�
Dim DestFile As String '���ڱ���Ŀ���ļ�
Dim SureCopy As Integer '���ڿ����Ƿ񵥻����ƻ���в˵�
Dim SureDell As Boolean '���ڿ����Ƿ񵥻�ɾ���ļ�
Dim sfn As String '���ڱ��汻ѡ���ļ����ļ���

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub Form_Load()
    Drive1.Drive = "F"
    SureCopy = 0
    SureDell = False
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    '����Ѿ�ѡ���ƻ��������򽫵�ǰ·����ΪĿ��·��
    If SureCopy = 1 Then
        If Right(Dir1.Path, 1) <> "\" Then
            DestFile = Dir1.Path + "\" + sfn
        Else
            DestFile = Dir1.Path + sfn
        End If
     '���û�У�����ǰ·����ΪԴ·��
    Else
        If Right(Dir1.Path, 1) <> "\" Then
            SourPath = Dir1.Path + "\"
        Else
            SourPath = Dir1.Path
        End If
    End If
End Sub

Private Sub File1_DblClick()

    '�õ��༭�ļ�����ϸ·��
    fName = Dir1.Path + "\" + File1.FileName
    '���ļ��༭�鿴����
    Form2.Show
    Form1.Hide

End Sub

Private Sub File1_Click()
    sfn = File1.FileName
    Form1.Caption = Dir1.Path + "\" + File1.FileName
    'ѡ��Դ�ļ�
    If Right(Dir1.Path, 1) <> "\" Then
        Sourfile = Dir1.Path + "\" + File1.FileName
    Else
        Sourfile = Dir1.Path + File1.FileName
    End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
    'ѡ���ļ����������"D"������ѯ���Ƿ�ɾ���ļ�
    If KeyAscii = 100 Then
        SureDel = MsgBox("ȷʵҪɾ���ļ���", vbYesNo + vbQuestion, "ɾ���ļ�")
        '�������"��"��ť����ɾ��ѡ���ļ����������"��"��ť����ɾ���ļ�
        Select Case SureDel
        Case vbYes
            'ɾ���ļ�
            Kill (Sourfile)
            '�����ļ��б�
            File1.Refresh
        Case vbNo
            Exit Sub
        End Select
    End If
End Sub

Private Sub MnuCheck_Click()
    '�õ��༭�ļ�����ϸ·��
    fName = Dir1.Path + "\" + File1.FileName
    '���ļ��༭�鿴����
    Form3.Show
    Form1.Hide
End Sub

Private Sub mnuCopy_Click()
    'ѡ����������Ժ�·������ΪĿ��·��
    If sfn = "" Then
        MsgBox "δѡ���ļ�", vbOKOnly + vbCritical, "����"
        SureCopy = 0
     Else
        SureCopy = 1
        SureDell = False
    End If
End Sub

Private Sub mnuCut_Click()
    'ѡ�����������Ժ�·������ΪĿ��·����ͬʱɾ����ѡ�е��ļ�
    If sfn = "" Then
        MsgBox "δѡ���ļ�", vbOKOnly + vbCritical, "����"
         SureCopy = 0
        SureDell = False
    Else
        SureCopy = 1
        SureDell = True
    End If

End Sub

Private Sub MnuModi_Click()
    '�õ��༭�ļ�����ϸ·��
    fName = Dir1.Path + "\" + File1.FileName
    '���ļ��༭�鿴����
    Form2.Show
    Form1.Hide
End Sub

Private Sub mnuPaste_Click()
    '����ļ����Ѿ����ڣ���ѯ���Ƿ񸲸��ļ�
    If SureCopy = 1 Then
       If Dir(DestFile) <> "" Then
            intfile = MsgBox("�ļ�" + DestFile + "�Ѿ����ڣ��Ƿ񸲸ǣ�", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "�����ļ�")
            Select Case intfile
            '�����ļ�
              Case vbYes
                FileCopy Sourfile, DestFile
              Case vbNo
                Exit Sub
              End Select
        Else
        '�����ļ�
         FileCopy Sourfile, DestFile
        End If
    End If
    '���ѡ����Ǽ��������ɾ��Դ�ļ�
    If SureDell = True Then
        Kill (Sourfile)
    End If
    File1.Refresh
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���ļ��б���е�������Ҽ����������༭���˵��Ŀ�ݲ˵�
    If Button = 2 Then
        PopupMenu MnuEdit
    End If
End Sub

