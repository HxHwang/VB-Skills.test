VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   450
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
      Pattern         =   "*.txt"
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
   
End Sub


