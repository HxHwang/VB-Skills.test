VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6210
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdSmall 
      Caption         =   "��С"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CmdLarge 
      Caption         =   "�Ŵ�"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "��ʾͼƬ"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   240
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const small As Single = 0.5
Private Const large As Single = -1

'�Ŵ���С�������
Private Sub Zoom(ByVal img As Image, ByVal ratio As Single)
    'ͨ���ı�ͼƬ��ĳߴ��λ����ʵ�ֶ�ͼƬ�ķŴ����С
    Image1.Left = Image1.Left + img.Width * ratio / 2
    Image1.Top = Image1.Top + Image1.Height * ratio / 2
    Image1.Width = Image1.Width - Image1.Width * ratio
    Image1.Height = Image1.Height - Image1.Height * ratio
End Sub

Private Sub CmdLarge_Click()
    Zoom Image1, large
End Sub

Private Sub CmdShow_Click()
    CommonDialog1.Filter = "ͼ���ļ�(*.bmp;*.jpg)|*.bmp;*.jpg|�����ļ�|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DialogTitle = "��ͼƬ�ļ�"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub CmdSmall_Click()
    Zoom Image1, small
End Sub
