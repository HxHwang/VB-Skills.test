VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   2715
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdResize 
      Caption         =   "�� С"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdLarge 
      Caption         =   "�� ��"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdTurn 
      Caption         =   "�� ת"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdMove 
      Caption         =   "�� ��"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox PicCat 
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Integer

Private Sub cmdMove_Click()
     '������ťһ�Σ���ͼƬ������X����Y����ֱ��ƶ�80����λ
     j = j + 80
     '�ƶ�ͼƬ
     PicCat.PaintPicture PicCat.Picture, 0 + j, 0 + j, PicCat.Width, _
     PicCat.Height
End Sub

Private Sub cmdTurn_Click()
     '��ͼƬ��ԭ�㿪ʼ�ƶ�
     j = 0
     '��תͼƬ
     '���ݵ�����ť�Ĵ�������תͼƬ��ÿ����һ��ͼƬ��תһ��
     If i Mod 2 = 0 Then
     '�����Ĵ���Ϊż����ͼƬ��ת����
         PicCat.PaintPicture PicCat.Picture, PicCat.Width, _
         PicCat.Height, -PicCat.Width, -PicCat.Height
     Else
     '������ť����Ϊ������ͼƬ��ԭ
         PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width, _
         PicCat.Height
     End If
     '������ťһ�Σ���ť�������Ĵ�������һ��
     i = i + 1
End Sub

Private Sub cmdLarge_Click()
     '��ͼƬ��ԭ�㿪ʼ�ƶ�
     j = 0
     '������ťһ�Σ�ͼƬ�Ŀ�Ⱥͳ��ȶ�������80����λ
     i = i + 80
     '�ֲ��Ŵ�ͼƬ
     PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width + i, _
     PicCat.Height + i
End Sub
Private Sub cmdResize_Click()
     '��ͼƬ��ԭ�㿪ʼ�ƶ�
     j = 0
     '������ťһ�Σ�ͼƬ�Ŀ�Ⱥͳ��ȶ�����С80����λ
     i = i - 80
     '�ֲ���СͼƬ
     PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width + i, _
      PicCat.Height + i
End Sub


