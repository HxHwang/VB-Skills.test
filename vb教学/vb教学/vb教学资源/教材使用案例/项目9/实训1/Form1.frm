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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdResize 
      Caption         =   "缩 小"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdLarge 
      Caption         =   "放 大"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdTurn 
      Caption         =   "翻 转"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdMove 
      Caption         =   "移 动"
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
     '单击按钮一次，将图片顶点在X方向、Y方向分别移动80个单位
     j = j + 80
     '移动图片
     PicCat.PaintPicture PicCat.Picture, 0 + j, 0 + j, PicCat.Width, _
     PicCat.Height
End Sub

Private Sub cmdTurn_Click()
     '让图片从原点开始移动
     j = 0
     '翻转图片
     '根据单击按钮的次数来翻转图片，每单击一次图片翻转一次
     If i Mod 2 = 0 Then
     '单击的次数为偶数，图片倒转过来
         PicCat.PaintPicture PicCat.Picture, PicCat.Width, _
         PicCat.Height, -PicCat.Width, -PicCat.Height
     Else
     '单击按钮次数为奇数，图片还原
         PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width, _
         PicCat.Height
     End If
     '单击按钮一次，按钮被单击的次数增加一次
     i = i + 1
End Sub

Private Sub cmdLarge_Click()
     '让图片从原点开始移动
     j = 0
     '单击按钮一次，图片的宽度和长度都被拉伸80个单位
     i = i + 80
     '局部放大图片
     PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width + i, _
     PicCat.Height + i
End Sub
Private Sub cmdResize_Click()
     '让图片从原点开始移动
     j = 0
     '单击按钮一次，图片的宽度和长度都被缩小80个单位
     i = i - 80
     '局部缩小图片
     PicCat.PaintPicture PicCat.Picture, 0, 0, PicCat.Width + i, _
      PicCat.Height + i
End Sub


