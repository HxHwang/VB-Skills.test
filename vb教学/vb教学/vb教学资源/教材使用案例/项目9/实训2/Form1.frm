VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   3030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdLarge 
      Caption         =   "放大"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSmall 
      Caption         =   "缩小"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgCat 
      Height          =   3315
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const small As Single = 0.5
Private Const large As Single = -1

Private Sub cmdLarge_Click()
    Zoom imgCat, large
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSmall_Click()
    Zoom imgCat, small
End Sub

'放大、缩小处理过程
Private Sub Zoom(ByVal img As Image, ByVal ratio As Single)
    '通过改变图片框的尺寸和位置来实现对图片的放大和缩小
    imgCat.Left = imgCat.Left + img.Width * ratio / 2
    imgCat.Top = imgCat.Top + imgCat.Height * ratio / 2
    imgCat.Width = imgCat.Width - imgCat.Width * ratio
    imgCat.Height = imgCat.Height - imgCat.Height * ratio
End Sub

