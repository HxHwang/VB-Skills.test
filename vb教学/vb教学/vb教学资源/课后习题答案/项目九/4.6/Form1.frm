VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4620
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdEllip 
      Caption         =   "椭圆"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdArc 
      Caption         =   "圆弧"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdCirc 
      Caption         =   "圆"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdRect 
      Caption         =   "矩形"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdLine 
      Caption         =   "直线"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox picDraw 
      BackColor       =   &H80000009&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2000
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLine_Click()
     Dim i As Integer
     Dim y As Long
     picDraw.Cls
     For i = 0 To 4
         '设置线的类型
         picDraw.DrawStyle = i
         y = (300 * i) + 300
         '如果线型为实线，则设置线宽为5个像素
         If picDraw.DrawStyle = 0 Then
             picDraw.DrawWidth = 5
         '如果不是实线，则线宽只能为1
         Else
             picDraw.DrawWidth = 1
         End If
         '设置画图模式
         picDraw.DrawMode = i + 3
         picDraw.Line (850, y)-(3100, y), vbGreen
     Next
End Sub

Private Sub cmdClear_Click()
     picDraw.Cls
End Sub

Private Sub cmdQuit_Click()
     Unload Form1
End Sub

Private Sub cmdRect_Click()
     Dim R, i, x, y As Integer
     x = 100
     y = 1000
     picDraw.Cls
     picDraw.DrawMode = 13
     For i = 0 To 7
         '设置填充样式
         picDraw.FillStyle = i
         '设置填充颜色
         picDraw.FillColor = vbRed
         '绘制矩形
         picDraw.Line (x, y - 350)-(x + 200, y + 350), , B
         x = x + 500
     Next i
End Sub

Private Sub cmdCirc_Click()
     Dim R, i, x, y As Integer
     x = 500
     y = 1000
     picDraw.Cls
     picDraw.DrawMode = 13
     For i = 0 To 6
         '设置填充样式
         picDraw.FillStyle = i
         '设置填充颜色
         picDraw.FillColor = vbRed
         picDraw.Circle (x, y), 200
         x = x + 500
     Next i
End Sub

Private Sub cmdArc_Click()
     pi = 3.14
     picDraw.DrawMode = 13
     picDraw.Cls
     picDraw.Circle (350, 1000), 400, , pi / 4, 3 * pi / 4
     picDraw.Circle (1550, 1000), 400, , -pi / 4, 3 * pi / 4
     picDraw.Circle (2550, 1000), 400, , pi / 4, -3 * pi / 4
     picDraw.Circle (3550, 1000), 400, , -pi / 4, -3 * pi / 4
End Sub

Private Sub cmdEllip_Click()
Dim x, y, i As Integer
Dim j As Double
x = 500
y = 1000
j = 0
picDraw.Cls
picDraw.DrawMode = 13
For i = 1 To 3
'设置填充样式
picDraw.FillStyle = i
'设置填充颜色
picDraw.FillColor = vbRed
j = (1 - j) * 0.5
picDraw.Circle (x, y), 300, , , , j
x = x + 800
Next i
For i = 4 To 5
'设置填充样式
picDraw.FillStyle = i
'设置填充颜色
picDraw.FillColor = vbRed
picDraw.Circle (x, y), 400, , , , i - 2
x = x + 800
Next i
End Sub

