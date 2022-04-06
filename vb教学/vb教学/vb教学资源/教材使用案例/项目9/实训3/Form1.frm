VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   21.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4515
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Left            =   3720
      Top             =   2640
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      X1              =   2040
      X2              =   2040
      Y1              =   720
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   840
      Shape           =   3  'Circle
      Top             =   360
      Width           =   2655
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2040
      Y1              =   720
      Y2              =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'定义用于记录时针和分针长度的变量
Dim hLine, mLine As Integer
'定义用于记录时针和分针每次所转的角度
Dim i, j As Integer
Const pi = 3.14159


Private Sub Form_Load()
    '初始化，让时针、分针停在12点
    Line1.X1 = Shape1.Left + Shape1.Width / 2
    Line1.X2 = Line1.X2
    Line1.Y1 = Shape1.Top + Shape1.Height / 2
    Line1.Y2 = Line1.Y1 - Shape1.Height / 2 + 400
    Line2.X1 = Shape1.Left + Shape1.Width / 2
    Line2.X2 = Line2.X2
    Line2.Y1 = Shape1.Top + Shape1.Height / 2
    Line2.Y2 = Line2.Y1 - Shape1.Height / 2 + 150
    mLine = Line2.Y2 - Line2.Y1
    hLine = Line1.Y2 - Line1.Y1
    i = 0
    j = 0
    '分针开始走
    Timer1.Interval = 100
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X _
As Single, Y As Single)
    If Button = 1 Then
      Unload Form1
    End If
End Sub


Private Sub Timer1_Timer()
    '分针每隔100毫秒转10度
    Timer2.Interval = 0
    Line2.X2 = Line2.X1 + mLine * Cos((i + 90) * pi / 180)
    Line2.Y2 = Line2.Y1 + mLine * Sin((i + 90) * pi / 180)
    i = i + 10
    '每转3600，重新开始转，并且时针开始走
    If i = 360 Then
       i = 0
       Timer2.Interval = 1
    End If
End Sub


Private Sub Timer2_Timer()
    Line1.X2 = Line1.X1 + hLine * Cos((j + 90) * pi / 180)
    Line1.Y2 = Line1.Y1 + hLine * Sin((j + 90) * pi / 180)
    j = j + 5
    If j = 360 Then
       j = 0
End If
End Sub


