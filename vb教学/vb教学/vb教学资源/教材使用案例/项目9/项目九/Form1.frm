VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "一个简易的画图程序"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6435
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdEnd 
      Caption         =   "退出"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "清除画布"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存文件"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox ComboSize 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Text            =   "2"
      Top             =   3480
      Width           =   1410
   End
   Begin VB.CommandButton CmdColor 
      Caption         =   "画笔颜色"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1530
   End
   Begin VB.PictureBox PicDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      Height          =   2625
      Left            =   0
      ScaleHeight     =   2565
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   960
      Top             =   4080
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   600
      Y1              =   4080
      Y2              =   4440
   End
   Begin VB.Label LblShape 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label LblShape 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label LblShape 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label LblShape 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image ImgShow 
      Height          =   1455
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "设置画笔的尺寸:"
      Height          =   180
      Left            =   1920
      TabIndex        =   3
      Top             =   3240
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldx, oldy, shape As Integer

Private Sub ComboSize_Change()
  PicDraw.DrawWidth = Int(ComboSize.Text)
End Sub

Private Sub ComboSize_Click()
  PicDraw.DrawWidth = Int(ComboSize.Text)
End Sub

Private Sub Form_Load()
  Dim i As Integer
  shape = 1
  Do While i <= 40
     i = i + 2
     ComboSize.AddItem Str(i)
  Loop
End Sub

Private Sub LblShape_Click(Index As Integer)
  Select Case Index
    Case 0
      shape = 0
    Case 1
      shape = 1
    Case 2
      shape = 2
    Case 3
      shape = 3
  End Select
End Sub

Private Sub PicDraw_Change()
    ImgShow.Picture = PicDraw.Image
End Sub

Private Sub PicDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oldx = X
  oldy = Y
End Sub

Private Sub PicDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If shape = 0 Then PicDraw.Line (oldx, oldy)-(X, Y)
  If shape = 1 Then
'     PicDraw.Line (oldx, oldy)-(oldx, Y)
'     PicDraw.Line (oldx, oldy)-(X, oldy)
'     PicDraw.Line (oldx, Y)-(X, Y)
'     PicDraw.Line (X, oldy)-(X, Y)
    PicDraw.Line (oldx, oldy)-(X, Y), , B
  End If
  If shape = 2 Then
     If Abs(X - oldx) > Abs(Y - oldy) Then radius = Abs(Y - oldy) Else radius = Abs(X - oldx)
     PicDraw.Circle (oldx, oldy), radius
  End If
  If shape = 3 Then
     If Abs(X - oldx) > Abs(Y - oldy) Then radius = Abs(Y - oldy) Else radius = Abs(X - oldx)
     PicDraw.Circle (oldx, oldy), radius, , , , 0.5
  End If
  ImgShow.Picture = PicDraw.Image
  
End Sub

Private Sub CmdColor_Click()
  CommonDialog1.ShowColor
  PicDraw.ForeColor = CommonDialog1.Color
End Sub

Private Sub CmdOpen_Click()
  CommonDialog1.Filter = "BMP文件|*.bmp"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
     PicDraw.Picture = LoadPicture(CommonDialog1.FileName)
  End If
End Sub

Private Sub CmdSave_Click()
  CommonDialog1.Filter = "BMP文件|*.bmp"
  CommonDialog1.ShowSave
  If CommonDialog1.FileName <> "" Then
     SavePicture PicDraw.Image, CommonDialog1.FileName
  End If
End Sub

Private Sub CmdClear_Click()
  PicDraw.Cls
End Sub

Private Sub CmdEnd_Click()
  End
End Sub

