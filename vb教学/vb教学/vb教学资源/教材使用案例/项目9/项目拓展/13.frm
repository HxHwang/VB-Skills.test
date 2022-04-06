VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4560
      Top             =   480
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   8
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   0
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   7
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   6
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   5
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   4
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   3
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   2
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   1
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 8
  Shape(i).Visible = False
Next i

End Sub

Private Sub Timer1_Timer()
Static j As Integer
  Shape(j).Visible = False
  j = (j + 1) Mod 9
  Shape(j).Visible = True
  
  
End Sub
