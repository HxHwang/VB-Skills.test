VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结果"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer, m As Long, i As Integer, s As Double


Private Sub Command1_Click()
m = 1
n = Val(Text1.Text)
For i = 2 To n
 m = m * i
Next i
Text2.Text = Str$(m)

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""


End Sub

