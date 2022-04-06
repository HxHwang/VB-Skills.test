VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5565
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Í£Ö¹"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¿ªÊ¼"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "¿¼ÊÔ³É¹¦£¡"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 10
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Label1.Left < Form1.Width - Label1.Width Then
Label1.Left = Label1.Left + 10
Else
Label1.Left = 120
End If




End Sub
