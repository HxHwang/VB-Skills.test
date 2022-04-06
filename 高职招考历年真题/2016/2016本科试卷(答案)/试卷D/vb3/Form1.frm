VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5595
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "ÔËËã"
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "Õû³ý"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "³Ë"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¼õ"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¼Ó"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "½á¹û£º"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ÇëÊäÈëÁ½¸öÕýÕûÊý£º"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click(Index As Integer)

Dim a, b, c As Single
a = Val(Text1.Text)
b = Val(Text2.Text)

Select Case Index
Case 0
Label3.Caption = "+"
c = a + b

Case 1
c = a - b
Label3.Caption = "-"
Case 2
c = a * b
Label3.Caption = "x"
Case 3
c = a \ b
Label3.Caption = "\"
End Select
Label2.Caption = "½á¹û£º" & c

End Sub
