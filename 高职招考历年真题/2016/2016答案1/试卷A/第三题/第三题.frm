VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   12060
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   2295
      Left            =   6360
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
      Begin VB.CheckBox Check2 
         Caption         =   "б��"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�Ӵ�"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   2655
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "����"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text1.FontBold = Not (Text1.FontBold)
End Sub

Private Sub Check2_Click()
Text1.FontItalic = Not (Text1.FontItalic)
End Sub

Private Sub Option1_Click()
Text1.FontName = "����"
End Sub

Private Sub Option2_Click()
Text1.FontName = "����"
End Sub
