VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   5925
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdpaste 
      Caption         =   "Õ³Ìù(&v)"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "¸´ÖÆ(&c)"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdcut 
      Caption         =   "¼ôÇÐ(&x)"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Height          =   2175
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcopy_Click()
Clipboard.Clear
Clipboard.SetText Txt1.SelText
End Sub

Private Sub cmdcut_Click()
Clipboard.Clear
Clipboard.SetText Txt1.SelText
Txt1.SelText = ""

End Sub

Private Sub cmdpaste_Click()
Txt1.SelText = Clipboard.GetText()
End Sub

