VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   5925
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   3990
      Width           =   5925
      Begin VB.Image Image3 
         Height          =   1095
         Left            =   4080
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   1920
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   3120
      Picture         =   "form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   480
      Picture         =   "form1.frx":E1042
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Shape1.Visible = False
End Sub

Private Sub Image1_Click()
Shape1.Left = Image1.Left
Shape1.Top = Image1.Top
Shape1.Visible = True
Picture1.Cls
Picture1.Print "picture1"
Image3.Picture = Image1.Picture
Image3.Visible = True
End Sub


Private Sub Image2_Click()
Shape1.Left = Image2.Left
Shape1.Top = Image2.Top
Shape1.Visible = True
Picture1.Cls
Picture1.Print "picture2"
Image3.Picture = Image2.Picture
Image3.Visible = True
End Sub
