VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ѭ���ṹ"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   6285
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Private Sub Command1_Click()
Print
For i = 1 To 9
  For j = 1 To i
    Print Tab(3 * j); i;
  Next j
  Print
Next i
For i = 8 To 1 Step -1
  For j = 1 To i
    Print Tab(3 * j); i;
  Next j
Next i
Command1.Caption = "���"
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
main.Show
Unload Me
End Sub
