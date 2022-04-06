VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8490
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(20) As Integer
Dim n As Integer
Dim k As Integer
Dim sushu As Boolean
Private Sub Command1_Click()
For i = 1 To 20
  a(i) = CInt(Rnd * 100)
  n = n + 1
  Print a(i);
  If n Mod 5 = 0 Then
     Print
  End If
Next i
End Sub

Private Sub Command2_Click()

Print "其中素数有："
For i = 1 To 20
 k = 2
 sushu = True
 Do While k < a(i)
    If a(i) Mod k = 0 Then
      sushu = False
      Exit Do
    Else
      k = k + 1
    End If
 Loop
 If sushu = True Then
    Print a(i);
 End If
Next i
End Sub

Private Sub Form_Load()
sushu = False
End Sub
