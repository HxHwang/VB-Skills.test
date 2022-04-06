VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "产生随机数并查询奇偶数"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   10725
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查询奇/偶数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生随机数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   0
      Top             =   4920
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim n As Integer
Dim a(20) As Integer
Private Sub Command1_Click()
Print
Print "生成20个0~100间的随机整数："
For i = 1 To 20
  a(i) = CInt(Rnd * 100)
  n = n + 1
  Print a(i) & "    ";
  If n Mod 5 = 0 Then
    Print
  End If
Next i
End Sub

Private Sub Command2_Click()

n = 0
Print
Print "其中奇数有："
For i = 1 To 20
  If a(i) Mod 2 <> 0 Then
    Print a(i) & "    ";
    n = n + 1
    If n Mod 5 = 0 Then
      Print
    End If
  End If
Next i
n = 0
Print
Print "其中偶数有："
For i = 1 To 20
  If a(i) Mod 2 = 0 Then
    Print a(i) & "    ";
    n = n + 1
    If n Mod 5 = 0 Then
      Print
    End If
  End If
Next i
End Sub

Private Sub Command3_Click()
main.Show
Unload Me
End Sub

Private Sub Form_Load()
Randomize
End Sub
