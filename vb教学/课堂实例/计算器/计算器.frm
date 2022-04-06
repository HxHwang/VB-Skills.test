VERSION 5.00
Begin VB.Form form1 
   Caption         =   "我的计算器"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4875
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Text            =   "0"
      Top             =   240
      Width           =   4020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "="
      Height          =   1935
      Left            =   3840
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CLS"
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   15
      Top             =   3240
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   14
      Top             =   2520
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   13
      Top             =   1800
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   12
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MOD"
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   2520
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   650
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ysf As Integer
Dim s1 As Double
Dim s2 As Double
Dim c As Double
Dim xsd As Boolean
Dim flag As Integer

Private Sub Command1_Click(Index As Integer)
If flag = 1 Then
  Text1.Text = ""
  flag = 0
End If
If Text1.Text = "" Or Text1.Text = "0" Then
   Text1.Text = CStr(Index)
Else
   Text1.Text = Text1.Text & CStr(Index)
End If

End Sub

Private Sub Command2_Click()
If flag = 1 Then
  Text1.Text = ""
  flag = 0
End If
If xsd = False Then
  If Text1.Text = "" Or Text1.Text = "0" Then
   Text1.Text = "0."
   xsd = True
  Else
   Text1.Text = Text1.Text + "."
   xsd = True
  End If
End If
End Sub



Private Sub Command3_Click(Index As Integer)
flag = 1
s1 = Val(Text1.Text)
xsd = False
ysf = Index
End Sub

Private Sub Command4_Click()
s1 = 0
s2 = 0
c = 0
Text1.Text = ""
xsd = False
flag = 0
End Sub

Private Sub Command5_Click()
s2 = Val(Text1.Text)
Select Case ysf
 Case 0
    c = s1 + s2
 Case 1
    c = s1 - s2
 Case 2
    c = s1 * s2
 Case 3
    c = s1 / s2
 Case 4
    c = s1 Mod s2
End Select
If Abs(c) < 1 Then
  Text1.Text = FormatNumber(c, Len(Str(c)) - 2, vbTrue)
Else
   Text1.Text = c
End If
End Sub

