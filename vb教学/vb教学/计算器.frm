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
   Begin VB.CommandButton Command9 
      Caption         =   "="
      Height          =   1935
      Left            =   3840
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CLS"
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command7 
      Caption         =   "/"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   3240
      Width           =   650
   End
   Begin VB.CommandButton Command6 
      Caption         =   "*"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   2520
      Width           =   650
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1800
      Width           =   650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1080
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MOD"
      Height          =   495
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
Dim a As Double
Dim b As Double
Dim c As Double
Dim xsd As Boolean

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 0
Else
Text1.Text = Text1.Text & "0"
End If
Case 1
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 1
Else
Text1.Text = Text1.Text & "1"
End If
Case 2
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 2
Else
Text1.Text = Text1.Text & "2"
End If
Case 3
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 3
Else
Text1.Text = Text1.Text + "3"
End If
Case 4
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 4
Else
Text1.Text = Text1.Text + "4"
End If
Case 5
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 5
Else
Text1.Text = Text1.Text + "5"
End If
Case 6
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 6
Else
Text1.Text = Text1.Text + "6"
End If
Case 7
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 7
Else
Text1.Text = Text1.Text + "7"
End If
Case 8
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 8
Else
Text1.Text = Text1.Text + "8"
End If
Case 9
If Text1.Text = "" Or Text1.Text = "0" Then
Text1.Text = 9
Else
Text1.Text = Text1.Text & "9"
End If
End Select
End Sub

Private Sub Command2_Click()
If xsd = False Then
  If Text1.Text = "" Then
   Text1.Text = "0."
   xsd = True
  Else
   Text1.Text = Text1.Text + "."
   xsd = True
  End If
End If
End Sub

Private Sub Command3_Click()
a = Val(Text1.Text)
Text1.Text = ""
ysf = 5
xsd = False
End Sub

Private Sub Command4_Click()
a = Val(Text1.Text)
Text1.Text = ""
ysf = 1
xsd = False
End Sub

Private Sub Command5_Click()
a = Val(Text1.Text)
Text1.Text = ""
ysf = 2
xsd = False
End Sub

Private Sub Command6_Click()
If Text1.Text <> "" Then
a = Val(Text1.Text)
Text1.Text = ""
Else
Text1.Text = ""
End If
ysf = 3
xsd = False
End Sub

Private Sub Command7_Click()
a = Val(Text1.Text)
Text1.Text = ""
ysf = 4
xsd = False
End Sub

Private Sub Command8_Click()
a = 0
b = 0
c = 0
Text1.Text = ""
xsd = False
End Sub

Private Sub Command9_Click()
b = Val(Text1.Text)
Select Case ysf
Case 1
c = a + b
Case 2
c = a - b
Case 3
c = a * b
Case 4
c = a / b
If Abs(c) < 1 Then
c = "0" + CStr(c)
End If
Case 5
c = a Mod b
End Select
Text1.Text = c
End Sub

