VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   8415
   ClientTop       =   3855
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "判断质数"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个大于2的正整数："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim i, n, y As Integer
n = Val(Text1.Text)
'y = 0 '不是质数
For i = 2 To n - 1
If n Mod i = 0 Then
Label2.Caption = n & "是质数"
'Print n & "是质数"
Exit For
'y = 1 '是质数
Else
Label2.Caption = n & "不是质数"
'Print n & "不是质数"
End If
Next i
'If y = 0 Then Label2.Caption = n & "不是质数"
End Sub
