VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "计算器"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCal 
      Caption         =   "+"
      Height          =   735
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE"
      Height          =   735
      Left            =   2880
      TabIndex        =   18
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdReCal 
      Caption         =   "C"
      Height          =   735
      Left            =   2160
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   735
      Left            =   2880
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "="
      Height          =   735
      Index           =   4
      Left            =   1440
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "."
      Height          =   735
      Index           =   10
      Left            =   720
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "2"
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Appearance      =   0  'Flat
      Caption         =   "1"
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "3"
      Height          =   735
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "4"
      Height          =   735
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "*"
      Height          =   735
      Index           =   3
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "/"
      Height          =   735
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "0"
      Height          =   735
      Index           =   9
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "－"
      Height          =   735
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "9"
      Height          =   735
      Index           =   8
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "8"
      Height          =   735
      Index           =   7
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "7"
      Height          =   735
      Index           =   6
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "6"
      Height          =   735
      Index           =   5
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdNumb 
      Caption         =   "5"
      Height          =   735
      Index           =   4
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prenum As Integer '用于保存第一个数
Dim lastnum As Integer '用于保存第二个数
Dim cal As Integer  '用于保存所按的运算符号
Dim cflag As Integer '用于判断是否按下运算符号，当cflag＝0时，按下的数字键为一个数
                    '当cflag＝1时，按下的数字键为第二个数

Private Sub cmdCal_Click(Index As Integer)
On Error GoTo errorpro
    '按下等号后，才计算结果
    If Index = 4 Then
    '计算最后结果
        Select Case cal
            Case 0
                lblResult.Caption = Str(prenum + lastnum) '数字转为字符
            Case 1
                lblResult.Caption = Str(prenum + lastnum)
             Case 2
               lblResult.Caption = Str(prenum / lastnum)
             Case 3
                lblResult.Caption = Str(prenum * lastnum)
        End Select
    End If
'将计算出来的结果作为新的第一个数，用于连续计算
    prenum = Val(lblResult.Caption)
    '保存所按的运算符号
    cal = Index
    cflag = 1 '表示已按下运算符号，如果再按数字键，输入的为第二个数
errorpro:
    If Err.Number = 11 Then
        MsgBox "除数不能为零！", vbOKOnly + vbCritical, "错误"
        cflag = 0
        Exit Sub
    End If
   End Sub

Private Sub cmdClear_Click()
'清除所有设置
prenum = 0
lastnum = 0
cflag = 0
lblResult.Caption = "0."
End Sub


Private Sub cmdExit_Click()
Unload Form1 '退出程序
End Sub

Private Sub cmdNumb_Click(Index As Integer)
lblResult.Caption = cmdNumb(Index).Caption
'判断输入的是第一个数还是第二个数，在没按下运算符号之前，所输入
'的数字为第一个数，否则为第二个数
If cflag = 0 Then
    prenum = Val(lblResult.Caption) '将字符转化为数字
Else
    lastnum = Val(lblResult.Caption)
End If
End Sub

Private Sub cmdReCal_Click()
lblResult.Caption = "0." '清除错误输入的数字
End Sub

Private Sub Form_Load()
'初始化所有设置
prenum = 0
lastnum = 0
cflag = 0
End Sub
