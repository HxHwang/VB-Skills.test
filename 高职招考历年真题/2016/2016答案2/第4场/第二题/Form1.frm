VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "最大公约数"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As Integer, y As Integer
Dim i As Integer, max As Integer
x = Val(Text1)
y = Val(Text2)
For i = 1 To x
    If x Mod i = 0 And y Mod i = 0 Then
        'Print "公约数：" & i
        If i > max Then
            max = i
        End If
    End If
Next i
Label1 = "最大公约数：" & max
End Sub
'这道题 需要理解,最大公约数：一个数能被一个数整除
'例如 4 可以被1、2、4整除  2 可以被 1、2整除
'    那么2和4的公约数就是 1、2  最大的就是2 这个就是最大公约数
'可以先求一个数能被哪些数整除

Private Sub Form_Load()
Text1 = ""
Text2 = ""
Label1 = ""
End Sub
