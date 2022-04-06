VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6825
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "产生随机数"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查询偶数"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(0 To 20) As Integer
Dim j, m, n As Integer



Private Sub Command1_Click()
Command2.Enabled = True
FontSize = 24 '设置文本显示字体
Print  '在窗体上大隐换行
Print "产生的20个随机数为："
For j = 1 To 20
 a(j) = CInt(Rnd * 100) '产生100以内的随机数
 Print a(j);  '在窗体上显示没个随机数
 If j Mod 5 = 0 Then  '设置内5个随机数为一行
  Print
 End If
Next j
End Sub



Private Sub Command2_Click()
n = 0
Print "其中偶数为："
For j = 1 To 20
m = Sushu(a(j)) '引用函数判断a(j)这个随机数是否为素数
If m = 0 Then  '在窗体上打印素数
Print a(j);
n = n + 1
End If
If n Mod 5 = 0 And n <> 0 Then  '设置每5个素数为一行
n = 0
Print
End If
Next j

End Sub
