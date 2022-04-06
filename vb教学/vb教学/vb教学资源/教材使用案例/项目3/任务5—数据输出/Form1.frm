VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "数据输出"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "输出"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '定义变量
    Dim i As Integer, j As Integer      '定义整形变量
    Dim S As String                     '定义变长字符串
    Dim S1 As String * 10, S2 As String * 5, S3 As String * 1       '定义定长字符串
    
    i = 2:    j = -5
    Print "输出数值数据:"   '输出字符串常量
    Print "i="; i           '输出数值数据
    Print "j="; j
    Print "i+j="; i + j     '输出计算表达式的值
    Print                   '输出一个空行
    
    S = "abcde"
    Print "使用分号输出变长字符串数据："
    Print S; "ABCDE"        '使用分号输出变长字符串变量和字符串常量
    Print
    Print "使用逗号输出变长字符串数据："
    Print S, "ABCDE"         '使用逗号输出变长字符串变量和字符串常量
    Print
    
    S1 = "xyz": S2 = "xyz": S3 = "xyz"
    Print "使用分号输出定长字符串数据S1,S2,S3："
    Print S1; S2;           '尾部加分号表示下一个变量输出不换行
    Print S3
End Sub


