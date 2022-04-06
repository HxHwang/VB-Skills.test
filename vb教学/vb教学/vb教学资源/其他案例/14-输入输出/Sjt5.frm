VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   2520
   ClientTop       =   1875
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   Begin VB.CommandButton C3 
      Caption         =   "存盘"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton C2 
      Caption         =   "计算"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton C1 
      Caption         =   "读入数据"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim Arr1(20) As Integer
Dim Arr2(20) As Integer
Dim Sum As Integer

Sub ReadData1()
    Open App.Path & "\" & "datain1.txt" For Input As #1
    For i = 1 To 20
        Input #1, Arr1(i)
    Next i
    Close #1
End Sub

Sub ReadData2()
    Open App.Path & "\" & "datain2.txt" For Input As #1
    For i = 1 To 20
        Input #1, Arr2(i)
    Next i
    Close #1
End Sub

Sub WriteData(Filename As String, Num As Integer)
    Open App.Path & "\" & Filename For Output As #1
    Print #1, Num
    Close #1
End Sub

Private Sub C1_Click()
ReadData1
    ReadData2

End Sub

Private Sub C2_Click()
Dim arr3(20) As Integer
    Sum = 0
    For i = 1 To 20
        arr3(i) = Arr1(i) \ Arr2(i)
        Sum = Sum + arr3(i)
    Next
    Print Sum

End Sub

Private Sub C3_Click()
WriteData "dataout.txt", Sum
End Sub
