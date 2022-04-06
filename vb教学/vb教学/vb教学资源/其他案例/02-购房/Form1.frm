VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   8625
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "卧室2"
      Height          =   2055
      Index           =   3
      Left            =   3360
      TabIndex        =   55
      Top             =   7440
      Width           =   2775
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   57
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   56
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "面积"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "单价"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "售价"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   59
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "卧室1"
      Height          =   2055
      Index           =   2
      Left            =   3360
      TabIndex        =   48
      Top             =   5280
      Width           =   2775
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   49
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "面积"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "单价"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "售价"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "客厅"
      Height          =   2055
      Index           =   1
      Left            =   3360
      TabIndex        =   41
      Top             =   3120
      Width           =   2775
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   42
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "面积"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "单价"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "售价"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "卧室2"
      Height          =   2055
      Index           =   3
      Left            =   240
      TabIndex        =   34
      Top             =   7440
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "面积"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "单价"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "售价"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "卧室1"
      Height          =   2055
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   5280
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "面积"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "单价"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "售价"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "客厅"
      Height          =   2055
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "面积"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "单价"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "售价"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   6600
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   6600
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "厨房"
      Height          =   2055
      Index           =   0
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   2775
      Begin VB.TextBox Text6 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "售价"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "单价"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "面积"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "厨房"
      Height          =   2055
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2775
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "售价"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "单价"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "面积"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label10 
      Caption         =   "第二套"
      Height          =   615
      Left            =   3360
      TabIndex        =   63
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "第一套"
      Height          =   615
      Left            =   240
      TabIndex        =   62
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "选择"
      Height          =   180
      Left            =   6720
      TabIndex        =   17
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "总价"
      Height          =   180
      Left            =   6360
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(3) As Long, b(3) As Long, c(3) As Long, d(3) As Long, sum1 As Long
Dim e(3) As Long, f(3) As Long, g(3) As Long, h(3) As Long, sum2 As Long
Private Sub Command1_Click()
sum1 = sum2 = 0
For i = 0 To 3 Step 1
   a(i) = Val(Text1(i).Text)
   b(i) = Val(Text2(i).Text)
   c(i) = Val(Text3(i).Text)
   d(i) = a(i) * b(i) * c(i)
   sum1 = sum1 + d(i)
   e(i) = Val(Text4(i).Text)
   f(i) = Val(Text5(i).Text)
   g(i) = Val(Text6(i).Text)
   h(i) = e(i) * f(i) * g(i)
   sum2 = sum2 + h(i)
   Next i
    Text7.Text = Str$(sum1)
    Text8.Text = Str$(sum2)
If sum1 <= sum2 Then
 Text9.Text = "第一套"
    Else
 Text9.Text = "第二套"
 End If
 
End Sub

