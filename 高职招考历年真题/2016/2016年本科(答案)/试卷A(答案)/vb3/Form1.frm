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
      Caption         =   "�ж�����"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "������һ������2����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
'y = 0 '��������
For i = 2 To n - 1
If n Mod i = 0 Then
Label2.Caption = n & "������"
'Print n & "������"
Exit For
'y = 1 '������
Else
Label2.Caption = n & "��������"
'Print n & "��������"
End If
Next i
'If y = 0 Then Label2.Caption = n & "��������"
End Sub
