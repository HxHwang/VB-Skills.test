VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ѧ�Ǽǳ���"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8400
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   4215
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "��ѧѧԺ��"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "��"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "��ѧ���£�"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "�Ա�"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stradd As String

Private Sub Command1_Click()
If Text1.Text <> "" Then
  stradd = Text1.Text & "   " & Combo1.Text & "   " & Combo2.Text & "��" & Combo3.Text & "��" & "   " & Combo4.Text
  List1.AddItem stradd
End If
End Sub

Private Sub Command2_Click()
If List1.ListIndex >= 0 Then
  List1.RemoveItem List1.ListIndex
Else
  Exit Sub
End If
End Sub

Private Sub Command3_Click()
List1.Clear
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Combo1.AddItem "��"
Combo1.AddItem "Ů"
Combo2.AddItem "2005"
Combo2.AddItem "2006"
Combo2.AddItem "2007"
Combo2.AddItem "2008"
Combo3.AddItem "1"
Combo3.AddItem "2"
Combo3.AddItem "3"
Combo3.AddItem "4"
Combo3.AddItem "5"
Combo3.AddItem "6"
Combo3.AddItem "7"
Combo3.AddItem "8"
Combo3.AddItem "9"
Combo3.AddItem "10"
Combo3.AddItem "11"
Combo3.AddItem "12"
Combo4.AddItem "�����ѧԺ"
Combo4.AddItem "����ѧԺ"
Combo4.AddItem "����ѧԺ"
Combo1.Text = Combo1.List(0)
Combo2.Text = Combo2.List(0)
Combo3.Text = Combo3.List(0)
Combo4.Text = Combo4.List(0)
End Sub
