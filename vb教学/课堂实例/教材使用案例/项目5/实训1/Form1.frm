VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   3225
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'����Ͽ�������б���
With Combo1
.AddItem "����"
.AddItem "�Ϻ�"
.AddItem "�����人"
.AddItem "���ϳ�ɳ"
.AddItem "�Ĵ��ɶ�"
.AddItem "�㶫����"
.ListIndex = 0
End With
End Sub


Private Sub Combo1_Click()
'���ı����з�����ʾ��ѡ�е��б���
Text1.Text = Text1.Text + Combo1.Text + Chr(13) + Chr(10)
End Sub

