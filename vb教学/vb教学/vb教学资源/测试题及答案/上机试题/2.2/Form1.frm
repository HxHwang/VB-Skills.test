VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   7710
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton C1 
      Caption         =   "��ʾ"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox Ch3 
      Caption         =   "����"
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CheckBox Ch1 
      Caption         =   "����"
      Height          =   735
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Ch2 
      Caption         =   "����"
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C1_Click()
Print
Print "�ҵİ�����";
If Ch1.Value Then Print "����";
If Ch2.Value Then Print "����";
If Ch3.Value Then Print "����";

End Sub
