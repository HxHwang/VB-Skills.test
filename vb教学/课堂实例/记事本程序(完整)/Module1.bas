Attribute VB_Name = "Module1"
Option Explicit
Public dirty As Boolean
Dim Filename As String

Public Sub NewFile()
Dim s As Integer
'����ı����е����ݷ����ı䣬ѯ���Ƿ񱣴��ļ�
If dirty = True Then
s = MsgBox("�ļ��Ѹı�,�Ƿ񱣴�?", vbYesNoCancel + vbInformation, _
"����")
Select Case s
'�����ļ���Ȼ���½��ļ�
Case vbYes
SaveFile
GoTo ss
'�������ļ���ֱ���½��ļ�
Case vbNo
GoTo ss
'ֱ�ӷ���������
Case vbCancel
Exit Sub
End Select
Else
GoTo ss
End If
'�½��ļ�
ss:
Form1.txtText.text = ""
dirty = False
Form1.Caption = "Untitled"
End Sub

Public Sub OpenFile()
Dim s As Integer
Dim text As String
Dim textbuff As String
'����ı����е����ݷ����ı䣬ѯ���Ƿ񱣴��ļ�
If dirty = True Then
s = MsgBox("�ļ��Ѹı�,�Ƿ񱣴�?", vbYesNoCancel + vbInformation, _
"����")
Select Case s
'�ȱ����ļ���Ȼ����ļ�
Case vbYes
SaveFile
GoTo ss
'ֱ�Ӵ��ļ�
Case vbNo
GoTo ss
Case vbCancel
Exit Sub
End Select
Else
GoTo ss
End If
'���ļ�
ss:
'�����ļ�������
Form1.CommonDialog1.Filter = "�ı��ļ�(*.txt)|*.txt"
'��ʾ"��"�Ի���
Form1.CommonDialog1.ShowOpen
Filename = Form1.CommonDialog1.Filename
Form1.Caption = Filename
If Filename <> "" Then
Form1.Caption = Filename
'��˳���ļ�
Open Filename For Input As #1
'��ȡ˳���ļ��е����ݣ���������ʾ���ı�����
Do While Not EOF(1)
Line Input #1, text
textbuff = textbuff + text
Form1.txtText.text = textbuff
Loop
'�ر��ļ�
Close #1
End If
End Sub

Public Sub SaveFile()
'�����ļ�������
Form1.CommonDialog1.Filter = "�ı��ļ�(*.txt)|*.txt"
'��ʾ"���Ϊ"�Ի���
Filename = Form1.Caption
If Form1.Caption = "Untitled" Then
SaveAsFile
Else
'��˳���ļ�
Open Filename For Output As #1
'���ı����е�����д���ļ�
Print #1, Form1.txtText.text
'�ر��ļ�
Close #1
End If
End Sub

Public Sub SaveAsFile()
Dim text As String
Dim textbuff As String
'�����ļ�������
Form1.CommonDialog1.Filter = "�ı��ļ�(*.txt)|*.txt"
'��ʾ"���Ϊ"�Ի���
Form1.CommonDialog1.ShowSave
Filename = Form1.CommonDialog1.Filename
Form1.Caption = Filename
If Filename <> "" Then
'��˳���ļ�
Open Filename For Output As #1
'���ı����е�����д���ļ�
Print #1, Form1.txtText.text
'�ر��ļ�
Close #1
End If
End Sub


