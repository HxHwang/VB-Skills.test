Attribute VB_Name = "Module1"
Option Explicit
Public dirty As Boolean
Dim Filename As String

Public Sub NewFile()
Dim s As Integer
'如果文本框中的内容发生改变，询问是否保存文件
If dirty = True Then
s = MsgBox("文件已改变,是否保存?", vbYesNoCancel + vbInformation, _
"保存")
Select Case s
'保存文件，然后新建文件
Case vbYes
SaveFile
GoTo ss
'不保存文件，直接新建文件
Case vbNo
GoTo ss
'直接返回主程序
Case vbCancel
Exit Sub
End Select
Else
GoTo ss
End If
'新建文件
ss:
Form1.txtText.text = ""
dirty = False
Form1.Caption = "Untitled"
End Sub

Public Sub OpenFile()
Dim s As Integer
Dim text As String
Dim textbuff As String
'如果文本框中的内容发生改变，询问是否保存文件
If dirty = True Then
s = MsgBox("文件已改变,是否保存?", vbYesNoCancel + vbInformation, _
"保存")
Select Case s
'先保存文件，然后打开文件
Case vbYes
SaveFile
GoTo ss
'直接打开文件
Case vbNo
GoTo ss
Case vbCancel
Exit Sub
End Select
Else
GoTo ss
End If
'打开文件
ss:
'设置文件过滤器
Form1.CommonDialog1.Filter = "文本文件(*.txt)|*.txt"
'显示"打开"对话框
Form1.CommonDialog1.ShowOpen
Filename = Form1.CommonDialog1.Filename
Form1.Caption = Filename
If Filename <> "" Then
Form1.Caption = Filename
'打开顺序文件
Open Filename For Input As #1
'读取顺序文件中的内容，并将它显示到文本框中
Do While Not EOF(1)
Line Input #1, text
textbuff = textbuff + text
Form1.txtText.text = textbuff
Loop
'关闭文件
Close #1
End If
End Sub

Public Sub SaveFile()
'设置文件过滤器
Form1.CommonDialog1.Filter = "文本文件(*.txt)|*.txt"
'显示"另存为"对话框
Filename = Form1.Caption
If Form1.Caption = "Untitled" Then
SaveAsFile
Else
'打开顺序文件
Open Filename For Output As #1
'将文本框中的内容写入文件
Print #1, Form1.txtText.text
'关闭文件
Close #1
End If
End Sub

Public Sub SaveAsFile()
Dim text As String
Dim textbuff As String
'设置文件过滤器
Form1.CommonDialog1.Filter = "文本文件(*.txt)|*.txt"
'显示"另存为"对话框
Form1.CommonDialog1.ShowSave
Filename = Form1.CommonDialog1.Filename
Form1.Caption = Filename
If Filename <> "" Then
'打开顺序文件
Open Filename For Output As #1
'将文本框中的内容写入文件
Print #1, Form1.txtText.text
'关闭文件
Close #1
End If
End Sub


