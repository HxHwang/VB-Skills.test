Attribute VB_Name = "Module1"

Public Function fun(n As Integer)
s = 0
For i = 1 To n
   s = s + i
Next i
fun = s
End Function
