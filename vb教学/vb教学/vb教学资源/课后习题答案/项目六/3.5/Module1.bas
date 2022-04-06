Attribute VB_Name = "Module1"

Public Function fun(a As Integer)
k = Int(Sqr(a))
swit = 0
i = 2
Do While i <= k And swit = 0
If a Mod i = 0 Then
swit = 1
Else
i = i + 1
End If
Loop
fun = swit
End Function
