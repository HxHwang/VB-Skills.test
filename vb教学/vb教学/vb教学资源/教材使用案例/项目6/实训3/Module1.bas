Attribute VB_Name = "Module1"

Public Function Sushu(ByVal x As Integer)
Dim k, i As Integer
k = Int(Sqr(x))
i = 2
Sushu = 0
Do While i <= k And Sushu = 0
If x Mod i = 0 Then
Sushu = 1
Else
i = i + 1
End If
Loop
End Function
