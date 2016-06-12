Attribute VB_Name = "ValidHora"
Function hora(Hs As Integer) As Boolean
Dim Mm As Integer

Mm = Hs Mod 100

If Hs > -1 And Hs < 2400 And Mm > -1 And Mm < 60 Then
    hora = True
Else
    hora = False
End If

Exit Function

End Function
