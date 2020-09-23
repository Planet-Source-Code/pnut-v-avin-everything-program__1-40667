Attribute VB_Name = "mdlScreen"
Public Function GetDec(Hx As String) As Integer
    On Error Resume Next
    If Asc(Left$(Hx, 1)) >= 65 Then
        Groups = Asc(Left$(Hx, 1)) - 55
    Else
        Groups = Val(Left$(Hx, 1))
    End If
        Groups = Groups * 16
    If Asc(Right$(Hx, 1)) >= 65 Then
        LeftOver = Asc(Right$(Hx, 1)) - 55
    Else
        LeftOver = Val(Right$(Hx, 1))
    End If
    GetDec = Groups + LeftOver
End Function
