<<<<<                           1
Private Sub btnIngresar_Click()

    Dim filaVacia As Long
    Dim i As Long
    Dim hayEspacio As Boolean

    hayEspacio = false

    For i = 9 to 33

    If Hoja1.Cells(i, 6).Value = "" Then

    filaVacia = i
    hayEspacio = true
    Exit For

    End If
    Next i

End Sub
