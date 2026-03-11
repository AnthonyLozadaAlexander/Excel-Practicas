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

    If hayEspacio = True Then

      Hoja1.Cells(filaVacia, 6).Value = txtCantidad.Text
      Hoja1.Cells(filaVacia, 7).Value = txtDescripcion.Text
      Hoja1.Cells(filaVacia, 8).Value = txtValorUnitario.Text

    Else

     MsgBox "La Tabla Se Encuentra LLena"

    End If

End Sub
